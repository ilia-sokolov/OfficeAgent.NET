using System.Security.Claims;
using System.Security.Cryptography;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.Abstractions;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Excel;
using OfficeAgent.Mcp;
using OfficeAgent.PowerPoint;
using OfficeAgent.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;

/// <summary>
/// Drives every tool into each response state a caller must handle, and checks that the fixture
/// really produced the state it is named after.
/// </summary>
/// <remarks>
/// One sample per tool freezes only the path that sample happened to take. A failure envelope,
/// a partial batch, an empty collection or a null receipt is a different shape, and a caller
/// written for it breaks just as badly when it changes. Each state here is its own baseline
/// section, so a field removed only from a failure envelope is a diff even though the success
/// shape is untouched. A state whose check fails stops generation: a label that no longer
/// describes its response would freeze the wrong thing.
/// </remarks>
internal static class ResponseStates
{
    public sealed record State(string Tool, string Name, string Json);

    public static async Task<IReadOnlyList<State>> Generate()
    {
        var folder = Path.Combine(Path.GetTempPath(), $"officeagent-states-{Guid.NewGuid():N}");
        Directory.CreateDirectory(folder);
        try
        {
            return await Generate(folder);
        }
        finally
        {
            Directory.Delete(folder, recursive: true);
        }
    }

    private static async Task<IReadOnlyList<State>> Generate(string folder)
    {
        var memory = new MemoryDocumentProvider("memory");
        var faulty = new FaultingProvider("faulty");
        var actor = new SwitchableActor();
        var services = new ServiceCollection();
        services.AddWordFormat();
        services.AddPowerPointFormat();
        services.AddExcelFormat();
        services.AddSingleton<IDocumentProvider>(memory);
        services.AddSingleton<IDocumentProvider>(faulty);
        services.AddFileSystemDocumentProvider("workspace", folder);
        services.AddSingleton<IAuditActorProvider>(actor);
        services.AddOfficeAgent();
        using var provider = services.BuildServiceProvider();
        var client = provider.GetRequiredService<OfficeAgentClient>();
        var tools = new OfficeAgentTools(client);
        var denied = new OfficeAgentTools(client, new DenyAll());
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();

        var states = new List<State>();
        async Task<string> Record(string tool, string name, Task<string> call, Func<JsonNode, bool> check, string expected)
        {
            var json = await call;
            var node = JsonNode.Parse(json)!;
            if (!check(node))
                throw new InvalidOperationException($"{tool} / {name}: expected {expected}, got {json}");
            states.Add(new State(tool, name, json));
            return json;
        }

        // Documents every state below starts from.
        var created = JsonNode.Parse(await tools.CreateDocument("memory", "contract.docx", CreatePlan("Direct")))!;
        var contract = Str(created, "outputDocumentId");
        var blank = Str(JsonNode.Parse(await tools.CreateDocument("memory", "blank.docx"))!, "outputDocumentId");
        var deck = Str(JsonNode.Parse(await tools.CreateDocument("memory", "deck.pptx"))!, "outputDocumentId");
        var book = Str(JsonNode.Parse(await tools.CreateDocument("memory", "book.xlsx"))!, "outputDocumentId");
        var template = Str(JsonNode.Parse(await tools.ImportDocumentContent("memory", "quote.docx",
            Convert.ToBase64String(QuoteTemplate())))!, "documentId");
        var editPlan = EditPlan(await tools.InspectDocument("memory", contract));
        var contractBytes = Convert.FromBase64String(Str(JsonNode.Parse(await tools.ExportDocumentContent("memory", contract))!, "contentBase64"));
        File.WriteAllBytes(Path.Combine(folder, "contract.docx"), contractBytes);
        const string badPlan = "{\"operations\":[{\"op\":\"changeText\",\"target\":{\"paraId\":\"00000000\",\"expect\":\"absent\"},\"with\":\"x\"}]}";
        const string emptyPlan = "{\"operations\":[]}";

        // describe_capabilities and list_connections
        await Record("describe_capabilities", "connections visible", tools.DescribeCapabilities(),
            n => Arr(n, "connections").Count > 0, "a non-empty connections list");
        await Record("describe_capabilities", "no connection visible to the caller", denied.DescribeCapabilities(),
            n => Arr(n, "connections").Count == 0, "an empty connections list");
        var mcpOptions = new OfficeAgentMcpOptions
        {
            AllowCreation = true,
            FileSystemConnections = { new FileSystemConnectionOptions { ConnectionId = "workspace", RootPath = folder } }
        };
        await Record("list_connections", "connections visible", OfficeAgentMcpServer.ConnectionsPayloadAsync(tools, mcpOptions),
            n => n.AsArray().Count > 0, "a non-empty array");
        await Record("list_connections", "no connection visible to the caller", OfficeAgentMcpServer.ConnectionsPayloadAsync(denied, mcpOptions),
            n => n.AsArray().Count == 0, "an empty array");

        // inspect_document
        await Record("inspect_document", "Word, populated collections", tools.InspectDocument("memory", template),
            n => Arr(n, "paragraphs").Count > 0 && Arr(n, "contentControls").Count > 0, "paragraphs and content controls");
        await Record("inspect_document", "Word, blank document", tools.InspectDocument("memory", blank),
            n => Arr(n, "contentControls").Count == 0, "no content controls");
        await Record("inspect_document", "PowerPoint deck", tools.InspectDocument("memory", deck), IsSuccess, "an inspection");
        await Record("inspect_document", "Excel workbook", tools.InspectDocument("memory", book), IsSuccess, "an inspection");
        await Record("inspect_document", "unknown document", tools.InspectDocument("memory", "no-such-document"),
            IsError, "an error envelope");
        await Record("inspect_document", "connection denied", denied.InspectDocument("memory", contract),
            n => HasCode(n, ToolErrorCodes.ConnectionForbidden), "connection-forbidden");
        await Record("inspect_document", "cancelled before start", tools.InspectDocument("memory", contract, cancellationToken: cancelled.Token),
            n => HasCode(n, ToolErrorCodes.Cancelled), "cancelled");

        // find_in_document
        await Record("find_in_document", "hits", tools.FindInDocument("memory", contract, "Northwind"),
            n => n.AsArray().Count > 0, "a non-empty array");
        await Record("find_in_document", "no hits", tools.FindInDocument("memory", contract, "zzzz-absent"),
            n => n is JsonArray { Count: 0 }, "an empty array");
        await Record("find_in_document", "unknown document", tools.FindInDocument("memory", "no-such-document", "x"),
            IsError, "an error envelope");

        // preview_plan
        await Record("preview_plan", "valid, changes proposed", tools.PreviewPlan("memory", contract, editPlan),
            n => Bool(n, "isValid") && Arr(n, "changes").Count > 0, "a valid preview with changes");
        await Record("preview_plan", "valid, empty plan", tools.PreviewPlan("memory", contract, emptyPlan),
            n => Bool(n, "isValid") && Arr(n, "changes").Count == 0, "a valid preview with no changes");
        await Record("preview_plan", "validation failure", tools.PreviewPlan("memory", contract, badPlan),
            n => !Bool(n, "isValid") && Arr(n, "errors").Count > 0 && Arr(n, "errors")[0]!["target"] is JsonObject,
            "errors with a target");
        await Record("preview_plan", "unreadable plan", tools.PreviewPlan("memory", contract, "{\"bogus\":1}"),
            n => HasCode(n, ToolErrorCodes.InvalidJson), "invalid-json");

        // apply_plan: success, rejection and every storage boundary the faulting provider can reach
        await Record("apply_plan", "committed, anonymous caller", tools.ApplyPlan("memory", contract, editPlan),
            n => Bool(n, "committed") && n["receipt"]!["actor"] is null && n["receipt"]!["outputSha256"] is not null,
            "a commit with a receipt and no actor");
        var afterCommit = EditPlan(await tools.InspectDocument("memory", contract));
        actor.Actor = new AuditActor { Subject = "user-1", Issuer = "https://issuer.example", DisplayName = "Fixture User" };
        await Record("apply_plan", "committed, authenticated actor", tools.ApplyPlan("memory", contract, afterCommit, saveMode: "NewVersion", newName: "contract-v2.docx"),
            n => Bool(n, "committed") && n["receipt"]!["actor"] is JsonObject, "a commit whose receipt carries an actor");
        actor.Actor = null;
        await Record("apply_plan", "validation failure", tools.ApplyPlan("memory", contract, badPlan),
            n => !Bool(n, "committed") && Arr(n, "errors").Count > 0, "a rejected plan");
        await Record("apply_plan", "unknown document", tools.ApplyPlan("memory", "no-such-document", editPlan),
            IsError, "an error envelope");
        await Record("apply_plan", "connection denied", denied.ApplyPlan("memory", contract, editPlan),
            n => HasCode(n, ToolErrorCodes.ConnectionForbidden), "connection-forbidden");

        foreach (var (fault, name, code) in new[]
                 {
                     (Fault.VersionConflict, "stale version at save", "version-conflict"),
                     (Fault.RefuseWrite, "storage refused the write", "io-error"),
                     (Fault.AfterAccept, "storage accepted, then failed to confirm", "io-error"),
                     (Fault.CancelAfterAccept, "cancelled after storage accepted", ToolErrorCodes.Cancelled),
                 })
        {
            // A fresh document per case: an accepted write really lands, and the next case
            // must not start from its text.
            var faultyDoc = faulty.Add("contract.docx", contractBytes);
            var faultyPlan = EditPlan(await tools.InspectDocument("faulty", faultyDoc));
            faulty.Next = fault;
            await Record("apply_plan", name, tools.ApplyPlan("faulty", faultyDoc, faultyPlan),
                n => IsError(n) && (code == "io-error" || HasCode(n, code)), $"an error envelope ({code})");
            faulty.Next = Fault.None;
        }

        // create_document
        await Record("create_document", "created blank", tools.CreateDocument("memory", "second.docx"),
            n => Bool(n, "committed") && Arr(n, "changes").Count == 0, "a blank creation");
        await Record("create_document", "created from a plan", tools.CreateDocument("memory", "third.docx", CreatePlan(null)),
            n => Bool(n, "committed") && Arr(n, "changes").Count > 0, "a creation with changes");
        await Record("create_document", "validation failure", tools.CreateDocument("memory", "fourth.docx", badPlan),
            n => !Bool(n, "committed") && Arr(n, "errors").Count > 0, "a rejected initial plan");
        await Record("create_document", "name already exists", tools.CreateDocument("memory", "blank.docx"),
            n => HasCode(n, ToolErrorCodes.AlreadyExists), "already-exists");
        faulty.Next = Fault.AfterAccept;
        await Record("create_document", "storage accepted, then failed to confirm", tools.CreateDocument("faulty", "made.docx"),
            IsError, "an error envelope");
        faulty.Next = Fault.None;

        // register_document, open_document, edit_document, remove_document
        await Record("register_document", "registered", tools.RegisterDocument("workspace", "contract.docx"),
            n => n["documentId"] is not null, "a document id");
        await Record("register_document", "file not found", tools.RegisterDocument("workspace", "missing.docx"),
            IsError, "an error envelope");
        await Record("register_document", "path outside the root", tools.RegisterDocument("workspace", "../outside.docx"),
            IsError, "an error envelope");
        var opened = await Record("open_document", "opened", tools.OpenDocument("workspace", "contract.docx"),
            n => n["documentId"] is not null, "a document id and inspection");
        await Record("open_document", "file not found", tools.OpenDocument("workspace", "missing.docx"),
            IsError, "an error envelope");
        await Record("edit_document", "committed", tools.EditDocument("workspace", "contract.docx", EditPlan(opened)),
            n => Bool(n, "committed") && n["sourceDocumentId"] is not null, "a commit naming its source");
        await Record("edit_document", "validation failure", tools.EditDocument("workspace", "contract.docx", badPlan),
            n => !Bool(n, "committed") && n["sourceDocumentId"] is not null && Arr(n, "errors").Count > 0,
            "errors that still name the source document");
        // A text target is bound before validation, and a miss has its own envelope.
        await Record("edit_document", "find target not bound", tools.EditDocument("workspace", "contract.docx",
                "{\"operations\":[{\"op\":\"changeText\",\"target\":{\"find\":\"zzzz-absent\"},\"with\":\"x\"}]}"),
            n => !Bool(n, "committed") && Arr(n, "errors").Count > 0 && Arr(n, "errors")[0]!["itemId"] is not null,
            "binding errors naming the source document");
        var removable = Str(JsonNode.Parse(await tools.CreateDocument("memory", "removable.docx"))!, "outputDocumentId");
        await Record("remove_document", "removed", tools.RemoveDocument("memory", removable),
            n => n["removed"]?.GetValue<bool>() == true, "removed");
        await Record("remove_document", "unknown document", tools.RemoveDocument("memory", "no-such-document"),
            IsError, "an error envelope");

        // inline content and the session connection
        var inline = await Record("create_document_content", "created", tools.CreateDocumentContent("inline.docx", CreatePlan(null)),
            n => Bool(n, "committed") && n["contentBase64"] is not null, "content");
        await Record("create_document_content", "validation failure", tools.CreateDocumentContent("inline.docx", badPlan),
            n => !Bool(n, "committed") && n["contentBase64"] is null, "no content");
        var inlineContent = Str(JsonNode.Parse(inline)!, "contentBase64");
        var inlineInspect = await Record("inspect_document_content", "inspected", tools.InspectDocumentContent(inlineContent),
            IsSuccess, "an inspection");
        await Record("inspect_document_content", "content is not a package", tools.InspectDocumentContent(Convert.ToBase64String(new byte[] { 1, 2, 3 })),
            IsError, "an error envelope");
        await Record("edit_document_content", "committed", tools.EditDocumentContent(inlineContent, EditPlan(inlineInspect)),
            n => Bool(n, "committed") && n["contentBase64"] is not null, "content");
        await Record("edit_document_content", "preview only", tools.EditDocumentContent(inlineContent, EditPlan(inlineInspect), preview: true),
            n => !Bool(n, "committed") && n["contentBase64"] is null && Bool(n, "isValid"), "a valid preview with no content");
        await Record("edit_document_content", "validation failure", tools.EditDocumentContent(inlineContent, badPlan),
            n => !Bool(n, "isValid"), "a rejected plan");
        await Record("import_document_content", "imported", tools.ImportDocumentContent("memory", "imported.docx", inlineContent),
            n => n["documentId"] is not null, "a document id");
        await Record("import_document_content", "content is not a package", tools.ImportDocumentContent("memory", "bad.docx", Convert.ToBase64String(new byte[] { 1, 2, 3 })),
            IsError, "an error envelope");
        await Record("export_document_content", "exported", tools.ExportDocumentContent("memory", contract),
            n => n["contentBase64"] is not null, "content");
        await Record("export_document_content", "unknown document", tools.ExportDocumentContent("memory", "no-such-document"),
            IsError, "an error envelope");

        // compare_documents
        var original = contract;
        var revised = Str(JsonNode.Parse(await tools.CreateDocument("memory", "revised.docx",
            CreatePlan("Direct", "Prepared for Contoso Pharmaceuticals.")))!, "outputDocumentId");
        await Record("compare_documents", "identical documents", tools.CompareDocuments("memory", original, "memory", original),
            n => Arr(n, "differences").Count == 0, "no differences");
        await Record("compare_documents", "differences found", tools.CompareDocuments("memory", original, "memory", revised),
            n => Arr(n, "differences").Count > 0, "differences");
        await Record("compare_documents", "unknown document", tools.CompareDocuments("memory", original, "memory", "no-such-document"),
            IsError, "an error envelope");

        // template tools, including a partial batch and an uncertain item
        await Record("discover_template", "slots found", tools.DiscoverTemplate("memory", template),
            n => Arr(n, "slots").Count > 0 && Arr(n, "repeatingRows").Count > 0, "slots and repeating rows");
        await Record("discover_template", "no slots", tools.DiscoverTemplate("memory", blank),
            n => Arr(n, "slots").Count == 0, "no slots");
        var batchA = Batch("quote-a.docx");
        var previewA = await Record("preview_template_batch", "valid", tools.PreviewTemplateBatch("memory", template, batchA),
            n => Bool(n, "isValid") && n["token"] is JsonObject, "a valid preview with a token");
        await Record("preview_template_batch", "invalid binding", tools.PreviewTemplateBatch("memory", template,
                "{\"items\":[{\"outputName\":\"quote-x.docx\",\"binding\":{\"values\":{\"NoSuchSlot\":\"x\"},\"missingValueBehavior\":\"Fail\"}}]}"),
            n => !Bool(n, "isValid"), "an invalid preview");
        await Record("populate_template_batch", "every item committed",
            tools.PopulateTemplateBatch("memory", template, batchA, Json(previewA, "token")),
            n => Arr(n, "items").All(i => Str(i!, "outcome") == "Committed"), "all items committed");
        var batchB = Batch("quote-b.docx");
        var previewB = await tools.PreviewTemplateBatch("memory", template, batchB);
        await Record("populate_template_batch", "stale preview token",
            tools.PopulateTemplateBatch("memory", template, batchB, Json(previewA, "token")),
            n => !Bool(n, "committed"), "a refused commit");

        var faultyTemplate = faulty.Add("quote.docx", QuoteTemplate());
        var pair = Batch("pair-a.docx", "pair-b.docx");
        var pairPreview = await tools.PreviewTemplateBatch("faulty", faultyTemplate, pair);
        faulty.RefuseCreateNamed = "pair-b.docx";
        await Record("populate_template_batch", "partial: one committed, one refused before any write",
            tools.PopulateTemplateBatch("faulty", faultyTemplate, pair, Json(pairPreview, "token")),
            n => Arr(n, "items").Any(i => Str(i!, "outcome") == "Committed") &&
                 Arr(n, "items").Any(i => Str(i!, "outcome") != "Committed"), "mixed item outcomes");
        faulty.RefuseCreateNamed = null;
        var single = Batch("uncertain.docx");
        var singlePreview = await tools.PreviewTemplateBatch("faulty", faultyTemplate, single);
        faulty.Next = Fault.AfterAccept;
        await Record("populate_template_batch", "item uncertain after storage accepted",
            tools.PopulateTemplateBatch("faulty", faultyTemplate, single, Json(singlePreview, "token")),
            n => Arr(n, "items").Any(i => Str(i!, "outcome") == "Uncertain"), "an uncertain item");
        faulty.Next = Fault.None;

        // merge
        var left = Str(JsonNode.Parse(await tools.CreateDocument("memory", "left.docx", CreatePlan("Direct")))!, "outputDocumentId");
        var right = Str(JsonNode.Parse(await tools.CreateDocument("memory", "right.docx", CreatePlan("Direct")))!, "outputDocumentId");
        var mergeRequest = "{\"sources\":[{\"connectionId\":\"memory\",\"documentId\":\"" + left + "\"}," +
                           "{\"connectionId\":\"memory\",\"documentId\":\"" + right + "\"}]}";
        var mergePreview = await Record("preview_document_merge", "valid", tools.PreviewDocumentMerge(mergeRequest),
            n => Bool(n, "isValid") && n["plan"] is JsonObject, "a valid plan");
        await Record("preview_document_merge", "unsupported source content", tools.PreviewDocumentMerge(
                "{\"sources\":[{\"connectionId\":\"memory\",\"documentId\":\"" + contract + "\"}," +
                "{\"connectionId\":\"memory\",\"documentId\":\"" + deck + "\"}]}"),
            n => !Bool(n, "isValid") && n["plan"] is null, "an invalid preview with no plan");
        var mergePlan = Json(mergePreview, "plan");
        await Record("merge_documents", "committed", tools.MergeDocuments(mergePlan, "memory", "packet.docx"),
            n => Bool(n, "committed") && n["receipt"] is JsonObject, "a commit with a receipt");
        faulty.Next = Fault.AfterAccept;
        await Record("merge_documents", "storage accepted, then failed to confirm", tools.MergeDocuments(mergePlan, "faulty", "packet.docx"),
            IsError, "an error envelope");
        faulty.Next = Fault.None;
        await tools.ApplyPlan("memory", left, EditPlan(await tools.InspectDocument("memory", left)));
        await Record("merge_documents", "source changed since preview", tools.MergeDocuments(mergePlan, "memory", "stale.docx"),
            n => !Bool(n, "committed"), "a refused commit");

        return states;
    }

    // ── fixture helpers ────────────────────────────────────────────────────────────────────

    private static string CreatePlan(string? mode, string text = "Prepared for Northwind Labs.") =>
        "{\"operations\":[{\"op\":\"changeText\",\"target\":{\"paraId\":\"auto-0000\",\"expect\":\"\"}," +
        (mode is null ? "" : "\"mode\":\"" + mode + "\",") + "\"with\":\"" + text + "\"}]}";

    private static string Batch(params string[] outputs) =>
        "{\"items\":[" + string.Join(",", outputs.Select(output =>
            "{\"outputName\":\"" + output + "\",\"binding\":{\"values\":{\"CustomerName\":\"Fabrikam\"},\"missingValueBehavior\":\"Ignore\"}}")) + "]}";

    /// <summary>A one-operation plan that edits the first non-empty paragraph an inspection reported.</summary>
    internal static string EditPlan(string inspection)
    {
        var root = JsonNode.Parse(inspection)!;
        var inspected = root["paragraphs"] is not null ? root : root["inspection"]!;
        var paragraph = inspected["paragraphs"]!.AsArray()
            .First(item => (item!["text"]?.GetValue<string>() ?? "").Length > 0)!;
        return "{\"operations\":[{\"op\":\"changeText\",\"target\":{\"paraId\":\"" +
               paragraph["paraId"]!.GetValue<string>() + "\",\"expect\":\"" +
               paragraph["text"]!.GetValue<string>() + "\"},\"mode\":\"Direct\",\"with\":\"Revised for Contoso.\"}]}";
    }

    private static bool IsSuccess(JsonNode node) => node is JsonObject obj && !obj.ContainsKey("errors") || Arr(node, "errors").Count == 0;

    private static bool IsError(JsonNode node) =>
        node is JsonObject obj && obj.ContainsKey("errors") && Arr(node, "errors").Count > 0 && !Bool(node, "committed");

    private static bool HasCode(JsonNode node, string code) =>
        IsError(node) && Arr(node, "errors").Any(error => error!["code"]?.GetValue<string>() == code);

    private static JsonArray Arr(JsonNode node, string name) => node[name]?.AsArray() ?? new JsonArray();

    private static bool Bool(JsonNode node, string name) => node[name]?.GetValue<bool>() == true;

    private static string Str(JsonNode node, string name) =>
        node[name]?.GetValue<string>() ?? throw new InvalidOperationException($"Fixture response has no '{name}': {node.ToJsonString()}");

    private static string Json(string response, string name) =>
        JsonNode.Parse(response)![name]?.ToJsonString() ?? throw new InvalidOperationException($"Fixture response has no '{name}': {response}");

    /// <summary>A template with one content-control slot and one repeating table row.</summary>
    internal static byte[] QuoteTemplate()
    {
        static W.TableRow Row(params string[] values) => new(values.Select(value =>
            new W.TableCell(new W.Paragraph(new W.Run(new W.Text(value))))));

        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new W.Document(new W.Body(
                new W.Paragraph(
                    new W.Run(new W.Text("Quote for ")),
                    new W.SdtRun(
                        new W.SdtProperties(new W.Tag { Val = "CustomerName" }, new W.SdtId { Val = 11 }, new W.SdtContentText()),
                        new W.SdtContentRun(new W.Run(new W.Text("CUSTOMER"))))),
                new W.Table(
                    new W.TableGrid(new W.GridColumn { Width = "3600" }, new W.GridColumn { Width = "1200" }),
                    Row("Description", "Quantity"),
                    Row("{{Description}}", "{{Quantity}}")),
                new W.Paragraph()));
            main.Document.Save();
        }
        return stream.ToArray();
    }

    private sealed class SwitchableActor : IAuditActorProvider
    {
        public AuditActor? Actor { get; set; }

        public AuditActor? GetCurrentActor() => Actor;
    }

    private sealed class DenyAll : IConnectionAccessPolicy
    {
        public ValueTask<bool> IsAllowedAsync(ClaimsPrincipal principal, string connectionId,
            ConnectionCapability capability, CancellationToken cancellationToken = default) => new(false);
    }

    private enum Fault { None, VersionConflict, RefuseWrite, AfterAccept, CancelAfterAccept }

    /// <summary>
    /// A memory store whose next write fails at a chosen boundary. References carry a name and
    /// content type like the built-in providers', so fault states compare like for like with
    /// success states.
    /// </summary>
    private sealed class FaultingProvider : IDocumentCreatingProvider
    {
        private readonly Dictionary<string, (string Name, byte[] Bytes)> _items = new(StringComparer.Ordinal);
        private int _next;

        public FaultingProvider(string connectionId) => ConnectionId = connectionId;

        public string Provider => "faulting";

        public string ConnectionId { get; }

        public Fault Next { get; set; }

        public string? RefuseCreateNamed { get; set; }

        public string Add(string name, byte[] content)
        {
            var id = $"item-{++_next}";
            _items[id] = (name, content);
            return id;
        }

        public Task<DocumentReference> RegisterAsync(string source, CancellationToken cancellationToken = default) =>
            throw new DocumentProviderException(ProviderErrorCode.InvalidArgument, "Registration is not supported.", Provider, ConnectionId, null);

        public Task<DocumentContent> OpenReadAsync(DocumentReference reference, CancellationToken cancellationToken = default)
        {
            if (!_items.TryGetValue(reference.ItemId, out var item))
                throw new DocumentProviderException(ProviderErrorCode.NotFound, "Not found.", Provider, ConnectionId, reference.ItemId);
            return Task.FromResult(new DocumentContent(Reference(reference.ItemId), new MemoryStream(item.Bytes, writable: false)));
        }

        public async Task<DocumentReference> SaveAsync(DocumentReference source, Stream content, SaveDocumentOptions options,
            CancellationToken cancellationToken = default)
        {
            if (options.NewName is { } newName && newName == RefuseCreateNamed)
                throw new DocumentProviderException(ProviderErrorCode.AlreadyExists,
                    $"A document named '{newName}' already exists.", Provider, ConnectionId, null);
            var bytes = await Before(source.ItemId, content, cancellationToken);
            var id = options.Mode == SaveMode.Replace ? source.ItemId : $"item-{++_next}";
            _items[id] = (options.NewName ?? _items[source.ItemId].Name, bytes);
            After(id);
            return Reference(id);
        }

        public async Task<DocumentReference> CreateAsync(string name, Stream content, CancellationToken cancellationToken = default)
        {
            if (name == RefuseCreateNamed)
                throw new DocumentProviderException(ProviderErrorCode.IO, "The storage refused the write.", Provider, ConnectionId, null);
            var bytes = await Before(null, content, cancellationToken);
            var id = $"item-{++_next}";
            _items[id] = (name, bytes);
            After(id);
            return Reference(id);
        }

        public Task RemoveAsync(DocumentReference reference, CancellationToken cancellationToken = default)
        {
            _items.Remove(reference.ItemId);
            return Task.CompletedTask;
        }

        private async Task<byte[]> Before(string? itemId, Stream content, CancellationToken cancellationToken)
        {
            if (Next == Fault.VersionConflict)
                throw new DocumentVersionConflictException("expected", "actual", Provider, ConnectionId, itemId);
            if (Next == Fault.RefuseWrite)
                throw new DocumentProviderException(ProviderErrorCode.IO, "The storage refused the write.", Provider, ConnectionId, itemId);
            using var buffer = new MemoryStream();
            await content.CopyToAsync(buffer, cancellationToken);
            return buffer.ToArray();
        }

        private void After(string id)
        {
            if (Next == Fault.AfterAccept)
                throw new DocumentProviderException(ProviderErrorCode.IO,
                    "The storage accepted the document but the result could not be confirmed.", Provider, ConnectionId, id);
            if (Next == Fault.CancelAfterAccept)
                throw new OperationCanceledException("Cancelled after the storage accepted the document.");
        }

        private DocumentReference Reference(string id) => new()
        {
            Provider = Provider,
            ConnectionId = ConnectionId,
            ItemId = id,
            Version = Convert.ToHexString(SHA256.HashData(_items[id].Bytes)),
            Name = _items[id].Name,
            ContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        };
    }
}
