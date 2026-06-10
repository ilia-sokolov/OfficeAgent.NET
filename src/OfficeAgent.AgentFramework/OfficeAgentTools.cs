using System.Text.Json;
using System.Text.Json.Serialization;
using Microsoft.Extensions.AI;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;

namespace OfficeAgent.AgentFramework;

/// <summary>
/// Projects <see cref="OfficeAgentClient"/> as Microsoft.Extensions.AI tools that
/// address documents by an opaque, provider-assigned id. The host registers documents
/// with a provider connection (<see cref="OfficeAgentClient.RegisterAsync"/>),
/// receives a <see cref="DocumentReference"/>, and the LLM drives inspect /
/// find / preview / apply by <c>(connectionId, documentId)</c>. The agent never
/// sees a file path; it cannot register documents, escape the connection, or
/// delete content the provider only references.
/// </summary>
public sealed class OfficeAgentTools
{
    private static readonly JsonSerializerOptions Json = new() { WriteIndented = false };

    private static readonly JsonSerializerOptions PlanJson = new()
    {
        PropertyNameCaseInsensitive = true,
        Converters = { new JsonStringEnumConverter() }
    };

    internal static readonly AIJsonSchemaCreateOptions StrictSchemaOptions = new()
    {
        TransformOptions = new AIJsonSchemaTransformOptions
        {
            RequireAllProperties = true,
            DisallowAdditionalProperties = true
        }
    };

    private readonly OfficeAgentClient _client;

    /// <summary>Initializes the tool projection over an <see cref="OfficeAgentClient"/>.</summary>
    public OfficeAgentTools(OfficeAgentClient client) => _client = client;

    /// <summary>
    /// System-prompt guidance to concatenate into the agent's instructions. Teaches
    /// the model the <c>(connectionId, documentId)</c> contract, the safety loop,
    /// and the structured-error vocabulary the tools surface.
    /// </summary>
    public const string SystemPromptGuidance = """
        You are editing a Microsoft Word document through the OfficeAgent tools.

        Document addressing
        - Storage connections are host-configured and so is each document's registration. The host gives you an OPAQUE, provider-assigned documentId for every document you are allowed to work with; all document tools address it as (connectionId, documentId). Never invent a documentId, never pass a filename or path as one, and never ask the user to send raw file bytes through this conversation.
        - The connectionId and documentId are already in your instructions or in the conversation context. NEVER ask the user for them - the user does not know or manage these values. If a request mentions "the document", it means the current document you were given; start working with it immediately.
        - apply_plan returns outputDocumentId, outputName, and outputContentType for the saved revision. Use outputDocumentId as the next call's documentId if you keep editing. When the work is complete, tell the user the document is ready; the host retrieves its bytes and presents the download or attachment. Do not place document base64 in the final response.

        Plan shape, anchors, safety loop
        - Plan body is { "operations": [ ... ] }. Do NOT set contractVersion or snapshot - the engine fills them.
        - Call inspect_document or find_in_document before building a plan to obtain anchor ids; never invent paragraph ids, occurrence numbers, content-control tags, or node paths.
        - Tables and images only appear in inspect_document.nodes (kind="table"/"image", path="table#N"/"image#N"). They are NOT in the paragraphs list. To recognise table content, look for paragraphs whose `in` field equals a table path.
        - Preview before you apply. If preview reports stale-snapshot, re-inspect and rebuild. If preview reports expect-mismatch, the document drifted - re-inspect/find that operation.
        - Default to ChangeMode "Tracked" unless the user explicitly approves direct edits.
        - Reject operations that need a renderer (pagination, field recalculation); explain the limitation instead.
        """;

    /// <summary>Returns the four AIFunctions the host registers with its agent.</summary>
    public AIFunction[] AsAIFunctions() => new[]
    {
        AIFunctionFactory.Create(InspectDocument, Opts(
            "inspect_document",
            "Inspect a Word document by (connectionId, documentId). Returns outline, paragraphs (with their `in` table containment), content controls, nodes (tables/images/docProperties/revisions - paths for table-row and image operations come from here), styles, and a snapshot etag for drift detection. Use paragraphOffset/paragraphLimit to page; fidelity='outline'|'structure'|'content' to control payload size.")),
        AIFunctionFactory.Create(FindInDocument, Opts(
            "find_in_document",
            "Find text in a Word document by (connectionId, documentId). Returns content-verified anchors (paragraphId + expected + occurrence) usable as plan targets.")),
        AIFunctionFactory.Create(PreviewPlan, Opts(
            "preview_plan",
            "Dry-run a DocumentPlan JSON against (connectionId, documentId). Returns {isValid, changes, errors} without writing. " +
            "Plan shape: { \"operations\": [ ... ] }. Do NOT set contractVersion or snapshot. Each operation is one object. Concrete examples:\n\n" +
            "// Replace text:\n" +
            "{ \"op\": \"changeText\", \"target\": { \"paraId\": \"w14:...\", \"expect\": \"Acme Corp\", \"occurrence\": 0 }, \"with\": \"Globex Inc.\", \"mode\": \"Tracked\" }\n\n" +
            "// Unified formatting (paragraph/run/table/row/cell/image):\n" +
            "{ \"op\": \"format\", \"target\": { \"paraId\": \"w14:...\", \"expect\": \"important\", \"occurrence\": 0 }, \"highlight\": \"yellow\", \"bold\": true, \"color\": \"FF0000\" }\n" +
            "{ \"op\": \"format\", \"target\": { \"kind\": \"table\",     \"path\": \"table#0\" }, \"styleId\": \"TableGrid\", \"borderStyle\": \"single\" }\n" +
            "{ \"op\": \"format\", \"target\": { \"kind\": \"image\",     \"path\": \"image#0\" }, \"widthPx\": 320, \"heightPx\": 200 }\n\n" +
            "// Fill / comment / insert paragraph or new table / setProperty / revision:\n" +
            "{ \"op\": \"fill\", \"target\": { \"tag\": \"ClientName\" }, \"value\": \"Globex\" }\n" +
            "{ \"op\": \"comment\", \"target\": { \"paraId\": \"w14:...\", \"expect\": \"...\" }, \"text\": \"Confirm this.\" }\n" +
            "{ \"op\": \"insert\", \"target\": { \"paraId\": \"w14:...\", \"expect\": \"...\" }, \"position\": \"After\", \"text\": \"New paragraph.\" }\n" +
            "{ \"op\": \"setProperty\", \"target\": { \"kind\": \"docProperty\", \"path\": \"core/title\" }, \"value\": \"My Title\" }\n" +
            "{ \"op\": \"revision\",   \"target\": { \"kind\": \"revision\",    \"path\": \"all\" }, \"action\": \"Accept\" }\n\n" +
            "// Add or remove table rows / columns; insert or remove image; copy or clear styles. Paths come from inspect_document.nodes:\n" +
            "{ \"op\": \"insertTableRows\", \"target\": { \"kind\": \"table\", \"path\": \"table#0\" }, \"rows\": [[\"NL\",\"17\",\"41850\"]], \"position\": \"End\" }\n" +
            "{ \"op\": \"removeTableRows\", \"target\": { \"kind\": \"table\", \"path\": \"table#0\" }, \"onlyIfEmpty\": true }\n" +
            "{ \"op\": \"insertImage\", \"target\": { \"paraId\": \"w14:...\", \"expect\": \"...\" }, \"base64Bytes\": \"iVBORw0KGgo...\", \"imageType\": \"png\", \"widthPx\": 200, \"heightPx\": 80 }\n" +
            "{ \"op\": \"insertImage\", \"target\": { \"paraId\": \"w14:...\", \"expect\": \"...\" }, \"imageConnectionId\": \"images\", \"imageDocumentId\": \"<opaque id from a prior add>\", \"imageType\": \"png\", \"widthPx\": 200, \"heightPx\": 80 }\n" +
            "{ \"op\": \"removeImage\", \"target\": { \"kind\": \"image\", \"path\": \"image#0\" } }")),
        AIFunctionFactory.Create(ApplyPlan, Opts(
            "apply_plan",
            "Apply a DocumentPlan JSON to (connectionId, documentId) and save through the provider. Returns {committed, outputConnectionId, outputDocumentId, outputVersion, outputName, outputContentType, changes, errors}. saveMode: 'NewVersion' (default, mints a new id under the same connection), 'NewDocument' (mints a fresh id with an optional newName for display), 'Replace' (overwrites the source after an optimistic version check). On any failure nothing is written."))
    };

    /// <summary>Inspects a document and returns paginated JSON.</summary>
    public Task<string> InspectDocument(
        string connectionId,
        string documentId,
        string fidelity = "content",
        int paragraphOffset = 0,
        int paragraphLimit = 200,
        CancellationToken cancellationToken = default)
        => SafeAsync(async () =>
        {
            var options = new InspectOptions { Fidelity = ParseFidelity(fidelity) };
            var result = await _client.InspectAsync(connectionId, documentId, options, cancellationToken).ConfigureAwait(false);

            var pagedParagraphs = result.Paragraphs
                .Skip(Math.Max(0, paragraphOffset))
                .Take(Math.Max(0, paragraphLimit))
                .Select(p => new { p.ParaId, style = p.StyleId, p.Text, @in = p.In });

            return JsonSerializer.Serialize(new
            {
                format = result.Format.ToString(),
                snapshot = result.Snapshot.ETag,
                outline = result.Outline.Select(MapOutline),
                paragraphsTotal = result.Paragraphs.Count,
                paragraphOffset,
                paragraphLimit,
                paragraphs = pagedParagraphs,
                contentControls = result.StructuralAnchors.Select(s => new { s.Tag, s.Kind }),
                nodes = result.Nodes.Select(n => new { n.Kind, n.Path, n.Summary }),
                styles = result.Styles.Styles.Select(s => new { s.Id, s.Name, s.InUseCount })
            }, Json);
        });

    /// <summary>Finds text in a document and returns content-verified anchors.</summary>
    public Task<string> FindInDocument(
        string connectionId,
        string documentId,
        string pattern,
        bool regex = false,
        bool wholeWord = false,
        bool caseSensitive = false,
        CancellationToken cancellationToken = default)
        => SafeAsync(async () =>
        {
            var query = new FindQuery
            {
                Pattern = pattern,
                Options = new MatchOptions { Regex = regex, WholeWord = wholeWord, CaseSensitive = caseSensitive }
            };
            var hits = await _client.FindAsync(connectionId, documentId, query, cancellationToken).ConfigureAwait(false);
            return JsonSerializer.Serialize(hits.Select(h => new
            {
                paraId = (h.Anchor as TextSpanAnchor)?.ParaId,
                expect = h.Text,
                occurrence = (h.Anchor as TextSpanAnchor)?.Occurrence ?? 0,
                context = h.Context
            }), Json);
        });

    /// <summary>Dry-runs a plan against the document.</summary>
    public Task<string> PreviewPlan(
        string connectionId,
        string documentId,
        string planJson,
        CancellationToken cancellationToken = default)
        => SafeAsync(async () =>
        {
            var plan = DeserializePlan(planJson);
            var report = await _client.PreviewAsync(connectionId, documentId, plan, cancellationToken).ConfigureAwait(false);
            return SerializeReport(report, committed: false, savedReference: null);
        });

    /// <summary>Applies a plan and saves through the provider.</summary>
    public Task<string> ApplyPlan(
        string connectionId,
        string documentId,
        string planJson,
        string saveMode = "NewVersion",
        string newName = "",
        CancellationToken cancellationToken = default)
        => SafeAsync(async () =>
        {
            var plan = DeserializePlan(planJson);
            var options = new SaveDocumentOptions
            {
                Mode = ParseSaveMode(saveMode),
                NewName = string.IsNullOrEmpty(newName) ? null : newName
            };
            var result = await _client.CommitAsync(connectionId, documentId, plan, options, cancellationToken).ConfigureAwait(false);
            return SerializeReport(result.Report, result.Committed, result.Committed ? result.Document : null);
        });

    private static async Task<string> SafeAsync(Func<Task<string>> work)
    {
        try { return await work().ConfigureAwait(false); }
        catch (OperationCanceledException) { return SerializeError("cancelled", "Operation was cancelled."); }
        catch (JsonException ex) { return SerializeError("invalid-json", ex.Message); }
        catch (DocumentProviderException ex)
        {
            return SerializeError(
                ProviderCodeToWire(ex.Code),
                ex.Message,
                ex.Provider, ex.ConnectionId, ex.ItemId);
        }
        catch (Exception ex) { return SerializeError("internal-error", ex.Message); }
    }

    private static AIFunctionFactoryOptions Opts(string name, string description) => new()
    {
        Name = name,
        Description = description,
        JsonSchemaCreateOptions = StrictSchemaOptions
    };

    private static Fidelity ParseFidelity(string fidelity) => fidelity?.ToLowerInvariant() switch
    {
        "outline" => Fidelity.Outline,
        "structure" => Fidelity.Structure,
        _ => Fidelity.Content
    };

    private static SaveMode ParseSaveMode(string mode) => mode?.Trim() switch
    {
        "NewDocument" => SaveMode.NewDocument,
        "Replace" => SaveMode.Replace,
        _ => SaveMode.NewVersion
    };

    private static string ProviderCodeToWire(ProviderErrorCode code) => code switch
    {
        ProviderErrorCode.NotFound => "not-found",
        ProviderErrorCode.AccessDenied => "access-denied",
        ProviderErrorCode.ContentTooLarge => "content-too-large",
        ProviderErrorCode.ExtensionNotAllowed => "extension-not-allowed",
        ProviderErrorCode.VersionConflict => "version-conflict",
        ProviderErrorCode.InvalidArgument => "invalid-argument",
        ProviderErrorCode.ConfigurationError => "configuration-error",
        ProviderErrorCode.IO => "io-error",
        _ => "provider-error"
    };

    private static DocumentPlan DeserializePlan(string planJson) =>
        JsonSerializer.Deserialize<DocumentPlan>(planJson, PlanJson)
        ?? throw new JsonException("Plan JSON did not deserialize to a DocumentPlan.");

    private static string SerializeReport(ChangeReport report, bool committed, DocumentReference? savedReference) =>
        JsonSerializer.Serialize(new
        {
            isValid = report.IsValid,
            committed,
            outputConnectionId = savedReference?.ConnectionId,
            outputDocumentId = savedReference?.ItemId,
            outputVersion = savedReference?.Version,
            outputName = savedReference?.Name,
            outputContentType = savedReference?.ContentType,
            changes = report.Changes.Select(c => new
            {
                c.Verb,
                target = SummariseAnchor(c.Target),
                c.Before, c.After, c.Context, c.BlastRadius,
                capability = c.Capability.ToString()
            }),
            errors = report.Errors.Select(e => new
            {
                e.Code, e.Message,
                target = SummariseAnchor(e.Target)
            })
        }, Json);

    private static string SerializeError(string code, string message, string? provider = null, string? connectionId = null, string? itemId = null) =>
        JsonSerializer.Serialize(new
        {
            isValid = false,
            committed = false,
            outputConnectionId = (string?)null,
            outputDocumentId = (string?)null,
            outputVersion = (string?)null,
            outputName = (string?)null,
            outputContentType = (string?)null,
            changes = Array.Empty<object>(),
            errors = new[] { new { Code = code, Message = message, target = (object?)null, provider, connectionId, itemId } }
        }, Json);

    private static object? SummariseAnchor(Anchor? anchor) => anchor switch
    {
        null => null,
        TextSpanAnchor t => new { kind = "textSpan", paraId = t.ParaId, expect = t.Expect, occurrence = t.Occurrence },
        StructuralAnchor s => new { kind = "structural", tag = s.Tag, structuralKind = s.Kind },
        NodeAnchor n => new { kind = "node", nodeKind = n.Kind, path = n.Path },
        StyleAnchor s => new { kind = "style", styleId = s.StyleId },
        _ => new { kind = anchor.GetType().Name, anchor.Id }
    };

    private static object MapOutline(OutlineNode node) => new
    {
        node.Level,
        node.Text,
        paraId = (node.Anchor as TextSpanAnchor)?.ParaId,
        children = node.Children.Select(MapOutline)
    };
}
