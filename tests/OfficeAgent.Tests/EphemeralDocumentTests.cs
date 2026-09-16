using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.Abstractions;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Mcp;
using OfficeAgent.PowerPoint;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// Documents the server holds for the session, addressed by an opaque id like any other.
/// </summary>
/// <remarks>
/// This exists because of what a live model did with the inline tools: asked to chain two
/// edits, it reproduced 2,928 characters of base64 with one character wrong, and every call
/// after that failed on content that was no longer a package. An id cannot be got wrong,
/// and the bytes never enter the conversation at all.
/// </remarks>
public class EphemeralDocumentTests
{
    [Fact]
    public async Task A_document_is_created_edited_and_exported_without_any_storage()
    {
        using var session = new Session();

        var created = await Json(session.Tools.CreateDocument("session", "agreement.docx", """
            { "operations": [
                { "op": "changeText", "target": { "paraId": "auto-0000", "expect": "" },
                  "with": "Master Services Agreement", "mode": "Direct" } ] }
            """));
        var documentId = created.GetProperty("outputDocumentId").GetString()!;

        // Saving stabilises the paragraph ids, so the positional one the create plan used
        // is gone; this is the ordinary inspect-then-edit loop, done by id.
        var inspected = await Json(session.Tools.InspectDocument("session", documentId));
        var headingId = inspected.GetProperty("paragraphs")[0].GetProperty("ParaId").GetString()!;

        // Every edit after this names the id, never the bytes.
        for (var step = 0; step < 5; step++)
        {
            var applied = await Json(session.Tools.ApplyPlan("session", documentId, $$"""
                { "operations": [
                    { "op": "insert", "target": { "paraId": "{{headingId}}", "expect": "Master Services Agreement" },
                      "position": "After", "text": "Clause {{step}}.", "mode": "Direct" } ] }
                """));
            Assert.True(applied.GetProperty("committed").GetBoolean(), applied.GetProperty("errors").ToString());
            documentId = applied.GetProperty("outputDocumentId").GetString()!;
        }

        var exported = await Json(session.Tools.ExportDocumentContent("session", documentId));
        var bytes = Convert.FromBase64String(exported.GetProperty("contentBase64").GetString()!);

        using var stream = new MemoryStream(bytes);
        using var package = WordprocessingDocument.Open(stream, isEditable: false);
        var text = package.MainDocumentPart!.Document.Body!.InnerText;

        Assert.Contains("Master Services Agreement", text);
        for (var step = 0; step < 5; step++) Assert.Contains($"Clause {step}.", text);
    }

    [Fact]
    public async Task A_document_supplied_once_is_edited_by_id_from_then_on()
    {
        using var session = new Session();

        var imported = await Json(session.Tools.ImportDocumentContent(
            "session", "contract.docx", Convert.ToBase64String(DocxFactory.Contract())));

        var documentId = imported.GetProperty("documentId").GetString()!;
        Assert.Equal("contract.docx", imported.GetProperty("name").GetString());

        var applied = await Json(session.Tools.ApplyPlan("session", documentId, """
            { "operations": [
                { "op": "changeText",
                  "target": { "paraId": "w14:00000002", "expect": "Acme Corp", "occurrence": 0 },
                  "with": "Globex Inc.", "mode": "Direct" } ] }
            """));
        Assert.True(applied.GetProperty("committed").GetBoolean(), applied.GetProperty("errors").ToString());

        var exported = await Json(session.Tools.ExportDocumentContent(
            "session", applied.GetProperty("outputDocumentId").GetString()!));

        using var stream = new MemoryStream(Convert.FromBase64String(exported.GetProperty("contentBase64").GetString()!));
        using var package = WordprocessingDocument.Open(stream, isEditable: false);
        Assert.Contains("Globex Inc.", package.MainDocumentPart!.Document.Body!.InnerText);
    }

    [Fact]
    public async Task Inspect_and_find_work_against_a_session_document()
    {
        using var session = new Session();
        var documentId = await session.Import(DocxFactory.Contract(), "contract.docx");

        var inspected = await Json(session.Tools.InspectDocument("session", documentId));
        Assert.Equal("Word", inspected.GetProperty("format").GetString());

        var hits = await Json(session.Tools.FindInDocument("session", documentId, "Acme Corp"));
        Assert.Equal(2, hits.GetArrayLength());
    }

    [Fact]
    public async Task A_tracked_edit_in_a_session_document_is_a_redline_like_anywhere_else()
    {
        using var session = new Session();
        var documentId = await session.Import(DocxFactory.Contract(), "contract.docx");

        // No mode stated: the session connection carries the same Tracked default.
        var applied = await Json(session.Tools.ApplyPlan("session", documentId, """
            { "operations": [
                { "op": "changeText",
                  "target": { "paraId": "w14:00000002", "expect": "Acme Corp", "occurrence": 0 },
                  "with": "Globex Inc." } ] }
            """));

        var exported = await Json(session.Tools.ExportDocumentContent(
            "session", applied.GetProperty("outputDocumentId").GetString()!));

        using var stream = new MemoryStream(Convert.FromBase64String(exported.GetProperty("contentBase64").GetString()!));
        using var package = WordprocessingDocument.Open(stream, isEditable: false);
        Assert.NotEmpty(package.MainDocumentPart!.Document.Body!.Descendants<InsertedRun>());
    }

    [Fact]
    public async Task Editing_by_id_never_puts_the_document_in_the_conversation()
    {
        using var session = new Session();
        var documentId = await session.Import(DocxFactory.Contract(), "contract.docx");

        // The whole point: the payloads an agent sees carry no content at all.
        var inspected = await session.Tools.InspectDocument("session", documentId);
        var applied = await session.Tools.ApplyPlan("session", documentId, """
            { "operations": [
                { "op": "changeText", "target": { "paraId": "w14:00000002", "expect": "Acme Corp", "occurrence": 0 },
                  "with": "Globex Inc.", "mode": "Direct" } ] }
            """);

        Assert.DoesNotContain("contentBase64", inspected);
        Assert.DoesNotContain("contentBase64", applied);
        // A .docx is a zip, so its base64 always starts UEsDB - a cheap check that no
        // payload smuggled one in.
        Assert.DoesNotContain("UEsDB", inspected);
        Assert.DoesNotContain("UEsDB", applied);
    }

    [Fact]
    public async Task A_removed_session_document_is_really_gone()
    {
        using var session = new Session();
        var documentId = await session.Import(DocxFactory.Contract(), "contract.docx");

        await session.Tools.RemoveDocument("session", documentId);

        var exported = await session.Tools.ExportDocumentContent("session", documentId);
        Assert.Contains("not-found", exported);
    }

    [Fact]
    public async Task Export_is_refused_for_a_connection_backed_by_real_storage()
    {
        // Exporting from storage would turn every readable document into base64 an agent
        // can quote, which is a different capability from editing one in place.
        var root = Path.Combine(Path.GetTempPath(), $"officeagent-eph-{Guid.NewGuid():N}");
        Directory.CreateDirectory(root);
        try
        {
            var services = new ServiceCollection();
            services.AddWordFormat();
            services.AddFileSystemDocumentProvider("workspace", root);
            services.AddMemoryDocumentProvider("session");
            services.AddOfficeAgent();
            using var provider = services.BuildServiceProvider();

            var client = provider.GetRequiredService<OfficeAgentClient>();
            var tools = new OfficeAgentTools(client);
            var document = await client.RegisterBytesAsync("workspace", root, DocxFactory.Contract(), "contract.docx");

            var refused = await tools.ExportDocumentContent("workspace", document.ItemId);
            Assert.Contains("invalid-argument", refused);
            Assert.Contains("not an in-memory connection", refused);
        }
        finally
        {
            try { Directory.Delete(root, recursive: true); } catch { }
        }
    }

    [Fact]
    public async Task A_session_connection_has_no_path_to_register_and_says_so()
    {
        using var session = new Session();

        var refused = await session.Tools.RegisterDocument("session", "contract.docx");
        Assert.Contains("invalid-argument", refused);
    }

    [Fact]
    public async Task The_session_connection_is_bounded_so_it_cannot_grow_without_limit()
    {
        // Room for one document and not two, whatever the fixture happens to weigh.
        var contract = DocxFactory.Contract();
        using var session = new Session(totalBytes: contract.Length + 512);

        var first = await session.Tools.ImportDocumentContent(
            "session", "a.docx", Convert.ToBase64String(contract));
        Assert.Contains("documentId", first);

        // The store is process memory, so its cap has to be a real one.
        var second = await session.Tools.ImportDocumentContent(
            "session", "b.docx", Convert.ToBase64String(contract));
        Assert.Contains("content-too-large", second);
    }

    [Fact]
    public async Task Session_import_rejects_malformed_packages_before_storing_them()
    {
        using var session = new Session();

        var refused = await session.Tools.ImportDocumentContent(
            "session", "broken.docx", Convert.ToBase64String(new byte[] { 1, 2, 3, 4 }));

        Assert.Contains("malformed-package", refused);
        Assert.DoesNotContain("documentId", refused);
    }

    [Fact]
    public async Task Session_import_uses_the_hosts_configured_compressed_ceiling()
    {
        var contract = DocxFactory.Contract();
        using var session = new Session(limits: new OpenXmlIngestionLimits
        {
            MaximumCompressedBytes = contract.LongLength - 1
        });

        var refused = await session.Tools.ImportDocumentContent(
            "session", "contract.docx", Convert.ToBase64String(contract));

        Assert.Contains("input-too-large", refused);
        Assert.Contains((contract.LongLength - 1).ToString(), refused);
        Assert.DoesNotContain("documentId", refused);
    }

    [Fact]
    public void A_deck_in_a_session_connection_is_created_like_any_other()
    {
        using var session = new Session(deck: true);

        var created = session.Client.CreateAsync("session", "review.pptx").GetAwaiter().GetResult();
        Assert.True(created.Committed);
        Assert.Equal("review.pptx", created.Document!.Name);
    }

    // ── MCP composition ──────────────────────────────────────────────────

    [Fact]
    public void A_session_connection_alone_is_enough_to_start_the_server()
    {
        var options = new OfficeAgentMcpOptions { EphemeralConnectionId = "session", AllowCreation = true };
        var names = OfficeAgentMcpServer.BuildToolset(options).Select(t => t.ProtocolTool.Name).ToArray();

        // The connection-addressed tools are offered, because there is now a connection
        // for them to name.
        Assert.Contains("inspect_document", names);
        Assert.Contains("apply_plan", names);
        Assert.Contains("create_document", names);
        Assert.Contains("import_document_content", names);
        Assert.Contains("export_document_content", names);
        Assert.Contains("list_connections", names);

        var connections = OfficeAgentMcpServer.ConnectionsPayload(options);
        Assert.Contains("\"connectionId\":\"session\"", connections);
        Assert.Contains("\"provider\":\"session\"", connections);
        Assert.Contains("\"canCreateDocuments\":true", connections);
    }

    [Fact]
    public void The_guidance_tells_an_agent_to_prefer_a_handle_over_the_bytes()
    {
        var instructions = OfficeAgentMcpServer.InstructionsFor(new OfficeAgentMcpOptions
        {
            EphemeralConnectionId = "session",
            AllowInlineContent = true
        });

        Assert.Contains("import_document_content", instructions);
        Assert.Contains("PREFER this over the _content tools", instructions);
        Assert.Contains("\"session\" (session)", instructions);
    }

    [Fact]
    public void The_ephemeral_tools_are_absent_without_a_session_connection()
    {
        var names = OfficeAgentMcpServer
            .BuildToolset(new OfficeAgentMcpOptions { AllowInlineContent = true })
            .Select(t => t.ProtocolTool.Name)
            .ToArray();

        Assert.DoesNotContain("import_document_content", names);
        Assert.DoesNotContain("export_document_content", names);
    }

    // ── helpers ──────────────────────────────────────────────────────────

    private static async Task<JsonElement> Json(Task<string> call)
    {
        using var parsed = JsonDocument.Parse(await call);
        return parsed.RootElement.Clone();
    }

    private sealed class Session : IDisposable
    {
        private readonly ServiceProvider _services;

        public OfficeAgentClient Client { get; }
        public OfficeAgentTools Tools { get; }

        public Session(
            long totalBytes = 100L * 1024 * 1024,
            bool deck = false,
            OpenXmlIngestionLimits? limits = null)
        {
            var services = new ServiceCollection();
            services.AddWordFormat();
            if (deck) services.AddPowerPointFormat();
            services.AddMemoryDocumentProvider("session", o => o.MaximumTotalBytes = totalBytes);
            if (limits is not null) services.AddSingleton(limits);
            services.AddOfficeAgent();
            _services = services.BuildServiceProvider();

            Client = _services.GetRequiredService<OfficeAgentClient>();
            Tools = new OfficeAgentTools(Client);
        }

        public async Task<string> Import(byte[] bytes, string name)
        {
            var imported = await Json(Tools.ImportDocumentContent("session", name, Convert.ToBase64String(bytes)));
            return imported.GetProperty("documentId").GetString()!;
        }

        public void Dispose() => _services.Dispose();
    }
}
