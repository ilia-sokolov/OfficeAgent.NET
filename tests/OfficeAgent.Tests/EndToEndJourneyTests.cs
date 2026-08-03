using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.PowerPoint;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// End-to-end journeys driven the way a real agent drives them: through the JSON tool
/// surface, one call at a time, reading each result to decide the next call. Nothing here
/// touches an internal type - if a journey passes, an LLM following the tool descriptions
/// can complete that task, and the file it produces opens in Office.
/// </summary>
/// <remarks>
/// These complement the per-verb tests, which prove a verb writes correct XML. A journey
/// proves the verbs compose: that ids returned by one call are accepted by the next, that
/// anchors survive a save, and that the document is still valid at the end.
/// </remarks>
public class EndToEndJourneyTests
{
    // ── Word: the original use case, still working ────────────────────────────

    [Fact]
    public async Task Journey_edit_a_contract_through_the_full_tool_loop()
    {
        using var host = new AgentHost();
        var tools = host.Tools;
        host.Stage("contract.docx", DocxFactory.Contract());

        // 1. The user names a file; the agent stages it.
        var documentId = Field(await tools.RegisterDocument("workspace", "contract.docx"), "documentId");

        // 2. Inspect to understand the document and capture the snapshot.
        using var inspected = Json(await tools.InspectDocument("workspace", documentId));
        Assert.Equal("Word", inspected.RootElement.GetProperty("format").GetString());
        Assert.True(inspected.RootElement.GetProperty("paragraphsTotal").GetInt32() > 0);

        // 3. Find the text the user asked about - anchors must come from the engine.
        using var hits = Json(await tools.FindInDocument("workspace", documentId, "Acme Corp"));
        var hit = hits.RootElement[0];
        var paraId = hit.GetProperty("paraId").GetString()!;

        var plan = $$"""
            { "operations": [ { "op": "changeText",
                                "target": { "paraId": "{{paraId}}", "expect": "Acme Corp", "occurrence": 0 },
                                "with": "Globex Inc.", "mode": "Direct" } ] }
            """;

        // 4. Preview before writing - the safety step the guidance insists on.
        using var preview = Json(await tools.PreviewPlan("workspace", documentId, plan));
        Assert.True(preview.RootElement.GetProperty("isValid").GetBoolean());
        Assert.False(preview.RootElement.GetProperty("committed").GetBoolean());

        // 5. Apply, then hand the new id onward.
        using var applied = Json(await tools.ApplyPlan("workspace", documentId, plan));
        Assert.True(applied.RootElement.GetProperty("committed").GetBoolean());
        var outputId = applied.RootElement.GetProperty("outputDocumentId").GetString()!;
        Assert.Equal(
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            applied.RootElement.GetProperty("outputContentType").GetString());

        // 6. The saved revision really contains the edit, and the source is untouched.
        Assert.Contains("Globex Inc.", await TextOf(tools, outputId));
        Assert.Contains("Acme Corp", await TextOf(tools, documentId));

        // 7. The host can fetch the bytes to deliver, and they open cleanly.
        AssertOpensAsWord(await host.BytesOf(outputId));
    }

    [Fact]
    public async Task Journey_draft_a_new_document_and_keep_editing_it()
    {
        using var host = new AgentHost();
        var tools = host.Tools;

        // "Draft a project brief" - nothing exists yet.
        var initial = """
            [ { "op": "insert",
                "target": { "paraId": "auto-0000", "expect": "" },
                "position": "Before", "text": "Project Brief", "styleId": "Heading1" } ]
            """;
        using var created = Json(await tools.CreateDocument("workspace", "brief.docx", initial));
        Assert.True(created.RootElement.GetProperty("committed").GetBoolean(),
            created.RootElement.GetProperty("errors").ToString());
        var documentId = created.RootElement.GetProperty("outputDocumentId").GetString()!;

        // The agent continues from the id it was handed, with anchors it re-reads.
        using var hits = Json(await tools.FindInDocument("workspace", documentId, "Project Brief"));
        var paraId = hits.RootElement[0].GetProperty("paraId").GetString()!;

        var follow = $$"""
            [ { "op": "insert",
                "target": { "paraId": "{{paraId}}", "expect": "Project Brief" },
                "position": "After", "text": "Scope: migrate the billing service." } ]
            """;
        using var applied = Json(await tools.ApplyPlan("workspace", documentId, follow));
        Assert.True(applied.RootElement.GetProperty("committed").GetBoolean(),
            applied.RootElement.GetProperty("errors").ToString());

        var text = await TextOf(tools, applied.RootElement.GetProperty("outputDocumentId").GetString()!);
        Assert.Contains("Project Brief", text);
        Assert.Contains("migrate the billing service", text);
    }

    [Fact]
    public async Task Journey_one_call_edit_by_path_with_a_find_target()
    {
        using var host = new AgentHost();
        var tools = host.Tools;
        host.Stage("contract.docx", DocxFactory.Contract());

        // The composite path: no register, no find, no preview - one call.
        using var edited = Json(await tools.EditDocument("workspace", "contract.docx",
            """
            [ { "op": "changeText", "target": { "find": "Service Agreement" },
                "with": "Master Agreement", "mode": "Direct" } ]
            """));

        Assert.True(edited.RootElement.GetProperty("committed").GetBoolean(),
            edited.RootElement.GetProperty("errors").ToString());
        Assert.Contains("Master Agreement",
            await TextOf(tools, edited.RootElement.GetProperty("outputDocumentId").GetString()!));
    }

    // ── PowerPoint: the new use case ──────────────────────────────────────────

    [Fact]
    public async Task Journey_build_a_deck_from_nothing()
    {
        using var host = new AgentHost();
        var tools = host.Tools;

        // 1. "Make me a quarterly review deck."
        using var created = Json(await tools.CreateDocument("workspace", "review.pptx", ""));
        Assert.True(created.RootElement.GetProperty("committed").GetBoolean(),
            created.RootElement.GetProperty("errors").ToString());
        var deckId = created.RootElement.GetProperty("outputDocumentId").GetString()!;
        Assert.Equal(
            "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            created.RootElement.GetProperty("outputContentType").GetString());

        // 2. Inspect to learn the deck's shape - the agent must not invent paths.
        using var inspected = Json(await tools.InspectDocument("workspace", deckId));
        Assert.Equal("PowerPoint", inspected.RootElement.GetProperty("format").GetString());
        var slidePath = inspected.RootElement.GetProperty("nodes").EnumerateArray()
            .First(n => n.GetProperty("Kind").GetString() == "slide")
            .GetProperty("Path").GetString()!;
        var titleParaId = inspected.RootElement.GetProperty("paragraphs")[0].GetProperty("ParaId").GetString()!;

        // 3. Title the slide, add a table and a picture in one plan.
        var build = $$"""
            [ { "op": "changeText",
                "target": { "paraId": "{{titleParaId}}", "expect": "" },
                "with": "Quarterly Review", "mode": "Direct" },
              { "op": "insertTable",
                "target": { "kind": "slide", "path": "{{slidePath}}" },
                "table": { "headers": ["Region", "Q1"], "rows": [["EMEA", "41850"], ["APAC", "22100"]] } },
              { "op": "insertImage",
                "target": { "kind": "slide", "path": "{{slidePath}}" },
                "base64Bytes": "{{OnePixelPng}}", "imageType": "png",
                "widthPx": 240, "heightPx": 120, "altText": "Revenue chart" } ]
            """;

        using var built = Json(await tools.ApplyPlan("workspace", deckId, build));
        Assert.True(built.RootElement.GetProperty("committed").GetBoolean(),
            built.RootElement.GetProperty("errors").ToString());
        Assert.Equal(3, built.RootElement.GetProperty("changes").GetArrayLength());

        // 4. Everything the agent added is addressable afterwards.
        var finalId = built.RootElement.GetProperty("outputDocumentId").GetString()!;
        using var final = Json(await tools.InspectDocument("workspace", finalId));
        var kinds = final.RootElement.GetProperty("nodes").EnumerateArray()
            .Select(n => n.GetProperty("Kind").GetString()).ToList();
        Assert.Contains("table", kinds);
        Assert.Contains("image", kinds);
        Assert.Contains("Quarterly Review", await TextOf(tools, finalId));
        Assert.Contains("EMEA", await TextOf(tools, finalId));

        // 5. And PowerPoint can open the result.
        AssertOpensAsPresentation(await host.BytesOf(finalId));
    }

    [Fact]
    public async Task Journey_review_a_deck_then_resolve_the_comment()
    {
        using var host = new AgentHost();
        var tools = host.Tools;
        host.Stage("deck.pptx", PptxFactory.Deck());

        // 1. Open by path: register + inspect together.
        using var opened = Json(await tools.OpenDocument("workspace", "deck.pptx"));
        var deckId = opened.RootElement.GetProperty("documentId").GetString()!;
        var slidePath = opened.RootElement.GetProperty("nodes").EnumerateArray()
            .First(n => n.GetProperty("Kind").GetString() == "slide")
            .GetProperty("Path").GetString()!;

        // 2. Leave review feedback.
        using var commented = Json(await tools.ApplyPlan("workspace", deckId, $$"""
            [ { "op": "comment", "target": { "kind": "slide", "path": "{{slidePath}}" },
                "text": "Confirm the EMEA figure.", "author": "Reviewer", "initials": "RV" } ]
            """));
        Assert.True(commented.RootElement.GetProperty("committed").GetBoolean(),
            commented.RootElement.GetProperty("errors").ToString());
        var reviewedId = commented.RootElement.GetProperty("outputDocumentId").GetString()!;

        // 3. Later, the comment is discoverable with its own path.
        using var reviewed = Json(await tools.InspectDocument("workspace", reviewedId));
        var comment = reviewed.RootElement.GetProperty("nodes").EnumerateArray()
            .Single(n => n.GetProperty("Kind").GetString() == "comment");
        Assert.Contains("Reviewer", comment.GetProperty("Summary").GetString());
        Assert.DoesNotContain("(resolved)", comment.GetProperty("Summary").GetString());

        // 4. Resolve it - the text survives, only the status changes.
        using var resolved = Json(await tools.ApplyPlan("workspace", reviewedId, $$"""
            [ { "op": "comment", "action": "Resolve",
                "target": { "kind": "comment", "path": "{{comment.GetProperty("Path").GetString()}}" } } ]
            """));
        Assert.True(resolved.RootElement.GetProperty("committed").GetBoolean(),
            resolved.RootElement.GetProperty("errors").ToString());

        using var after = Json(await tools.InspectDocument(
            "workspace", resolved.RootElement.GetProperty("outputDocumentId").GetString()!));
        var settled = after.RootElement.GetProperty("nodes").EnumerateArray()
            .Single(n => n.GetProperty("Kind").GetString() == "comment");
        Assert.Contains("(resolved)", settled.GetProperty("Summary").GetString());
        Assert.Contains("Confirm the EMEA figure", settled.GetProperty("Summary").GetString());
    }

    [Fact]
    public async Task Journey_update_the_numbers_in_an_existing_deck()
    {
        using var host = new AgentHost();
        var tools = host.Tools;
        host.Stage("deck.pptx", PptxFactory.DeckWithTable());

        // Editing a slide table cell is ordinary text editing, via find.
        using var edited = Json(await tools.EditDocument("workspace", "deck.pptx",
            """
            [ { "op": "changeText", "target": { "find": "41850" },
                "with": "44120", "mode": "Direct" } ]
            """));

        Assert.True(edited.RootElement.GetProperty("committed").GetBoolean(),
            edited.RootElement.GetProperty("errors").ToString());

        var outputId = edited.RootElement.GetProperty("outputDocumentId").GetString()!;
        Assert.Contains("44120", await TextOf(tools, outputId));
        AssertOpensAsPresentation(await host.BytesOf(outputId));
    }

    // ── Cross-cutting behaviour a user actually feels ─────────────────────────

    [Fact]
    public async Task One_host_serves_both_formats_from_the_same_connection()
    {
        using var host = new AgentHost();
        var tools = host.Tools;
        host.Stage("contract.docx", DocxFactory.Contract());
        host.Stage("deck.pptx", PptxFactory.Deck());

        using var word = Json(await tools.OpenDocument("workspace", "contract.docx"));
        using var deck = Json(await tools.OpenDocument("workspace", "deck.pptx"));

        Assert.Equal("Word", word.RootElement.GetProperty("format").GetString());
        Assert.Equal("PowerPoint", deck.RootElement.GetProperty("format").GetString());
    }

    [Fact]
    public async Task A_deck_edit_that_asks_for_tracked_changes_is_told_what_to_do_instead()
    {
        using var host = new AgentHost();
        var tools = host.Tools;
        host.Stage("deck.pptx", PptxFactory.Deck());

        // "Tracked" is the guidance's default for Word. An agent carrying that habit to a
        // deck must get a message it can act on, not a bare failure.
        using var refused = Json(await tools.EditDocument("workspace", "deck.pptx",
            """
            [ { "op": "changeText", "target": { "find": "Quarterly Review" },
                "with": "Annual Review", "mode": "Tracked" } ]
            """));

        Assert.False(refused.RootElement.GetProperty("committed").GetBoolean());
        var message = refused.RootElement.GetProperty("errors")[0].GetProperty("Message").GetString()!;
        Assert.Contains("Direct", message);
    }

    [Fact]
    public async Task An_ambiguous_edit_names_its_candidates_and_writes_nothing()
    {
        using var host = new AgentHost();
        var tools = host.Tools;
        host.Stage("contract.docx", DocxFactory.Contract());
        var before = File.ReadAllBytes(Path.Combine(host.Root, "contract.docx"));

        // The fixture says "Acme Corp" twice.
        using var refused = Json(await tools.EditDocument("workspace", "contract.docx",
            """
            [ { "op": "changeText", "target": { "find": "Acme Corp" }, "with": "Globex" } ]
            """));

        Assert.Equal("ambiguous-anchor",
            refused.RootElement.GetProperty("errors")[0].GetProperty("Code").GetString());
        Assert.Equal(before, File.ReadAllBytes(Path.Combine(host.Root, "contract.docx")));

        // …and the retry the error tells the agent to make actually works.
        using var retried = Json(await tools.EditDocument("workspace", "contract.docx",
            """
            [ { "op": "changeText", "target": { "find": "Acme Corp", "match": 1 },
                "with": "Globex", "mode": "Direct" } ]
            """));
        Assert.True(retried.RootElement.GetProperty("committed").GetBoolean(),
            retried.RootElement.GetProperty("errors").ToString());
    }

    [Fact]
    public async Task A_verb_a_deck_cannot_do_fails_the_whole_plan_and_writes_nothing()
    {
        using var host = new AgentHost();
        var tools = host.Tools;
        host.Stage("deck.pptx", PptxFactory.Deck());
        var documentId = Field(await tools.RegisterDocument("workspace", "deck.pptx"), "documentId");

        using var inspected = Json(await tools.InspectDocument("workspace", documentId));
        var slidePath = inspected.RootElement.GetProperty("nodes").EnumerateArray()
            .First(n => n.GetProperty("Kind").GetString() == "slide").GetProperty("Path").GetString()!;
        var paraId = inspected.RootElement.GetProperty("paragraphs")[0].GetProperty("ParaId").GetString()!;

        // A valid table insert paired with a verb PowerPoint does not implement: the good
        // half must not land on its own.
        using var refused = Json(await tools.ApplyPlan("workspace", documentId, $$"""
            [ { "op": "insertTable", "target": { "kind": "slide", "path": "{{slidePath}}" },
                "table": { "headers": ["A"], "rows": [["1"]] } },
              { "op": "setProperty", "target": { "kind": "docProperty", "path": "core/title" },
                "value": "Nope" } ]
            """));

        Assert.False(refused.RootElement.GetProperty("committed").GetBoolean());
        using var unchanged = Json(await tools.InspectDocument("workspace", documentId));
        Assert.DoesNotContain(unchanged.RootElement.GetProperty("nodes").EnumerateArray(),
            n => n.GetProperty("Kind").GetString() == "table");
    }

    // ── helpers ───────────────────────────────────────────────────────────────

    private const string OnePixelPng =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";

    private static JsonDocument Json(string payload) => JsonDocument.Parse(payload);

    private static string Field(string payload, string name)
    {
        using var document = JsonDocument.Parse(payload);
        var value = document.RootElement.TryGetProperty(name, out var property)
            ? property.GetString()
            : null;
        Assert.False(string.IsNullOrEmpty(value),
            $"expected '{name}' in tool result but got: {payload}");
        return value!;
    }

    private static async Task<string> TextOf(OfficeAgentTools tools, string documentId)
    {
        using var inspected = JsonDocument.Parse(
            await tools.InspectDocument("workspace", documentId, paragraphLimit: 1000));
        return string.Join(" ", inspected.RootElement.GetProperty("paragraphs").EnumerateArray()
            .Select(p => p.GetProperty("Text").GetString()));
    }

    private static void AssertOpensAsWord(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, isEditable: false);
        AssertNoValidationErrors(new OpenXmlValidator(FileFormatVersions.Office2019).Validate(document));
    }

    private static void AssertOpensAsPresentation(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = PresentationDocument.Open(stream, isEditable: false);
        AssertNoValidationErrors(new OpenXmlValidator(FileFormatVersions.Office2019).Validate(document));
    }

    private static void AssertNoValidationErrors(IEnumerable<ValidationErrorInfo> errors)
    {
        var found = errors.ToList();
        Assert.True(found.Count == 0,
            string.Join("; ", found.Select(e => $"{e.Path?.XPath}: {e.Description}")));
    }

    /// <summary>
    /// A host wired the way the docs tell a host to wire it: both format modules, one
    /// filesystem connection, tools over the resulting client.
    /// </summary>
    private sealed class AgentHost : IDisposable
    {
        private readonly ServiceProvider _services;

        public AgentHost()
        {
            Root = Path.Combine(Path.GetTempPath(), $"officeagent-e2e-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Root);

            var services = new ServiceCollection();
            services.AddWordFormat();
            services.AddPowerPointFormat();
            services.AddFileSystemDocumentProvider("workspace", Root, o =>
                o.AllowedExtensions = new[] { ".docx", ".pptx" });
            services.AddOfficeAgent();

            _services = services.BuildServiceProvider();
            Client = _services.GetRequiredService<OfficeAgentClient>();
            Tools = new OfficeAgentTools(Client);
        }

        public string Root { get; }
        public OfficeAgentClient Client { get; }
        public OfficeAgentTools Tools { get; }

        public void Stage(string name, byte[] bytes) =>
            File.WriteAllBytes(Path.Combine(Root, name), bytes);

        /// <summary>What the host does to deliver a result: fetch the bytes by id.</summary>
        public async Task<byte[]> BytesOf(string documentId)
        {
            using var content = await Client.OpenReadAsync(
                OfficeAgent.Abstractions.DocumentReference.ForFileSystem("workspace", documentId));
            using var buffer = new MemoryStream();
            await content.Stream.CopyToAsync(buffer);
            return buffer.ToArray();
        }

        public void Dispose()
        {
            _services.Dispose();
            if (Directory.Exists(Root)) Directory.Delete(Root, recursive: true);
        }
    }
}
