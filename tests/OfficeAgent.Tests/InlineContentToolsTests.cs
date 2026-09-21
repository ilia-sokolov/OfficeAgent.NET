using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.PowerPoint;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// Creating and editing documents with no provider registered at all: the document arrives
/// as base64 and leaves the same way, and nothing is stored anywhere.
/// </summary>
public class InlineContentToolsTests
{
    /// <summary>A client with format modules and deliberately no providers.</summary>
    private static OfficeAgentTools Tools() =>
        new(new OfficeAgentClient(new WordModule(), new PowerPointModule()));

    // ── Creating from nothing ────────────────────────────────────────────

    [Fact]
    public async Task A_document_is_created_with_no_connection_in_existence()
    {
        var result = await Parse(Tools().CreateDocumentContent("report.docx"));

        Assert.True(result.Committed, result.Errors);
        Assert.Equal("report.docx", result.Name);

        var bytes = result.Content;
        Assert.NotNull(bytes);
        AssertValidWord(bytes!);
    }

    [Fact]
    public async Task An_initial_plan_authors_the_document_before_it_comes_back()
    {
        var result = await Parse(Tools().CreateDocumentContent("brief.docx", """
            { "operations": [
                { "op": "changeText",
                  "target": { "paraId": "auto-0000", "expect": "" },
                  "with": "Project Brief", "mode": "Direct" },
                { "op": "insert",
                  "target": { "paraId": "auto-0000", "expect": "" },
                  "position": "After",
                  "text": "Prepared for the steering group.", "mode": "Direct" } ] }
            """));

        Assert.True(result.Committed, result.Errors);
        using var body = Body(result.Content!);
        Assert.Contains("Project Brief", body.Root.InnerText);
        Assert.Contains("Prepared for the steering group.", body.Root.InnerText);
    }

    [Fact]
    public async Task A_deck_is_created_the_same_way()
    {
        var result = await Parse(Tools().CreateDocumentContent("review.pptx", """
            { "operations": [
                { "op": "changeText",
                  "target": { "paraId": "slide256/shape2/p0", "expect": "" },
                  "with": "Quarterly Review" } ] }
            """));

        Assert.True(result.Committed, result.Errors);

        using var stream = new MemoryStream(result.Content!);
        using var deck = PresentationDocument.Open(stream, isEditable: false);
        Assert.Single(deck.PresentationPart!.SlideParts);
    }

    [Fact]
    public async Task An_extension_no_module_can_mint_names_the_ones_that_exist()
    {
        var result = await Parse(Tools().CreateDocumentContent("ledger.xlsx"));

        Assert.False(result.Committed);
        Assert.Contains("invalid-argument", result.Errors);
        Assert.Contains(".docx", result.Errors);
    }

    [Fact]
    public async Task A_path_is_refused_as_a_document_name()
    {
        var result = await Parse(Tools().CreateDocumentContent("../escape.docx"));

        Assert.False(result.Committed);
        Assert.Contains("invalid-argument", result.Errors);
    }

    // ── Inspecting and editing content in hand ───────────────────────────

    [Fact]
    public async Task A_document_supplied_inline_is_inspected_like_any_other()
    {
        var json = await Tools().InspectDocumentContent(Base64(DocxFactory.Contract()));
        using var parsed = JsonDocument.Parse(json);

        Assert.Equal("Word", parsed.RootElement.GetProperty("format").GetString());
        Assert.Contains(
            parsed.RootElement.GetProperty("paragraphs").EnumerateArray(),
            p => p.GetProperty("text").GetString()!.Contains("shall provide services"));
    }

    [Fact]
    public async Task An_edit_returns_the_edited_document()
    {
        var result = await Parse(Tools().EditDocumentContent(Base64(DocxFactory.Contract()), """
            { "operations": [
                { "op": "changeText",
                  "target": { "paraId": "w14:00000002", "expect": "Acme Corp", "occurrence": 0 },
                  "with": "Globex Inc.", "mode": "Direct" } ] }
            """));

        Assert.True(result.Committed, result.Errors);
        using var body = Body(result.Content!);
        Assert.Contains("Globex Inc.", body.Root.InnerText);
    }

    [Fact]
    public async Task A_find_target_binds_without_inspecting_first()
    {
        // The point of the find target inline: an inspect round trip would cost another
        // full copy of the document in each direction.
        var result = await Parse(Tools().EditDocumentContent(Base64(DocxFactory.Contract()), """
            { "operations": [
                { "op": "changeText",
                  "target": { "find": "Effective date: 2020-01-01." },
                  "with": "Effective date: 2026-06-01.", "mode": "Direct" } ] }
            """));

        Assert.True(result.Committed, result.Errors);
        using var body = Body(result.Content!);
        Assert.Contains("2026-06-01", body.Root.InnerText);
    }

    [Fact]
    public async Task Ambiguous_text_is_refused_with_its_candidates()
    {
        var result = await Parse(Tools().EditDocumentContent(Base64(DocxFactory.Contract()), """
            { "operations": [
                { "op": "changeText",
                  "target": { "find": "Acme Corp" },
                  "with": "Globex Inc.", "mode": "Direct" } ] }
            """));

        Assert.False(result.Committed);
        Assert.Contains("ambiguous-anchor", result.Errors);
        Assert.Contains("match", result.Errors);
        Assert.Null(result.Content);
    }

    [Fact]
    public async Task A_preview_validates_without_producing_a_document()
    {
        var result = await Parse(Tools().EditDocumentContent(Base64(DocxFactory.Contract()), """
            { "operations": [
                { "op": "changeText",
                  "target": { "find": "Effective date: 2020-01-01." },
                  "with": "Effective date: 2026-06-01.", "mode": "Direct" } ] }
            """, preview: true));

        Assert.True(result.IsValid, result.Errors);
        Assert.False(result.Committed);
        // Nothing to hand back, and the agent can tell that from the payload alone.
        Assert.Null(result.Content);
    }

    [Fact]
    public async Task A_plan_that_fails_returns_no_document()
    {
        var result = await Parse(Tools().EditDocumentContent(Base64(DocxFactory.Contract()), """
            { "operations": [
                { "op": "changeText",
                  "target": { "paraId": "w14:00000002", "expect": "Not in the document" },
                  "with": "Anything", "mode": "Direct" } ] }
            """));

        Assert.False(result.Committed);
        Assert.Null(result.Content);
        Assert.Contains("expect-mismatch", result.Errors);
    }

    [Fact]
    public async Task Content_that_is_not_base64_names_the_fix()
    {
        var json = await Tools().InspectDocumentContent("this is not base64!!");
        Assert.Contains("invalid-argument", json);
        Assert.Contains("base64", json);
    }

    [Fact]
    public async Task Empty_content_is_refused_before_anything_is_decoded()
    {
        var json = await Tools().EditDocumentContent("", "{ \"operations\": [] }");
        Assert.Contains("invalid-argument", json);
    }

    [Theory]
    [InlineData("one character altered")]
    [InlineData("truncated")]
    [InlineData("not a package at all")]
    public async Task Content_that_did_not_arrive_intact_says_so(string damage)
    {
        // The likeliest failure of the whole inline design: the document has to survive
        // being reproduced in full on the way back in. This used to answer with
        // "internal-error - an unexpected internal error occurred", which gives an agent
        // nothing to act on and no reason to suspect its own copy of the content.
        var good = Convert.ToBase64String(DocxFactory.Contract());
        var damaged = damage switch
        {
            "one character altered" => Alter(good),
            "truncated" => good.Substring(0, good.Length - 400),
            _ => Convert.ToBase64String(new byte[] { 1, 2, 3, 4, 5, 6, 7, 8 })
        };

        var result = await Parse(Tools().EditDocumentContent(damaged, """
            { "operations": [
                { "op": "changeText", "target": { "find": "Acme Corp" }, "with": "Globex Inc." } ] }
            """));

        Assert.False(result.Committed);
        Assert.Null(result.Content);
        Assert.Contains("invalid-argument", result.Errors);
        Assert.Contains("not a readable", result.Errors);
        // Retrying the same string reproduces the same copy, so the error says not to.
        Assert.Contains("Do NOT re-send", result.Errors);
        Assert.DoesNotContain("internal-error", result.Errors);
    }

    [Fact]
    public async Task An_inline_failure_answers_in_the_shape_its_success_uses()
    {
        // A provider-shaped error here would name outputDocumentId and friends - which
        // mean nothing without a connection - while omitting the one field the tool told
        // the agent to read.
        using var parsed = JsonDocument.Parse(
            await Tools().EditDocumentContent(Convert.ToBase64String(DocxFactory.Contract()), "not json at all"));
        var root = parsed.RootElement;

        Assert.True(root.TryGetProperty("contentBase64", out var content));
        Assert.Equal(JsonValueKind.Null, content.ValueKind);
        Assert.False(root.GetProperty("committed").GetBoolean());
        Assert.Contains("invalid-json", root.GetProperty("errors").ToString());

        Assert.False(root.TryGetProperty("outputDocumentId", out _));
        Assert.False(root.TryGetProperty("sourceDocumentId", out _));
    }

    [Fact]
    public async Task A_plan_naming_a_verb_that_does_not_exist_lists_the_ones_that_do()
    {
        var result = await Parse(Tools().CreateDocumentContent("brief.docx", """
            { "operations": [ { "op": "insertParagraph", "target": { "paraId": "auto-0000" } } ] }
            """));

        Assert.False(result.Committed);
        Assert.Null(result.Content);
        Assert.Contains("insertParagraph", result.Errors);
        Assert.Contains("insert", result.Errors);
    }

    /// <summary>Changes one character in the middle, leaving valid base64 that is no longer a package.</summary>
    private static string Alter(string base64)
    {
        var middle = base64.Length / 2;
        var replacement = base64[middle] == 'A' ? 'B' : 'A';
        return base64.Substring(0, middle) + replacement + base64.Substring(middle + 1);
    }

    // ── The change mode a document with no connection inherits ───────────

    [Fact]
    public async Task A_word_document_edited_inline_still_defaults_to_tracked()
    {
        var result = await Parse(Tools().EditDocumentContent(Base64(DocxFactory.Contract()), """
            { "operations": [
                { "op": "changeText",
                  "target": { "paraId": "w14:00000002", "expect": "Acme Corp", "occurrence": 0 },
                  "with": "Globex Inc." } ] }
            """));

        Assert.True(result.Committed, result.Errors);
        using var body = Body(result.Content!);
        Assert.NotEmpty(body.Root.Descendants<InsertedRun>());
    }

    [Fact]
    public async Task A_deck_edited_inline_defaults_to_direct()
    {
        // There is no connection to carry the deck policy, and PresentationML cannot record
        // a revision at all - so the bytes decide, and the edit goes through rather than
        // failing on a mode the format can never honour.
        var deck = await Parse(Tools().CreateDocumentContent("deck.pptx"));

        var result = await Parse(Tools().EditDocumentContent(Base64(deck.Content!), """
            { "operations": [
                { "op": "changeText",
                  "target": { "paraId": "slide256/shape2/p0", "expect": "" },
                  "with": "Quarterly Review" } ] }
            """));

        Assert.True(result.Committed, result.Errors);
    }

    [Fact]
    public async Task A_deck_asked_explicitly_for_tracked_is_still_refused()
    {
        var deck = await Parse(Tools().CreateDocumentContent("deck.pptx"));

        var result = await Parse(Tools().EditDocumentContent(Base64(deck.Content!), """
            { "operations": [
                { "op": "changeText",
                  "target": { "paraId": "slide256/shape2/p0", "expect": "" },
                  "with": "Quarterly Review", "mode": "Tracked" } ] }
            """));

        Assert.False(result.Committed);
        Assert.Contains("no tracked-changes representation", result.Errors);
    }

    // ── The loop end to end ──────────────────────────────────────────────

    [Fact]
    public async Task Create_then_edit_then_edit_again_keeps_every_step()
    {
        var tools = Tools();

        var created = await Parse(tools.CreateDocumentContent("contract.docx", """
            { "operations": [
                { "op": "changeText",
                  "target": { "paraId": "auto-0000", "expect": "" },
                  "with": "Master Services Agreement", "mode": "Direct" } ] }
            """));

        var first = await Parse(tools.EditDocumentContent(Base64(created.Content!), """
            { "operations": [
                { "op": "insert",
                  "target": { "find": "Master Services Agreement" },
                  "position": "After",
                  "text": "The Supplier shall provide the Services.", "mode": "Direct" } ] }
            """));

        var second = await Parse(tools.EditDocumentContent(Base64(first.Content!), """
            { "operations": [
                { "op": "note",
                  "target": { "find": "the Services" },
                  "text": "As scoped in Schedule 1.", "mode": "Direct" } ] }
            """));

        Assert.True(second.Committed, second.Errors);
        AssertValidWord(second.Content!);

        using var stream = new MemoryStream(second.Content!);
        using var package = WordprocessingDocument.Open(stream, isEditable: false);
        var body = package.MainDocumentPart!.Document.Body!;

        Assert.Contains("Master Services Agreement", body.InnerText);
        Assert.Contains("The Supplier shall provide the Services.", body.InnerText);
        Assert.Single(body.Descendants<FootnoteReference>());
    }

    [Fact]
    public void The_inline_tools_carry_the_operation_vocabulary_themselves()
    {
        // An inline-only deployment has no preview_plan, so pointing at it for the shape of
        // an operation points at nothing. A live model, given only the earlier wording, sent
        // insert as { "with": { "text": … } }, got an empty paragraph, and concluded the verb
        // could not carry text at all.
        var tools = new OfficeAgentClient(new WordModule())
            .Tools()
            .AsAIFunctions(new OfficeAgentToolsOptions { AllowInlineContent = true });

        foreach (var name in new[] { "create_document_content", "edit_document_content" })
        {
            var description = tools.Single(t => t.Name == name).Description;
            Assert.Contains("\"op\": \"insert\"", description);
            Assert.Contains("\"text\": \"New paragraph.\"", description);
            Assert.Contains("\"op\": \"changeText\"", description);
            Assert.DoesNotContain("preview_plan documents", description);
        }
    }

    [Fact]
    public async Task The_inline_tools_are_absent_unless_the_host_asks_for_them()
    {
        var client = new OfficeAgentClient(new WordModule());

        var off = client.Tools().AsAIFunctions(new OfficeAgentToolsOptions()).Select(f => f.Name).ToList();
        Assert.DoesNotContain("create_document_content", off);
        Assert.DoesNotContain("edit_document_content", off);

        var on = client.Tools()
            .AsAIFunctions(new OfficeAgentToolsOptions { AllowInlineContent = true })
            .Select(f => f.Name).ToList();
        Assert.Contains("create_document_content", on);
        Assert.Contains("inspect_document_content", on);
        Assert.Contains("edit_document_content", on);
    }

    // ── helpers ──────────────────────────────────────────────────────────

    private static string Base64(byte[] document) => Convert.ToBase64String(document);

    /// <summary>The fields of an inline result an assertion cares about.</summary>
    private readonly record struct Result(bool IsValid, bool Committed, string? Name, byte[]? Content, string Errors);

    private static async Task<Result> Parse(Task<string> call)
    {
        using var parsed = JsonDocument.Parse(await call);
        var root = parsed.RootElement;

        var content = root.TryGetProperty("contentBase64", out var encoded) &&
                      encoded.ValueKind == JsonValueKind.String
            ? Convert.FromBase64String(encoded.GetString()!)
            : null;

        return new Result(
            root.TryGetProperty("isValid", out var valid) && valid.GetBoolean(),
            root.TryGetProperty("committed", out var committed) && committed.GetBoolean(),
            root.TryGetProperty("name", out var name) && name.ValueKind == JsonValueKind.String
                ? name.GetString()
                : null,
            content,
            root.TryGetProperty("errors", out var errors) ? errors.ToString() : root.ToString());
    }

    private sealed class OpenBody : IDisposable
    {
        private readonly MemoryStream _stream;
        private readonly WordprocessingDocument _document;

        public Body Root { get; }

        public OpenBody(byte[] bytes)
        {
            _stream = new MemoryStream(bytes);
            _document = WordprocessingDocument.Open(_stream, isEditable: false);
            Root = _document.MainDocumentPart!.Document.Body!;
        }

        public void Dispose()
        {
            _document.Dispose();
            _stream.Dispose();
        }
    }

    private static OpenBody Body(byte[] document) => new(document);

    private static void AssertValidWord(byte[] document)
    {
        using var stream = new MemoryStream(document);
        using var opened = WordprocessingDocument.Open(stream, isEditable: false);
        var problems = new OpenXmlValidator(FileFormatVersions.Office2019)
            .Validate(opened)
            .Select(e => $"{e.Path?.XPath}: {e.Description}")
            .ToList();
        Assert.True(problems.Count == 0, string.Join("; ", problems.Take(3)));
    }
}

internal static class InlineToolsClientExtensions
{
    internal static OfficeAgentTools Tools(this OfficeAgentClient client) => new(client);
}
