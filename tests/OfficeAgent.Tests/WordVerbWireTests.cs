using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.Abstractions;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// The new verbs as an agent actually sends them: JSON through the tools, not C# objects.
/// A verb that works against a hand-built <see cref="PlanOperation"/> and fails on the wire
/// - a string enum that does not parse, an anchor shape the converter reads as the wrong
/// type, a missing target - is a verb no agent can use.
/// </summary>
public class WordVerbWireTests
{
    [Fact]
    public async Task PageSetup_arrives_with_no_target_and_string_enums()
    {
        using var workspace = new Workspace();
        var document = await workspace.Register();

        await workspace.Apply(document, """
            { "operations": [
                { "op": "pageSetup",
                  "paperSize": "A4",
                  "orientation": "Landscape",
                  "marginTopTwips": 720,
                  "gutterTwips": 360 } ] }
            """);

        using var opened = await workspace.Open(document);
        var section = opened.Body.GetFirstChild<SectionProperties>()!;

        Assert.Equal(16838u, section.GetFirstChild<PageSize>()?.Width?.Value);
        Assert.Equal(PageOrientationValues.Landscape, section.GetFirstChild<PageSize>()?.Orient?.Value);
        Assert.Equal(720, section.GetFirstChild<PageMargin>()?.Top?.Value);
        Assert.Equal(360u, section.GetFirstChild<PageMargin>()?.Gutter?.Value);
    }

    [Fact]
    public async Task InsertBreak_parses_its_section_kinds_from_strings()
    {
        using var workspace = new Workspace();
        var document = await workspace.Register();

        await workspace.Apply(document, """
            { "operations": [
                { "op": "insertBreak",
                  "target": { "paraId": "w14:00000002", "expect": "Acme Corp" },
                  "kind": "SectionOddPage",
                  "position": "Before",
                  "mode": "Direct" } ] }
            """);

        using var opened = await workspace.Open(document);
        var inline = opened.Body.Descendants<Paragraph>()
            .Select(p => p.ParagraphProperties?.SectionProperties)
            .Single(s => s is not null)!;

        Assert.Equal(SectionMarkValues.OddPage, inline.GetFirstChild<SectionType>()?.Val?.Value);
    }

    [Fact]
    public async Task A_note_is_added_updated_and_removed_over_the_wire()
    {
        using var workspace = new Workspace();
        var document = await workspace.Register();

        await workspace.Apply(document, """
            { "operations": [
                { "op": "note",
                  "target": { "paraId": "w14:00000002", "expect": "Acme Corp", "occurrence": 0 },
                  "kind": "Footnote",
                  "text": "The counterparty, as defined in clause 1.",
                  "mode": "Direct" } ] }
            """);

        // The path an agent would read back out of inspect_document.
        var path = (await workspace.Nodes(document, "note")).Single().Path;
        Assert.Equal("footnote#1", path);

        await workspace.Apply(document, $$"""
            { "operations": [
                { "op": "note",
                  "target": { "kind": "note", "path": "{{path}}" },
                  "action": "Update",
                  "text": "As defined in clause 1.2.",
                  "mode": "Direct" } ] }
            """);

        using (var updated = await workspace.Open(document))
            Assert.Contains("As defined in clause 1.2.",
                updated.Package.MainDocumentPart!.FootnotesPart!.Footnotes!.InnerText);

        await workspace.Apply(document, $$"""
            { "operations": [
                { "op": "note",
                  "target": { "kind": "note", "path": "{{path}}" },
                  "action": "Remove",
                  "mode": "Direct" } ] }
            """);

        Assert.Empty(await workspace.Nodes(document, "note"));
    }

    [Fact]
    public async Task A_comment_is_replied_to_and_resolved_over_the_wire()
    {
        using var workspace = new Workspace();
        var document = await workspace.Register();

        await workspace.Apply(document, """
            { "operations": [
                { "op": "comment",
                  "target": { "paraId": "w14:00000002", "expect": "Acme Corp" },
                  "text": "Is this the right entity?",
                  "author": "Reviewer", "initials": "RV" } ] }
            """);

        await workspace.Apply(document, """
            { "operations": [
                { "op": "comment",
                  "target": { "kind": "comment", "path": "comment#1" },
                  "action": "Reply",
                  "text": "Yes - Acme Corp Ltd.",
                  "author": "Counsel", "initials": "CO" } ] }
            """);

        await workspace.Apply(document, """
            { "operations": [
                { "op": "comment",
                  "target": { "kind": "comment", "path": "comment#1" },
                  "action": "Resolve" } ] }
            """);

        var comments = await workspace.Nodes(document, "comment");
        Assert.Equal(2, comments.Count);
        Assert.Contains(comments, c => c.Path == "comment#1" && c.Summary.Contains("(resolved)"));
        Assert.Contains(comments, c => c.Path == "comment#2" && c.Summary.Contains("reply to comment#1"));
    }

    [Fact]
    public async Task Revisions_are_addressable_by_author_over_the_wire()
    {
        using var workspace = new Workspace();
        var document = await workspace.Register();

        // Omitting the mode means Tracked on this connection.
        await workspace.Apply(document, """
            { "operations": [
                { "op": "changeText",
                  "target": { "paraId": "w14:00000002", "expect": "Acme Corp" },
                  "with": "Globex Inc." } ] }
            """);

        var revisions = await workspace.Nodes(document, "revision");
        Assert.NotEmpty(revisions);
        Assert.All(revisions, r => Assert.Contains("by OfficeAgent", r.Summary));

        await workspace.Apply(document, """
            { "operations": [
                { "op": "revision",
                  "target": { "kind": "revision", "path": "author:OfficeAgent" },
                  "action": "Accept" } ] }
            """);

        Assert.Empty(await workspace.Nodes(document, "revision"));

        using var opened = await workspace.Open(document);
        // The clause names Acme Corp twice and the plan replaced the first; accepting keeps
        // that edit and leaves the occurrence nobody targeted alone.
        Assert.Contains("Globex Inc. shall provide services to Acme Corp.", opened.Body.InnerText);
        Assert.Empty(opened.Body.Descendants<DeletedText>());
    }

    [Fact]
    public async Task One_plan_can_reshape_a_document_end_to_end()
    {
        using var workspace = new Workspace();
        var document = await workspace.Register();

        // Everything a plan is allowed to do at once, all-or-nothing, over the wire: the
        // shape a real "turn this contract into the signed version" request takes.
        await workspace.Apply(document, """
            { "operations": [
                { "op": "pageSetup", "paperSize": "Letter", "marginLeftTwips": 1080 },
                { "op": "changeText",
                  "target": { "paraId": "w14:00000002", "expect": "Acme Corp", "occurrence": 0 },
                  "with": "Globex Inc." },
                { "op": "insert",
                  "target": { "paraId": "w14:00000004", "expect": "Effective date: 2020-01-01." },
                  "position": "After",
                  "text": "The Supplier shall indemnify the Client." },
                { "op": "note",
                  "target": { "paraId": "w14:00000001", "expect": "Service Agreement" },
                  "kind": "Footnote",
                  "text": "Executed in counterparts." },
                { "op": "comment",
                  "target": { "paraId": "w14:00000004", "expect": "Effective date" },
                  "text": "Confirm the commencement date." },
                { "op": "insertBreak",
                  "target": { "paraId": "w14:00000001", "expect": "Service Agreement" },
                  "kind": "Page", "position": "After" } ] }
            """);

        using var opened = await workspace.Open(document);
        AssertValid(opened.Package);

        var main = opened.Package.MainDocumentPart!;
        Assert.Contains("Globex Inc.", opened.Body.InnerText);
        Assert.Single(main.Document.Body!.Descendants<FootnoteReference>());
        Assert.Single(main.WordprocessingCommentsPart!.Comments!.Elements<Comment>());
        Assert.Contains(opened.Body.Descendants<Break>(), b => b.Type?.Value == BreakValues.Page);
        Assert.Equal(12240u, opened.Body.GetFirstChild<SectionProperties>()?.GetFirstChild<PageSize>()?.Width?.Value);

        // Text, paragraph, note and break all arrived as one round of redlines.
        var revisions = await workspace.Nodes(document, "revision");
        var tags = revisions.Select(r => r.Path.Split('#')[0]).ToHashSet();
        Assert.Contains("ins", tags);
        Assert.Contains("del", tags);
        Assert.Contains("markIns", tags);
    }

    [Fact]
    public async Task A_whole_redline_round_trips_through_accept()
    {
        using var workspace = new Workspace();
        var document = await workspace.Register();

        await workspace.Apply(document, """
            { "operations": [
                { "op": "changeText",
                  "target": { "paraId": "w14:00000002", "expect": "Acme Corp", "occurrence": 0 },
                  "with": "Globex Inc." },
                { "op": "insert",
                  "target": { "paraId": "w14:00000004", "expect": "Effective date: 2020-01-01." },
                  "position": "After",
                  "text": "The Supplier shall indemnify the Client." },
                { "op": "insertTable",
                  "target": { "paraId": "w14:00000004", "expect": "Effective date: 2020-01-01." },
                  "position": "After",
                  "table": { "headers": ["Milestone", "Date"], "rows": [["Kickoff", "2026-06-01"]] } },
                { "op": "note",
                  "target": { "paraId": "w14:00000001", "expect": "Service Agreement" },
                  "text": "Executed in counterparts." } ] }
            """);

        await workspace.Apply(document, """
            { "operations": [
                { "op": "revision",
                  "target": { "kind": "revision", "path": "all" },
                  "action": "Accept" } ] }
            """);

        using var opened = await workspace.Open(document);
        AssertValid(opened.Package);

        // Nothing pending, and every edit survived.
        Assert.Empty(await workspace.Nodes(document, "revision"));
        Assert.Contains("Globex Inc.", opened.Body.InnerText);
        Assert.Contains("The Supplier shall indemnify the Client.", opened.Body.InnerText);
        Assert.Single(opened.Body.Descendants<Table>());
        Assert.Equal(2, opened.Body.Descendants<Table>().Single().Elements<TableRow>().Count());
        Assert.Single(opened.Package.MainDocumentPart!.Document.Body!.Descendants<FootnoteReference>());
    }

    [Fact]
    public async Task A_whole_redline_round_trips_through_reject()
    {
        using var workspace = new Workspace();
        var document = await workspace.Register();
        var before = await workspace.Text(document);

        await workspace.Apply(document, """
            { "operations": [
                { "op": "changeText",
                  "target": { "paraId": "w14:00000002", "expect": "Acme Corp", "occurrence": 0 },
                  "with": "Globex Inc." },
                { "op": "insert",
                  "target": { "paraId": "w14:00000004", "expect": "Effective date: 2020-01-01." },
                  "position": "After",
                  "text": "The Supplier shall indemnify the Client." },
                { "op": "format",
                  "target": { "paraId": "w14:00000001", "expect": "Service Agreement" },
                  "bold": true } ] }
            """);

        await workspace.Apply(document, """
            { "operations": [
                { "op": "revision",
                  "target": { "kind": "revision", "path": "all" },
                  "action": "Reject" } ] }
            """);

        using var opened = await workspace.Open(document);
        AssertValid(opened.Package);

        // Rejecting everything is what "undo this agent's pass" means, so the text has to
        // come back exactly - not approximately.
        Assert.Empty(await workspace.Nodes(document, "revision"));
        Assert.Equal(before, await workspace.Text(document));
        Assert.DoesNotContain(opened.Body.Descendants<Run>(),
            r => r.RunProperties?.GetFirstChild<Bold>() is not null);
    }

    [Fact]
    public async Task Accepting_a_tracked_pass_lands_where_doing_it_directly_would()
    {
        // The invariant a redline rests on. If accept(tracked) and direct disagree, one of
        // the two paths is quietly producing a different document than the reviewer was
        // shown - which is worse than either being wrong on its own.
        const string edits = """
            { "op": "changeText",
              "target": { "paraId": "w14:00000002", "expect": "Acme Corp", "occurrence": 0 },
              "with": "Globex Inc.", "mode": "{{mode}}" },
            { "op": "insert",
              "target": { "paraId": "w14:00000004", "expect": "Effective date: 2020-01-01." },
              "position": "After",
              "text": "The Supplier shall indemnify the Client.", "mode": "{{mode}}" },
            { "op": "insertTable",
              "target": { "paraId": "w14:00000004", "expect": "Effective date: 2020-01-01." },
              "position": "After",
              "table": { "headers": ["Milestone", "Date"], "rows": [["Kickoff", "2026-06-01"]] },
              "mode": "{{mode}}" },
            { "op": "format",
              "target": { "paraId": "w14:00000001", "expect": "Service Agreement" },
              "bold": true, "mode": "{{mode}}" },
            { "op": "fill", "target": { "tag": "ClientName" }, "value": "Globex", "mode": "{{mode}}" }
            """;

        using var direct = new Workspace();
        var directDocument = await direct.Register();
        await direct.Apply(directDocument, $$"""{ "operations": [ {{edits.Replace("{{mode}}", "Direct")}} ] }""");

        using var tracked = new Workspace();
        var trackedDocument = await tracked.Register();
        await tracked.Apply(trackedDocument, $$"""{ "operations": [ {{edits.Replace("{{mode}}", "Tracked")}} ] }""");
        await tracked.Apply(trackedDocument, """
            { "operations": [
                { "op": "revision", "target": { "kind": "revision", "path": "all" }, "action": "Accept" } ] }
            """);

        Assert.Empty(await tracked.Nodes(trackedDocument, "revision"));
        Assert.Equal(await direct.Text(directDocument), await tracked.Text(trackedDocument));

        using var accepted = await tracked.Open(trackedDocument);
        AssertValid(accepted.Package);
    }

    [Fact]
    public async Task Revisions_outside_the_body_are_found_and_resolved_too()
    {
        using var workspace = new Workspace();
        var document = await workspace.Register();

        await workspace.Apply(document, """
            { "operations": [
                { "op": "headerFooter", "footer": "Confidential draft" },
                { "op": "note",
                  "target": { "paraId": "w14:00000002", "expect": "Acme Corp" },
                  "text": "The counterparty.", "mode": "Direct" } ] }
            """);

        var footerParaId = await workspace.ParaIdIn(document, "footer");
        var noteParaId = await workspace.ParaIdIn(document, "footnote");

        // A header, a footer and a note are text hosts like the body: a redline in one has
        // to be enumerated and resolvable, or "accept everything" quietly skips it.
        await workspace.Apply(document, $$"""
            { "operations": [
                { "op": "changeText",
                  "target": { "paraId": "{{footerParaId}}", "expect": "draft" },
                  "with": "final" },
                { "op": "changeText",
                  "target": { "paraId": "{{noteParaId}}", "expect": "counterparty" },
                  "with": "client" } ] }
            """);

        Assert.Equal(4, (await workspace.Nodes(document, "revision")).Count);

        await workspace.Apply(document, """
            { "operations": [
                { "op": "revision", "target": { "kind": "revision", "path": "all" }, "action": "Accept" } ] }
            """);

        Assert.Empty(await workspace.Nodes(document, "revision"));

        using var opened = await workspace.Open(document);
        AssertValid(opened.Package);
        Assert.Contains("Confidential final",
            opened.Package.MainDocumentPart!.FooterParts.Single().Footer!.InnerText);
        Assert.Contains("The client.",
            opened.Package.MainDocumentPart.FootnotesPart!.Footnotes!.InnerText);
    }

    [Fact]
    public async Task An_edit_inside_a_table_cell_is_redlined_like_any_other()
    {
        using var workspace = new Workspace();
        var document = await workspace.Register();

        await workspace.Apply(document, """
            { "operations": [
                { "op": "insertTable",
                  "target": { "paraId": "w14:00000004", "expect": "Effective date: 2020-01-01." },
                  "position": "After",
                  "table": { "headers": ["Milestone", "Date"], "rows": [["Kickoff", "2026-06-01"]] },
                  "mode": "Direct" } ] }
            """);

        var cellParaId = await workspace.ParaIdWithText(document, "2026-06-01");

        await workspace.Apply(document, $$"""
            { "operations": [
                { "op": "changeText",
                  "target": { "paraId": "{{cellParaId}}", "expect": "2026-06-01" },
                  "with": "2026-07-01" } ] }
            """);

        using var opened = await workspace.Open(document);
        AssertValid(opened.Package);

        var cell = opened.Body.Descendants<TableCell>().Single(c => c.InnerText.Contains("2026-07-01"));
        Assert.Single(cell.Descendants<InsertedRun>());
        Assert.Single(cell.Descendants<DeletedRun>());
    }

    [Fact]
    public async Task Applying_the_same_tracked_edit_twice_does_not_stack_markup()
    {
        using var workspace = new Workspace();
        var document = await workspace.Register();

        // The second pass finds the text it would have written already there, so it fails
        // on the anchor rather than layering a second revision over the first.
        await workspace.Apply(document, """
            { "operations": [
                { "op": "format",
                  "target": { "paraId": "w14:00000001", "expect": "Service Agreement" },
                  "bold": true } ] }
            """);

        await workspace.Apply(document, """
            { "operations": [
                { "op": "format",
                  "target": { "paraId": "w14:00000001", "expect": "Service Agreement" },
                  "bold": true } ] }
            """);

        using var opened = await workspace.Open(document);
        AssertValid(opened.Package);

        // One rPrChange, holding the state before the first pass - which is what "reject"
        // has to restore.
        var run = opened.Body.Descendants<Run>().Single(r => r.InnerText == "Service Agreement");
        Assert.Single(run.RunProperties!.Elements<RunPropertiesChange>());
        Assert.Null(run.RunProperties.GetFirstChild<RunPropertiesChange>()!
            .GetFirstChild<PreviousRunProperties>()?.GetFirstChild<Bold>());
    }

    [Fact]
    public async Task Nothing_is_written_when_one_operation_in_the_plan_fails()
    {
        using var workspace = new Workspace();
        var document = await workspace.Register();
        var before = await workspace.Text(document);

        var report = await workspace.Tools.ApplyPlan("workspace", document.ItemId, """
            { "operations": [
                { "op": "note",
                  "target": { "paraId": "w14:00000002", "expect": "Acme Corp" },
                  "text": "A good note." },
                { "op": "pageSetup", "paperSize": "Papyrus" } ] }
            """);

        using var parsed = JsonDocument.Parse(report);
        Assert.False(parsed.RootElement.GetProperty("committed").GetBoolean());

        Assert.Equal(before, await workspace.Text(document));
        Assert.Empty(await workspace.Nodes(document, "note"));
    }

    // ── helpers ──────────────────────────────────────────────────────────

    private static void AssertValid(WordprocessingDocument package)
    {
        var problems = new OpenXmlValidator(FileFormatVersions.Office2019)
            .Validate(package)
            .Select(e => $"{e.Path?.XPath}: {e.Description}")
            .ToList();
        Assert.True(problems.Count == 0, string.Join("; ", problems.Take(3)));
    }

    /// <summary>A filesystem-backed connection, which is how the tools are really reached.</summary>
    private sealed class Workspace : IDisposable
    {
        private readonly ServiceProvider _services;
        private readonly string _root;

        public OfficeAgentTools Tools { get; }
        public OfficeAgentClient Client { get; }

        public Workspace()
        {
            _root = Path.Combine(Path.GetTempPath(), $"officeagent-wire-{Guid.NewGuid():N}");
            Directory.CreateDirectory(_root);

            var services = new ServiceCollection();
            services.AddWordFormat();
            services.AddFileSystemDocumentProvider("workspace", _root);
            services.AddOfficeAgent();
            _services = services.BuildServiceProvider();

            Client = _services.GetRequiredService<OfficeAgentClient>();
            Tools = new OfficeAgentTools(Client);
        }

        public Task<DocumentReference> Register() =>
            Client.RegisterBytesAsync("workspace", _root, DocxFactory.Contract(), "contract.docx");

        public async Task Apply(DocumentReference document, string planJson)
        {
            var report = await Tools.ApplyPlan("workspace", document.ItemId, planJson);
            using var parsed = JsonDocument.Parse(report);
            Assert.True(parsed.RootElement.GetProperty("committed").GetBoolean(),
                parsed.RootElement.GetProperty("errors").ToString());
        }

        public async Task<IReadOnlyList<(string Path, string Summary)>> Nodes(
            DocumentReference document, string kind)
        {
            var json = await Tools.InspectDocument("workspace", document.ItemId);
            using var parsed = JsonDocument.Parse(json);

            return parsed.RootElement.GetProperty("nodes").EnumerateArray()
                .Where(n => n.GetProperty("Kind").GetString() == kind)
                .Select(n => (n.GetProperty("Path").GetString()!, n.GetProperty("Summary").GetString()!))
                .ToList();
        }

        /// <summary>The id of the first paragraph living in the named text host.</summary>
        public async Task<string> ParaIdIn(DocumentReference document, string location) =>
            (await Paragraphs(document))
                .First(p => p.GetProperty("location").GetString() == location)
                .GetProperty("ParaId").GetString()!;

        public async Task<string> ParaIdWithText(DocumentReference document, string text) =>
            (await Paragraphs(document))
                .First(p => p.GetProperty("Text").GetString()?.Contains(text) == true)
                .GetProperty("ParaId").GetString()!;

        private async Task<IReadOnlyList<JsonElement>> Paragraphs(DocumentReference document)
        {
            var json = await Tools.InspectDocument("workspace", document.ItemId);
            // Cloned out of the document before it is disposed.
            using var parsed = JsonDocument.Parse(json);
            return parsed.RootElement.GetProperty("paragraphs").EnumerateArray()
                .Select(p => p.Clone()).ToList();
        }

        public async Task<string> Text(DocumentReference document)
        {
            using var opened = await Open(document);
            return opened.Body.InnerText;
        }

        public async Task<OpenDocument> Open(DocumentReference document)
        {
            using var content = await Client.OpenReadAsync(
                DocumentReference.ForFileSystem("workspace", document.ItemId));
            var buffer = new MemoryStream();
            await content.Stream.CopyToAsync(buffer);
            buffer.Position = 0;
            return new OpenDocument(buffer);
        }

        public void Dispose()
        {
            _services.Dispose();
            if (Directory.Exists(_root)) Directory.Delete(_root, recursive: true);
        }
    }

    /// <summary>Keeps the package open for the life of the assertions against it.</summary>
    private sealed class OpenDocument : IDisposable
    {
        private readonly MemoryStream _stream;

        public WordprocessingDocument Package { get; }
        public Body Body { get; }

        public OpenDocument(MemoryStream stream)
        {
            _stream = stream;
            Package = WordprocessingDocument.Open(stream, isEditable: false);
            Body = Package.MainDocumentPart!.Document.Body!;
        }

        public void Dispose()
        {
            Package.Dispose();
            _stream.Dispose();
        }
    }
}
