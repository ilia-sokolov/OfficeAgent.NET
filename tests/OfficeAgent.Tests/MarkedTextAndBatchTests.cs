using System.Security.Cryptography;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.Abstractions;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.PowerPoint;
using OfficeAgent.Word;
using A = DocumentFormat.OpenXml.Drawing;
using DocFormat = OfficeAgent.Abstractions.DocumentFormat;

namespace OfficeAgent.Tests;

/// <summary>
/// Translating a paragraph keeps its character formatting through markedText and withMarked,
/// and one plan edits many documents through "match": "all" and edit_documents.
/// </summary>
public class MarkedTextAndBatchTests
{
    private const string Payment = "Payment is due within 30 days of invoice.";

    // ── markedText ───────────────────────────────────────────────────────────

    [Fact]
    public void Inspection_marks_the_words_whose_formatting_differs_from_the_paragraph()
    {
        var inspect = Word().Inspect(PaymentMemo());

        var payment = inspect.Paragraphs.Single(p => p.Text == Payment);
        Assert.Equal("Payment is due within <1>30 days</1> of invoice.", payment.MarkedText);

        // One formatting throughout needs no tags.
        Assert.Null(inspect.Paragraphs.Single(p => p.Text == "Plain closing line.").MarkedText);
    }

    [Fact]
    public void The_base_formatting_is_the_one_most_of_the_paragraph_has_and_tags_repeat_per_formatting()
    {
        var bytes = Memo(Para("0000000A",
            Run("Mostly bold text ", bold: true), Run("then plain "), Run("and bold again", bold: true), Run(" and "),
            Run("italic", italic: true)));

        var paragraph = Assert.Single(Word().Inspect(bytes).Paragraphs);

        Assert.Equal("Mostly bold text <1>then plain </1>and bold again<1> and </1><2>italic</2>", paragraph.MarkedText);
    }

    [Fact]
    public void A_paragraph_crossing_a_content_control_has_no_marked_text_and_refuses_a_marked_rewrite()
    {
        var office = Word();
        var bytes = DocxFactory.Contract();
        var client = office.Inspect(bytes).Paragraphs.Single(p => p.Text.StartsWith("Client:", StringComparison.Ordinal));
        Assert.Null(client.MarkedText);

        using var result = office.Commit(new StreamHandle(new MemoryStream(bytes)), Plan(new ChangeTextOp
        {
            Target = new TextSpanAnchor { ParaId = client.ParaId, Expect = client.Text },
            WithMarked = "Klant: PLACEHOLDER",
            Mode = ChangeMode.Direct
        }));

        var error = Assert.Single(result.Report.Errors);
        Assert.Equal(ValidationErrorCodes.InvalidOperation, error.Code);
        Assert.Contains("content control", error.Message);
    }

    // ── withMarked in Word ───────────────────────────────────────────────────

    [Fact]
    public void A_marked_rewrite_moves_the_formatting_with_the_words_wherever_they_land()
    {
        var office = Word();
        var bytes = PaymentMemo();
        var paragraph = office.Inspect(bytes).Paragraphs.Single(p => p.Text == Payment);

        using var result = office.Commit(new StreamHandle(new MemoryStream(bytes)), Plan(new ChangeTextOp
        {
            Target = new TextSpanAnchor { ParaId = paragraph.ParaId, Expect = Payment },
            WithMarked = "<1>Binnen 30 dagen</1> na factuurdatum moet worden betaald.",
            Mode = ChangeMode.Direct
        }));

        Assert.True(result.Committed, Errors(result));
        var runs = FirstParagraphRuns(result.ToBytes());
        Assert.Equal(new[] { ("Binnen 30 dagen", true), (" na factuurdatum moet worden betaald.", false) }, runs);
        Assert.Equal("Binnen 30 dagen na factuurdatum moet worden betaald.", Assert.Single(result.Report.Changes).After);
        Assert.Empty(Validate(result.ToBytes()));

        var after = office.Inspect(result.ToBytes()).Paragraphs.Single(p => p.ParaId == paragraph.ParaId);
        Assert.Equal("<1>Binnen 30 dagen</1> na factuurdatum moet worden betaald.", after.MarkedText);
    }

    [Fact]
    public void A_tracked_marked_rewrite_is_one_deletion_and_one_insertion_carrying_the_formatting()
    {
        var office = Word();
        var bytes = PaymentMemo();
        var paragraph = office.Inspect(bytes).Paragraphs.Single(p => p.Text == Payment);

        using var result = office.Commit(new StreamHandle(new MemoryStream(bytes)), Plan(new ChangeTextOp
        {
            Target = new TextSpanAnchor { ParaId = paragraph.ParaId, Expect = Payment },
            WithMarked = "Zahlung innerhalb von <1>30 Tagen</1> nach Rechnungsstellung.",
            Mode = ChangeMode.Tracked
        }));

        Assert.True(result.Committed, Errors(result));
        using var stream = new MemoryStream(result.ToBytes());
        using var document = WordprocessingDocument.Open(stream, isEditable: false);
        var body = document.MainDocumentPart!.Document.Body!.Elements<Paragraph>().First();

        var deleted = Assert.Single(body.Elements<DeletedRun>());
        Assert.Equal(Payment, string.Concat(deleted.Descendants<DeletedText>().Select(t => t.Text)));
        var inserted = Assert.Single(body.Elements<InsertedRun>());
        Assert.Equal(
            new[] { ("Zahlung innerhalb von ", false), ("30 Tagen", true), (" nach Rechnungsstellung.", false) },
            inserted.Elements<Run>().Select(r => (Text(r), r.RunProperties?.Bold is not null)));
        Assert.Empty(Validate(result.ToBytes()));
    }

    [Theory]
    [InlineData("Payment is due within 30 days", "Betaling", "complete text")]
    [InlineData(Payment, "Betaling <3>binnen</3>", "does not have")]
    [InlineData(Payment, "Betaling <1>binnen <1>30</1></1>", "do not nest")]
    [InlineData(Payment, "Betaling <1>binnen", "never closes")]
    [InlineData(Payment, "Betaling </1>", "without opening")]
    public void A_marked_rewrite_that_does_not_fit_the_paragraph_is_refused_naming_why(
        string expect, string withMarked, string reason)
    {
        var office = Word();
        var bytes = PaymentMemo();
        var paragraph = office.Inspect(bytes).Paragraphs.Single(p => p.Text == Payment);

        var report = office.Preview(new StreamHandle(new MemoryStream(bytes)), Plan(new ChangeTextOp
        {
            Target = new TextSpanAnchor { ParaId = paragraph.ParaId, Expect = expect },
            WithMarked = withMarked,
            Mode = ChangeMode.Direct
        }));

        var error = Assert.Single(report.Errors);
        Assert.Equal(ValidationErrorCodes.InvalidOperation, error.Code);
        Assert.Contains(reason, error.Message);
    }

    [Fact]
    public void With_and_with_marked_together_are_refused()
    {
        var office = Word();
        var bytes = PaymentMemo();
        var paragraph = office.Inspect(bytes).Paragraphs.Single(p => p.Text == Payment);

        var report = office.Preview(new StreamHandle(new MemoryStream(bytes)), Plan(new ChangeTextOp
        {
            Target = new TextSpanAnchor { ParaId = paragraph.ParaId, Expect = Payment },
            With = "Betaling",
            WithMarked = "Betaling",
            Mode = ChangeMode.Direct
        }));

        Assert.Contains("not both", Assert.Single(report.Errors).Message);
    }

    [Fact]
    public void A_literal_angle_bracket_that_is_not_a_tag_is_text()
    {
        var office = Word();
        var bytes = PaymentMemo();
        var paragraph = office.Inspect(bytes).Paragraphs.Single(p => p.Text == Payment);

        using var result = office.Commit(new StreamHandle(new MemoryStream(bytes)), Plan(new ChangeTextOp
        {
            Target = new TextSpanAnchor { ParaId = paragraph.ParaId, Expect = Payment },
            WithMarked = "Pay if a < b within <1>30 days</1>.",
            Mode = ChangeMode.Direct
        }));

        Assert.True(result.Committed, Errors(result));
        Assert.Equal("Pay if a < b within 30 days.",
            office.Inspect(result.ToBytes()).Paragraphs.Single(p => p.ParaId == paragraph.ParaId).Text);
    }

    // ── withMarked in PowerPoint ─────────────────────────────────────────────

    [Fact]
    public void A_deck_paragraph_takes_a_marked_rewrite_and_keeps_its_bold_run()
    {
        var office = new OfficeAgentClient(new PowerPointModule());
        var deck = PptxFactory.DeckWithBoldRun();
        var paragraph = office.Inspect(deck).Paragraphs.Single(p => p.Text == "plain BOLD text tail");
        Assert.Equal("plain <1>BOLD text</1> tail", paragraph.MarkedText);

        using var result = office.Commit(new StreamHandle(new MemoryStream(deck)), new DocumentPlan
        {
            Format = DocFormat.PowerPoint,
            Operations = new PlanOperation[]
            {
                new ChangeTextOp
                {
                    Target = new TextSpanAnchor { ParaId = paragraph.ParaId, Expect = paragraph.Text },
                    WithMarked = "<1>VETTE tekst</1> eerst, dan gewoon",
                    Mode = ChangeMode.Direct
                }
            }
        });

        Assert.True(result.Committed, Errors(result));
        using var stream = new MemoryStream(result.ToBytes());
        using var document = PresentationDocument.Open(stream, isEditable: false);
        var runs = document.PresentationPart!.SlideParts.First().Slide.Descendants<A.Paragraph>()
            .Single(p => p.InnerText.Contains("VETTE", StringComparison.Ordinal))
            .Elements<A.Run>()
            .Select(r => (string.Concat(r.Elements<A.Text>().Select(t => t.Text)), r.RunProperties?.Bold?.Value == true));
        Assert.Equal(new[] { ("VETTE tekst", true), (" eerst, dan gewoon", false) }, runs);
        Assert.DoesNotContain(new OpenXmlValidator(FileFormatVersions.Office2019).Validate(document),
            e => e.Part?.Uri.ToString().Contains("/slides/", StringComparison.Ordinal) == true);
    }

    // ── "match": "all" ───────────────────────────────────────────────────────

    [Fact]
    public async Task Match_all_changes_every_occurrence_including_several_in_one_paragraph()
    {
        using var workspace = new Workspace();
        workspace.Stage("contract.docx", DocxFactory.Contract());

        var response = Parse(await workspace.Tools.EditDocument("workspace", "contract.docx",
            Rename("Acme Corp", "Globex Inc.")));

        Assert.True(response.GetProperty("committed").GetBoolean(), response.GetRawText());
        Assert.Equal(2, response.GetProperty("changes").GetArrayLength());
        var text = Paragraphs(workspace.Read("contract.docx"));
        Assert.Contains("Globex Inc. shall provide services to Globex Inc..", text);
        Assert.DoesNotContain(text, t => t.Contains("Acme", StringComparison.Ordinal));
    }

    [Fact]
    public async Task Match_all_honours_case_and_whole_word()
    {
        using var workspace = new Workspace();
        workspace.Stage("names.docx", Memo(
            Para("0000000A", Run("ACME and Acme and Acmes.")),
            Para("0000000B", Run("acme in lower case."))));

        var response = Parse(await workspace.Tools.EditDocument("workspace", "names.docx",
            """[{ "op": "changeText", "target": { "find": "Acme", "match": "all", "caseSensitive": true, "wholeWord": true }, "with": "Globex", "mode": "Direct" }]"""));

        Assert.True(response.GetProperty("committed").GetBoolean(), response.GetRawText());
        Assert.Equal(new[] { "ACME and Globex and Acmes.", "acme in lower case." }, Paragraphs(workspace.Read("names.docx")));
    }

    [Fact]
    public async Task Match_all_that_finds_nothing_saves_nothing_and_says_so()
    {
        using var workspace = new Workspace();
        workspace.Stage("contract.docx", DocxFactory.Contract());
        var before = workspace.Hash("contract.docx");

        var response = Parse(await workspace.Tools.EditDocument("workspace", "contract.docx",
            Rename("Initech", "Globex Inc.")));

        Assert.False(response.GetProperty("committed").GetBoolean());
        Assert.Equal(ValidationErrorCodes.AnchorNotFound,
            response.GetProperty("errors")[0].GetProperty("code").GetString());
        Assert.Equal(before, workspace.Hash("contract.docx"));
    }

    [Fact]
    public async Task A_match_that_is_neither_an_index_nor_all_is_refused()
    {
        using var workspace = new Workspace();
        workspace.Stage("contract.docx", DocxFactory.Contract());

        var response = Parse(await workspace.Tools.EditDocument("workspace", "contract.docx",
            """[{ "op": "changeText", "target": { "find": "Acme Corp", "match": "every" }, "with": "Globex" }]"""));

        Assert.Equal(ToolErrorCodes.InvalidArgument, response.GetProperty("errors")[0].GetProperty("code").GetString());
    }

    // ── edit_documents ───────────────────────────────────────────────────────

    [Fact]
    public async Task One_plan_edits_every_document_that_has_the_text_and_leaves_the_rest_unsaved()
    {
        using var workspace = new Workspace();
        workspace.Stage("a.docx", DocxFactory.Contract());
        workspace.Stage("b.docx", DocxFactory.Contract());
        workspace.Stage("c.docx", PaymentMemo());
        var untouched = workspace.Hash("c.docx");

        var response = Parse(await workspace.Tools.EditDocuments("workspace",
            new[] { "a.docx", "b.docx", "c.docx" }, Rename("Acme Corp", "Globex Inc.")));

        Assert.True(response.GetProperty("isValid").GetBoolean(), response.GetRawText());
        var summary = response.GetProperty("summary");
        Assert.Equal(3, summary.GetProperty("total").GetInt32());
        Assert.Equal(2, summary.GetProperty("committed").GetInt32());
        Assert.Equal(1, summary.GetProperty("unchanged").GetInt32());

        var items = response.GetProperty("items").EnumerateArray().ToList();
        Assert.Equal(new[] { "a.docx", "b.docx", "c.docx" }, items.Select(i => i.GetProperty("source").GetString()));
        Assert.Equal(new[] { "committed", "committed", "unchanged" }, items.Select(i => i.GetProperty("status").GetString()));
        Assert.All(items.Take(2), item =>
        {
            Assert.Equal(2, item.GetProperty("changes").GetArrayLength());
            Assert.False(string.IsNullOrEmpty(item.GetProperty("receipt").GetProperty("outputSha256").GetString()));
        });

        Assert.DoesNotContain(Paragraphs(workspace.Read("a.docx")), t => t.Contains("Acme", StringComparison.Ordinal));
        Assert.DoesNotContain(Paragraphs(workspace.Read("b.docx")), t => t.Contains("Acme", StringComparison.Ordinal));
        Assert.Equal(untouched, workspace.Hash("c.docx"));
    }

    [Fact]
    public async Task A_preview_reports_every_document_and_writes_none()
    {
        using var workspace = new Workspace();
        workspace.Stage("a.docx", DocxFactory.Contract());
        workspace.Stage("b.docx", DocxFactory.Contract());
        var before = (workspace.Hash("a.docx"), workspace.Hash("b.docx"));

        var response = Parse(await workspace.Tools.EditDocuments("workspace",
            new[] { "a.docx", "b.docx" }, Rename("Acme Corp", "Globex Inc."), preview: true));

        Assert.True(response.GetProperty("preview").GetBoolean());
        Assert.Equal(2, response.GetProperty("summary").GetProperty("previewed").GetInt32());
        Assert.All(response.GetProperty("items").EnumerateArray(), item =>
        {
            Assert.False(item.GetProperty("committed").GetBoolean());
            Assert.Equal(2, item.GetProperty("changes").GetArrayLength());
        });
        Assert.Equal(before, (workspace.Hash("a.docx"), workspace.Hash("b.docx")));
    }

    [Fact]
    public async Task A_document_that_fails_does_not_stop_the_others()
    {
        using var workspace = new Workspace();
        workspace.Stage("a.docx", DocxFactory.Contract());
        workspace.Stage("c.docx", DocxFactory.Contract());

        var response = Parse(await workspace.Tools.EditDocuments("workspace",
            new[] { "a.docx", "missing.docx", "c.docx" }, Rename("Acme Corp", "Globex Inc.")));

        Assert.False(response.GetProperty("isValid").GetBoolean());
        Assert.Equal(new[] { "committed", "failed", "committed" },
            response.GetProperty("items").EnumerateArray().Select(i => i.GetProperty("status").GetString()));
        var failed = response.GetProperty("items")[1];
        Assert.Equal("notWritten", failed.GetProperty("writeOutcome").GetString());
        Assert.NotEqual(0, failed.GetProperty("errors").GetArrayLength());
    }

    [Theory]
    [InlineData(0)]
    [InlineData(OfficeAgentTools.MaximumBatchDocuments + 1)]
    public async Task A_batch_with_no_documents_or_too_many_is_refused_before_anything_opens(int count)
    {
        using var workspace = new Workspace();
        var sources = Enumerable.Range(0, count).Select(i => $"doc{i}.docx").ToArray();

        var response = Parse(await workspace.Tools.EditDocuments("workspace", sources, Rename("Acme", "Globex")));

        Assert.Equal(ToolErrorCodes.InvalidArgument, response.GetProperty("errors")[0].GetProperty("code").GetString());
    }

    [Fact]
    public async Task A_batch_naming_one_document_twice_or_carrying_a_snapshot_is_refused()
    {
        using var workspace = new Workspace();
        workspace.Stage("a.docx", DocxFactory.Contract());

        var twice = Parse(await workspace.Tools.EditDocuments("workspace", new[] { "a.docx", "A.docx" }, Rename("Acme", "Globex")));
        Assert.Contains("more than once", twice.GetProperty("errors")[0].GetProperty("message").GetString());

        var snapshot = Parse(await workspace.Tools.EditDocuments("workspace", new[] { "a.docx" },
            """{ "snapshot": { "eTag": "abc" }, "operations": [] }"""));
        Assert.Contains("snapshot", snapshot.GetProperty("errors")[0].GetProperty("message").GetString());
    }

    [Fact]
    public async Task A_server_that_requires_snapshots_refuses_a_batch()
    {
        using var workspace = new Workspace();
        workspace.Stage("a.docx", DocxFactory.Contract());
        var tools = new OfficeAgentTools(workspace.Client) { RequirePlanSnapshot = true };

        var response = Parse(await tools.EditDocuments("workspace", new[] { "a.docx" }, Rename("Acme", "Globex")));

        Assert.Equal(ToolErrorCodes.SnapshotRequired, response.GetProperty("errors")[0].GetProperty("code").GetString());
    }

    // ── helpers ──────────────────────────────────────────────────────────────

    private static OfficeAgentClient Word() => new(new WordModule());

    private static DocumentPlan Plan(params PlanOperation[] operations) => new() { Operations = operations };

    private static string Rename(string from, string to) =>
        $$"""[{ "op": "changeText", "target": { "find": "{{from}}", "match": "all" }, "with": "{{to}}", "mode": "Direct" }]""";

    private static JsonElement Parse(string json) => JsonDocument.Parse(json).RootElement.Clone();

    private static string Errors(ApplyResult result) =>
        string.Join("; ", result.Report.Errors.Select(e => $"{e.Code}: {e.Message}"));

    private static byte[] PaymentMemo() => Memo(
        Para("0000000A", Run("Payment is due within "), Run("30 days", bold: true), Run(" of invoice.")),
        Para("0000000B", Run("Plain closing line.")));

    private static Run Run(string text, bool bold = false, bool italic = false)
    {
        var run = new Run();
        if (bold || italic)
        {
            var properties = new RunProperties();
            if (bold) properties.AppendChild(new Bold());
            if (italic) properties.AppendChild(new Italic());
            run.AppendChild(properties);
        }

        run.AppendChild(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        return run;
    }

    private static Paragraph Para(string paraId, params Run[] runs)
    {
        var paragraph = new Paragraph { ParagraphId = paraId };
        foreach (var run in runs) paragraph.AppendChild(run);
        return paragraph;
    }

    private static byte[] Memo(params Paragraph[] paragraphs)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(paragraphs));
            main.Document.Save();
        }

        return stream.ToArray();
    }

    private static string Text(Run run) => string.Concat(run.Elements<Text>().Select(t => t.Text));

    private static IReadOnlyList<(string, bool)> FirstParagraphRuns(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, isEditable: false);
        return document.MainDocumentPart!.Document.Body!.Elements<Paragraph>()
            .First()
            .Elements<Run>()
            .Select(r => (Text(r), r.RunProperties?.Bold is not null))
            .ToList();
    }

    private static IReadOnlyList<string> Paragraphs(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, isEditable: false);
        return document.MainDocumentPart!.Document.Body!.Descendants<Paragraph>()
            .Select(p => string.Concat(p.Descendants<Text>().Select(t => t.Text)))
            .ToList();
    }

    private static IReadOnlyList<ValidationErrorInfo> Validate(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, isEditable: false);
        return new OpenXmlValidator(FileFormatVersions.Office2019).Validate(document).ToList();
    }

    private sealed class Workspace : IDisposable
    {
        private readonly ServiceProvider _services;

        public Workspace()
        {
            Root = Path.Combine(Path.GetTempPath(), $"officeagent-batch-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Root);
            var services = new ServiceCollection();
            services.AddWordFormat();
            services.AddFileSystemDocumentProvider("workspace", Root, options => options.DefaultChangeMode = ChangeMode.Direct);
            services.AddOfficeAgent();
            _services = services.BuildServiceProvider();
            Client = _services.GetRequiredService<OfficeAgentClient>();
            Tools = new OfficeAgentTools(Client);
        }

        public string Root { get; }

        public OfficeAgentClient Client { get; }

        public OfficeAgentTools Tools { get; }

        public void Stage(string name, byte[] bytes) => File.WriteAllBytes(Path.Combine(Root, name), bytes);

        public byte[] Read(string name) => File.ReadAllBytes(Path.Combine(Root, name));

        public string Hash(string name) => Convert.ToHexString(SHA256.HashData(Read(name)));

        public void Dispose()
        {
            _services.Dispose();
            if (Directory.Exists(Root)) Directory.Delete(Root, recursive: true);
        }
    }
}
