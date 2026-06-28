using DocumentFormat.OpenXml.Packaging;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Word;
using DocumentFormat.OpenXml.Wordprocessing;
using DocFormat = OfficeAgent.Abstractions.DocumentFormat;

namespace OfficeAgent.Tests;

public class FlowTests
{
    private static OfficeAgentClient Office() => new(new WordModule());

    private static StreamHandle Handle(byte[] bytes) => new(new MemoryStream(bytes), "contract.docx");

    private static string ParaId(InspectResult inspect, string contains) =>
        inspect.Paragraphs.First(p => p.Text.Contains(contains, StringComparison.Ordinal)).ParaId;

    [Fact]
    public void Inspect_returns_outline_paragraphs_styles_and_content_controls()
    {
        var office = Office();
        var inspect = office.Inspect(Handle(DocxFactory.Contract()));

        Assert.Equal(DocFormat.Word, inspect.Format);
        Assert.False(string.IsNullOrEmpty(inspect.Snapshot.ETag));
        Assert.Equal(4, inspect.Paragraphs.Count);

        Assert.Contains(inspect.Outline, o => o.Text == DocxFactory.HeadingText && o.Level == 1);
        Assert.Contains(inspect.StructuralAnchors, s => s.Tag == DocxFactory.ClientControlTag && s.Kind == "contentControl");
        Assert.Contains(inspect.Styles.Styles, s => s.Id == "Heading1");

        var clause = inspect.Paragraphs.Single(p => p.ParaId == ParaId(inspect, "shall provide"));
        Assert.Equal(DocxFactory.ClauseText, clause.Text);
    }

    [Fact]
    public void Find_returns_two_content_verified_occurrences()
    {
        var office = Office();
        var hits = office.Find(Handle(DocxFactory.Contract()), new FindQuery("Acme Corp"));

        Assert.Equal(2, hits.Count);
        Assert.All(hits, h => Assert.Equal("Acme Corp", h.Text));
        Assert.Equal(0, ((TextSpanAnchor)hits[0].Anchor).Occurrence);
        Assert.Equal(1, ((TextSpanAnchor)hits[1].Anchor).Occurrence);
    }

    [Fact]
    public void ChangeText_direct_replaces_run_spanning_first_occurrence()
    {
        var office = Office();
        var bytes = DocxFactory.Contract();
        var inspect = office.Inspect(Handle(bytes));

        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new ChangeTextOp
                {
                    Target = new TextSpanAnchor { ParaId = ParaId(inspect, "shall provide"), Expect = "Acme Corp", Occurrence = 0 },
                    With = "Globex",
                    Mode = ChangeMode.Direct
                }
            }
        };

        var result = office.Commit(Handle(bytes), plan);

        Assert.True(result.Committed);
        var text = ClauseLogicalText(OfficeAgentClient.ToBytes(result));
        Assert.Equal("Globex shall provide services to Acme Corp.", text);
    }

    [Fact]
    public void ChangeText_tracked_lands_as_redline()
    {
        var office = Office();
        var bytes = DocxFactory.Contract();
        var inspect = office.Inspect(Handle(bytes));

        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new ChangeTextOp
                {
                    Target = new TextSpanAnchor { ParaId = ParaId(inspect, "shall provide"), Expect = "Acme Corp", Occurrence = 0 },
                    With = "Globex Inc.",
                    Mode = ChangeMode.Tracked
                }
            }
        };

        var result = office.Commit(Handle(bytes), plan);
        Assert.True(result.Committed);

        using var doc = WordprocessingDocument.Open(new MemoryStream(OfficeAgentClient.ToBytes(result)), false);
        var body = doc.MainDocumentPart!.Document.Body!;

        var inserted = body.Descendants<InsertedRun>().ToList();
        var deleted = body.Descendants<DeletedRun>().ToList();

        Assert.Contains(inserted, ins => ins.Descendants<Text>().Any(t => t.Text == "Globex Inc."));
        Assert.Contains(deleted, del => string.Concat(del.Descendants<DeletedText>().Select(t => t.Text)) == "Acme Corp");
    }

    [Fact]
    public void Fill_sets_content_control_value()
    {
        var office = Office();
        var bytes = DocxFactory.Contract();

        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new FillOp
                {
                    Target = new StructuralAnchor { Tag = DocxFactory.ClientControlTag, Kind = "contentControl" },
                    Value = "Globex Inc."
                }
            }
        };

        var result = office.Commit(Handle(bytes), plan);
        Assert.True(result.Committed);

        using var doc = WordprocessingDocument.Open(new MemoryStream(OfficeAgentClient.ToBytes(result)), false);
        var sdt = doc.MainDocumentPart!.Document.Body!.Descendants<SdtRun>().Single();
        var value = string.Concat(sdt.Descendants<Text>().Select(t => t.Text));
        Assert.Equal("Globex Inc.", value);
    }

    [Fact]
    public void Comment_attaches_comment_to_span()
    {
        var office = Office();
        var bytes = DocxFactory.Contract();
        var inspect = office.Inspect(Handle(bytes));

        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new CommentOp
                {
                    Target = new TextSpanAnchor { ParaId = ParaId(inspect, "shall provide"), Expect = "Acme Corp", Occurrence = 1 },
                    Text = "Confirm the counterparty name."
                }
            }
        };

        var result = office.Commit(Handle(bytes), plan);
        Assert.True(result.Committed);

        using var doc = WordprocessingDocument.Open(new MemoryStream(OfficeAgentClient.ToBytes(result)), false);
        var commentsPart = doc.MainDocumentPart!.WordprocessingCommentsPart;
        Assert.NotNull(commentsPart);
        var comment = commentsPart!.Comments.Elements<Comment>().Single();
        Assert.Contains("Confirm the counterparty name.", comment.InnerText);
        Assert.Single(doc.MainDocumentPart!.Document.Body!.Descendants<CommentRangeStart>());
    }

    [Fact]
    public void InsertTable_adds_structured_table()
    {
        var office = Office();
        var bytes = DocxFactory.Contract();
        var inspect = office.Inspect(Handle(bytes));

        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new InsertTableOp
                {
                    Target = new TextSpanAnchor { ParaId = ParaId(inspect, "shall provide"), Expect = DocxFactory.ClauseText },
                    Position = InsertPosition.After,
                    Table = new TableData
                    {
                        Headers = new[] { "Milestone", "Date" },
                        Rows = new[] { new[] { "Kickoff", "2026-06-01" } }
                    }
                }
            }
        };

        var result = office.Commit(Handle(bytes), plan);
        Assert.True(result.Committed);

        using var doc = WordprocessingDocument.Open(new MemoryStream(OfficeAgentClient.ToBytes(result)), false);
        var table = doc.MainDocumentPart!.Document.Body!.Descendants<Table>().Single();
        Assert.Contains("Milestone", table.InnerText);
        Assert.Contains("Kickoff", table.InnerText);
    }

    [Fact]
    public void Format_applies_paragraph_style()
    {
        var office = Office();
        var bytes = DocxFactory.Contract();
        var inspect = office.Inspect(Handle(bytes));

        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new FormatOp
                {
                    Target = new TextSpanAnchor { ParaId = ParaId(inspect, "Effective date"), Expect = DocxFactory.DateText },
                    StyleId = "Quote"
                }
            }
        };

        var result = office.Commit(Handle(bytes), plan);
        Assert.True(result.Committed);

        using var doc = WordprocessingDocument.Open(new MemoryStream(OfficeAgentClient.ToBytes(result)), false);
        var dateParagraph = doc.MainDocumentPart!.Document.Body!
            .Descendants<Paragraph>()
            .Single(p => p.InnerText.Contains("Effective date"));
        Assert.Equal("Quote", dateParagraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value);
    }

    [Fact]
    public void Stale_anchor_fails_validation_and_does_not_commit()
    {
        var office = Office();
        var bytes = DocxFactory.Contract();
        var inspect = office.Inspect(Handle(bytes));

        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new ChangeTextOp
                {
                    Target = new TextSpanAnchor { ParaId = ParaId(inspect, "shall provide"), Expect = "Nonexistent text", Occurrence = 0 },
                    With = "x"
                }
            }
        };

        var report = office.Preview(Handle(bytes), plan);
        Assert.False(report.IsValid);
        Assert.Contains(report.Errors, e => e.Code == ValidationErrorCodes.ExpectMismatch);

        var result = office.Commit(Handle(bytes), plan);
        Assert.False(result.Committed);
        Assert.Null(result.Output);
    }

    [Fact]
    public void DryRun_previews_changes_without_writing()
    {
        var office = Office();
        var bytes = DocxFactory.Contract();
        var inspect = office.Inspect(Handle(bytes));

        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new ChangeTextOp
                {
                    Target = new TextSpanAnchor { ParaId = ParaId(inspect, "shall provide"), Expect = "Acme Corp", Occurrence = 0 },
                    With = "Globex"
                }
            }
        };

        var result = office.Apply(Handle(bytes), plan, ApplyOptions.Preview);

        Assert.True(result.Report.IsValid);
        Assert.False(result.Committed);
        Assert.Null(result.Output);
        var change = Assert.Single(result.Report.Changes);
        Assert.Equal("Acme Corp", change.Before);
        Assert.Equal("Globex", change.After);
    }

    private static string ClauseLogicalText(byte[] bytes)
    {
        using var doc = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        var clause = doc.MainDocumentPart!.Document.Body!
            .Descendants<Paragraph>()
            .Single(p => p.ParagraphId?.Value == "00000002");
        return string.Concat(clause.Descendants<Text>().Select(t => t.Text));
    }
}
