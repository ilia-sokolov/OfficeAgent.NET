using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

public class NewVerbsTests
{
    private static OfficeAgentClient Office() => new(new WordModule());
    private static StreamHandle Handle(byte[] bytes) => new(new MemoryStream(bytes), "doc.docx");

    [Fact]
    public void Inspect_surfaces_tables_as_node_anchors()
    {
        var bytes = BuildDocWithTable(("Milestone", "Date"), ("Kickoff", "2026-06-01"));
        var inspect = Office().Inspect(Handle(bytes));

        var tableNode = inspect.Nodes.Single(n => n.Kind == "table");
        Assert.Equal("table#0", tableNode.Path);
        Assert.NotNull(tableNode.Anchor);
        Assert.Contains("Milestone", tableNode.Summary);
    }

    [Fact]
    public void AddTableRows_appends_rows_to_existing_table()
    {
        var bytes = BuildDocWithTable(("Milestone", "Date"), ("Kickoff", "2026-06-01"));
        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new InsertTableRowsOp
                {
                    Target = new NodeAnchor { Kind = "table", Path = "table#0" },
                    Rows = new[]
                    {
                        new[] { "Design review", "2026-07-01" },
                        new[] { "Beta",          "2026-08-15" }
                    }
                }
            }
        };

        var result = Office().Commit(Handle(bytes), plan);
        Assert.True(result.Committed);

        using var doc = WordprocessingDocument.Open(new MemoryStream(OfficeAgentClient.ToBytes(result)), false);
        var table = doc.MainDocumentPart!.Document.Body!.Descendants<Table>().Single();
        var rows = table.Elements<TableRow>().ToList();

        // Header + original 1 + 2 added
        Assert.Equal(4, rows.Count);
        Assert.Contains("Design review", rows[2].InnerText);
        Assert.Contains("2026-08-15", rows[3].InnerText);
    }

    [Fact]
    public void AddTableRows_fails_cleanly_when_table_anchor_not_found()
    {
        var bytes = BuildDocWithTable(("A", "B"), ("1", "2"));
        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new InsertTableRowsOp
                {
                    Target = new NodeAnchor { Kind = "table", Path = "table#99" },
                    Rows = new[] { new[] { "x", "y" } }
                }
            }
        };

        var report = Office().Preview(Handle(bytes), plan);
        Assert.False(report.IsValid);
        Assert.Contains(report.Errors, e => e.Code == ValidationErrorCodes.AnchorNotFound);
    }

    [Fact]
    public void FormatRun_highlights_a_word_yellow()
    {
        var bytes = BuildSimpleDoc("This is an important word in a paragraph.");
        var inspect = Office().Inspect(Handle(bytes));
        var paraId = inspect.Paragraphs.First().ParaId;

        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new FormatOp
                {
                    Target = new TextSpanAnchor { ParaId = paraId, Expect = "important", Occurrence = 0 },
                    Highlight = "yellow"
                }
            }
        };

        var result = Office().Commit(Handle(bytes), plan);
        Assert.True(result.Committed);

        using var doc = WordprocessingDocument.Open(new MemoryStream(OfficeAgentClient.ToBytes(result)), false);
        var runs = doc.MainDocumentPart!.Document.Body!.Descendants<Run>().ToList();

        var highlightedRun = runs.SingleOrDefault(r =>
            r.RunProperties?.GetFirstChild<Highlight>()?.Val?.Value == HighlightColorValues.Yellow);
        Assert.NotNull(highlightedRun);
        Assert.Equal("important", string.Concat(highlightedRun!.Elements<Text>().Select(t => t.Text)));
    }

    [Fact]
    public void FormatRun_can_combine_bold_italic_color()
    {
        var bytes = BuildSimpleDoc("Make this BOLD italic red please.");
        var inspect = Office().Inspect(Handle(bytes));
        var paraId = inspect.Paragraphs.First().ParaId;

        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new FormatOp
                {
                    Target = new TextSpanAnchor { ParaId = paraId, Expect = "BOLD italic red" },
                    Bold = true,
                    Italic = true,
                    Color = "FF0000"
                }
            }
        };

        var result = Office().Commit(Handle(bytes), plan);
        Assert.True(result.Committed);

        using var doc = WordprocessingDocument.Open(new MemoryStream(OfficeAgentClient.ToBytes(result)), false);
        var styled = doc.MainDocumentPart!.Document.Body!.Descendants<Run>()
            .Single(r => string.Concat(r.Elements<Text>().Select(t => t.Text)) == "BOLD italic red");
        Assert.NotNull(styled.RunProperties?.GetFirstChild<Bold>());
        Assert.NotNull(styled.RunProperties?.GetFirstChild<Italic>());
        Assert.Equal("FF0000", styled.RunProperties?.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Color>()?.Val?.Value);
    }

    [Fact]
    public void FormatRun_with_no_properties_is_rejected()
    {
        var bytes = BuildSimpleDoc("anything");
        var paraId = Office().Inspect(Handle(bytes)).Paragraphs.First().ParaId;
        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new FormatOp { Target = new TextSpanAnchor { ParaId = paraId, Expect = "anything" } }
            }
        };

        var report = Office().Preview(Handle(bytes), plan);
        Assert.False(report.IsValid);
        Assert.Contains(report.Errors, e => e.Code == ValidationErrorCodes.InvalidOperation);
    }

    private static byte[] BuildSimpleDoc(string text)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var p = new Paragraph { ParagraphId = "11111111" };
            p.AppendChild(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
            main.Document = new Document(new Body(p));
            main.Document.Save();
        }
        return ms.ToArray();
    }

    [Fact]
    public void RemoveTable_deletes_the_whole_table()
    {
        var bytes = BuildDocWithTable(("Milestone", "Date"), ("Kickoff", "2026-06-01"));
        Assert.Single(Office().Inspect(Handle(bytes)).Nodes.Where(n => n.Kind == "table"));

        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new RemoveTableOp
                {
                    Target = new NodeAnchor { Kind = "table", Path = "table#0" }
                }
            }
        };

        var result = Office().Commit(Handle(bytes), plan);
        Assert.True(result.Committed);

        using var doc = WordprocessingDocument.Open(new MemoryStream(OfficeAgentClient.ToBytes(result)), false);
        Assert.Empty(doc.MainDocumentPart!.Document.Body!.Descendants<Table>());
    }

    [Fact]
    public void RemoveTable_fails_cleanly_when_table_anchor_not_found()
    {
        var bytes = BuildDocWithTable(("A", "B"), ("1", "2"));
        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new RemoveTableOp
                {
                    Target = new NodeAnchor { Kind = "table", Path = "table#99" }
                }
            }
        };

        var report = Office().Preview(Handle(bytes), plan);
        Assert.False(report.IsValid);
        Assert.Contains(report.Errors, e => e.Code == ValidationErrorCodes.AnchorNotFound);
    }

    private static byte[] BuildDocWithTable(params (string A, string B)[] rows)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var table = new Table(new TableProperties(new TableStyle { Val = "TableGrid" }));
            foreach (var (a, b) in rows)
                table.AppendChild(new TableRow(
                    new TableCell(new Paragraph(new Run(new Text(a) { Space = SpaceProcessingModeValues.Preserve }))),
                    new TableCell(new Paragraph(new Run(new Text(b) { Space = SpaceProcessingModeValues.Preserve })))));
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Intro"))),
                table,
                new Paragraph(new Run(new Text("Outro")))));
            main.Document.Save();
        }
        return ms.ToArray();
    }
}
