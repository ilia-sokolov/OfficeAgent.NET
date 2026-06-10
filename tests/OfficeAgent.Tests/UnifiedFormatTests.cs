using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

public class UnifiedFormatTests
{
    private static OfficeAgentClient Office() => new(new WordModule());
    private static StreamHandle Handle(byte[] bytes) => new(new MemoryStream(bytes), "doc.docx");

    [Fact]
    public void Format_paragraph_with_styleId_alignment_and_run_properties()
    {
        var bytes = BuildContractDoc();
        var inspect = Office().Inspect(Handle(bytes));
        var paraId = inspect.Paragraphs.First(p => p.Text.Contains("Acme")).ParaId;

        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new FormatOp
                {
                    Target = new TextSpanAnchor { ParaId = paraId, Expect = "" },
                    StyleId = "Quote",
                    Alignment = "center",
                    Bold = true,
                    FontFamily = "Calibri",
                    SizeHalfPoints = 24,
                    Color = "0000FF"
                }
            }
        };

        var result = Office().Commit(Handle(bytes), plan);
        Assert.True(result.Committed);

        using var doc = WordprocessingDocument.Open(new MemoryStream(OfficeAgentClient.ToBytes(result)), false);
        var paragraph = doc.MainDocumentPart!.Document.Body!.Descendants<Paragraph>()
            .Single(p => p.InnerText.Contains("Acme"));
        Assert.Equal("Quote", paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value);
        Assert.Equal(JustificationValues.Center, paragraph.ParagraphProperties?.Justification?.Val?.Value);

        var firstRun = paragraph.Elements<Run>().First();
        Assert.NotNull(firstRun.RunProperties?.GetFirstChild<Bold>());
        Assert.Equal("Calibri", firstRun.RunProperties?.GetFirstChild<RunFonts>()?.Ascii?.Value);
        Assert.Equal("24", firstRun.RunProperties?.GetFirstChild<FontSize>()?.Val?.Value);
        Assert.Equal("0000FF", firstRun.RunProperties?.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Color>()?.Val?.Value);
    }

    [Fact]
    public void Format_text_span_highlights_only_the_matched_substring()
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
        var highlightedRun = doc.MainDocumentPart!.Document.Body!.Descendants<Run>().Single(r =>
            r.RunProperties?.GetFirstChild<Highlight>()?.Val?.Value == HighlightColorValues.Yellow);
        Assert.Equal("important", string.Concat(highlightedRun.Elements<Text>().Select(t => t.Text)));
    }

    [Fact]
    public void Format_table_applies_styleId_and_borders()
    {
        var bytes = BuildCountriesTable();
        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new FormatOp
                {
                    Target = new NodeAnchor { Kind = "table", Path = "table#0" },
                    StyleId = "TableGrid",
                    BorderStyle = "double",
                    BorderSizeEighths = 8,
                    BorderColor = "FF0000"
                }
            }
        };

        var result = Office().Commit(Handle(bytes), plan);
        Assert.True(result.Committed);

        using var doc = WordprocessingDocument.Open(new MemoryStream(OfficeAgentClient.ToBytes(result)), false);
        var table = doc.MainDocumentPart!.Document.Body!.Descendants<Table>().Single();
        var properties = table.GetFirstChild<TableProperties>();
        Assert.Equal("TableGrid", properties?.TableStyle?.Val?.Value);

        var borders = properties?.TableBorders!;
        Assert.Equal(BorderValues.Double, borders.TopBorder?.Val?.Value);
        Assert.Equal("FF0000", borders.TopBorder?.Color?.Value);
    }

    [Fact]
    public void Format_table_row_applies_run_properties_to_every_cell_paragraph()
    {
        var bytes = BuildCountriesTable();
        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new FormatOp
                {
                    Target = new NodeAnchor { Kind = "tableRow", Path = "table#0/row#0" },
                    Bold = true,
                    Highlight = "yellow"
                }
            }
        };

        var result = Office().Commit(Handle(bytes), plan);
        Assert.True(result.Committed);

        using var doc = WordprocessingDocument.Open(new MemoryStream(OfficeAgentClient.ToBytes(result)), false);
        var firstRow = doc.MainDocumentPart!.Document.Body!.Descendants<Table>().Single().Elements<TableRow>().First();
        var runs = firstRow.Descendants<Run>().ToList();
        Assert.NotEmpty(runs);
        Assert.All(runs, r =>
        {
            Assert.NotNull(r.RunProperties?.GetFirstChild<Bold>());
            Assert.Equal(HighlightColorValues.Yellow, r.RunProperties?.GetFirstChild<Highlight>()?.Val?.Value);
        });
    }

    [Fact]
    public void Format_table_cell_applies_alignment_and_border()
    {
        var bytes = BuildCountriesTable();
        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new FormatOp
                {
                    Target = new NodeAnchor { Kind = "tableCell", Path = "table#0/cell#1/2" },
                    Alignment = "right",
                    Color = "0000FF",
                    BorderStyle = "single",
                    BorderColor = "00AA00"
                }
            }
        };

        var result = Office().Commit(Handle(bytes), plan);
        Assert.True(result.Committed);

        using var doc = WordprocessingDocument.Open(new MemoryStream(OfficeAgentClient.ToBytes(result)), false);
        var cell = doc.MainDocumentPart!.Document.Body!.Descendants<Table>().Single()
            .Elements<TableRow>().ElementAt(1)
            .Elements<TableCell>().ElementAt(2);

        var paragraph = cell.Elements<Paragraph>().Single();
        Assert.Equal(JustificationValues.Right, paragraph.ParagraphProperties?.Justification?.Val?.Value);

        var run = paragraph.Elements<Run>().Single();
        Assert.Equal("0000FF", run.RunProperties?.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Color>()?.Val?.Value);

        var cellBorders = cell.TableCellProperties?.GetFirstChild<TableCellBorders>();
        Assert.NotNull(cellBorders);
        Assert.Equal("00AA00", cellBorders!.TopBorder?.Color?.Value);
    }

    [Fact]
    public void Format_image_resizes_inline_drawing()
    {
        // Start: insert an image so we have something to resize.
        var bytes = BuildSimpleDoc("paragraph");
        var paraId = Office().Inspect(Handle(bytes)).Paragraphs.First().ParaId;
        var insert = Office().Commit(Handle(bytes), new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new InsertImageOp
                {
                    Target = new TextSpanAnchor { ParaId = paraId, Expect = "" },
                    Base64Bytes = Convert.ToBase64String(MinimalPng()),
                    ImageType = "png",
                    WidthPx = 50, HeightPx = 50,
                    Position = InsertPosition.After
                }
            }
        });
        Assert.True(insert.Committed);

        var withImageBytes = OfficeAgentClient.ToBytes(insert);
        var resize = Office().Commit(Handle(withImageBytes), new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new FormatOp
                {
                    Target = new NodeAnchor { Kind = "image", Path = "image#0" },
                    WidthPx = 400,
                    HeightPx = 200
                }
            }
        });
        Assert.True(resize.Committed);

        using var doc = WordprocessingDocument.Open(new MemoryStream(OfficeAgentClient.ToBytes(resize)), false);
        var extent = doc.MainDocumentPart!.Document.Body!.Descendants<Drawing>()
            .Single().GetFirstChild<DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline>()!.Extent!;
        Assert.Equal(400L * 9525, extent.Cx?.Value);
        Assert.Equal(200L * 9525, extent.Cy?.Value);
    }

    private static byte[] BuildContractDoc()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var p = new Paragraph { ParagraphId = "AAAAAAAA" };
            p.AppendChild(new Run(new Text("Acme Corp shall provide services.") { Space = SpaceProcessingModeValues.Preserve }));
            main.Document = new Document(new Body(p));
            var stylePart = main.AddNewPart<StyleDefinitionsPart>();
            stylePart.Styles = new Styles(
                new Style { Type = StyleValues.Paragraph, StyleId = "Quote", StyleName = new StyleName { Val = "Quote" } });
            stylePart.Styles.Save();
            main.Document.Save();
        }
        return ms.ToArray();
    }

    private static byte[] BuildSimpleDoc(string text)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var p = new Paragraph { ParagraphId = "BBBBBBBB" };
            p.AppendChild(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
            main.Document = new Document(new Body(p));
            main.Document.Save();
        }
        return ms.ToArray();
    }

    private static byte[] BuildCountriesTable()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var table = new Table(
                new TableProperties(new TableStyle { Val = "TableGrid" }),
                Row("Country", "Population mil", "Surface sq km"),
                Row("RU", "146", "17098246"),
                Row("US", "332", "9833520"),
                Row("UK", "68", "243610"));
            main.Document = new Document(new Body(table));
            main.Document.Save();
        }
        return ms.ToArray();
    }

    private static TableRow Row(params string[] cells)
    {
        var row = new TableRow();
        foreach (var c in cells)
            row.AppendChild(new TableCell(new Paragraph(new Run(new Text(c) { Space = SpaceProcessingModeValues.Preserve }))));
        return row;
    }

    private static byte[] MinimalPng() => Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4//8/AwAI/AL+XJ/PNwAAAABJRU5ErkJggg==");
}
