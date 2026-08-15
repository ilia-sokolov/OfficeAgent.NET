using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// WordprocessingML property containers are <em>sequences</em>: <c>w:tblStyle</c> has to
/// precede <c>w:tblW</c>, <c>w:b</c> has to precede <c>w:color</c>, and a border set runs
/// top, left, bottom, right. Appending a property produces a document Word offers to
/// repair - and one that passes a per-operation test whenever the properties happen to be
/// written in schema order anyway. These validate the whole saved package instead.
/// </summary>
public class WordSchemaOrderTests
{
    private static OfficeAgentClient Client() => new(new WordModule());

    [Fact]
    public void A_styled_and_bordered_table_is_schema_valid()
    {
        var client = Client();

        // insertTable writes w:tblW; format then adds w:tblStyle, which must land *before*
        // it. Appending is what produced "unexpected child element 'tblStyle'".
        var withTable = Apply(client, Blank(), new InsertTableOp
        {
            Target = new TextSpanAnchor { ParaId = "auto-0000", Expect = string.Empty },
            Position = InsertPosition.After,
            Table = new TableData { Headers = new[] { "Role", "Rate" }, Rows = new[] { new[] { "Lead", "1450" } } }
        });

        var formatted = Apply(client, withTable, new FormatOp
        {
            Target = new NodeAnchor { Kind = "table", Path = "table#0" },
            StyleId = "TableGrid",
            BorderStyle = "single",
            BorderSizeEighths = 4,
            BorderColor = "000000"
        });

        AssertValid(formatted);

        using var stream = new MemoryStream(formatted);
        using var document = WordprocessingDocument.Open(stream, isEditable: false);
        var properties = document.MainDocumentPart!.Document.Body!.Descendants<TableProperties>().Single();
        var names = properties.ChildElements.Select(c => c.LocalName).ToList();
        Assert.True(names.IndexOf("tblStyle") < names.IndexOf("tblW"),
            $"tblStyle must precede tblW; got {string.Join(", ", names)}");

        var borders = properties.GetFirstChild<TableBorders>()!.ChildElements.Select(c => c.LocalName).ToList();
        Assert.Equal(new[] { "top", "left", "bottom", "right", "insideH", "insideV" }, borders);
    }

    [Fact]
    public void Run_properties_written_in_any_order_still_come_out_in_schema_order()
    {
        var client = Client();

        // colour before bold on the wire; w:b must still precede w:color in the file.
        var applied = Apply(client, Seeded(), new FormatOp
        {
            Target = new TextSpanAnchor { ParaId = "w14:00000002", Expect = "Acme Corp" },
            Color = "FF0000",
            Bold = true,
            SizeHalfPoints = 28,
            Underline = true,
            FontFamily = "Georgia"
        });

        AssertValid(applied);

        using var stream = new MemoryStream(applied);
        using var document = WordprocessingDocument.Open(stream, isEditable: false);
        var rPr = document.MainDocumentPart!.Document.Body!
            .Descendants<Run>()
            .First(r => r.RunProperties?.Bold is not null)
            .RunProperties!;

        var names = rPr.ChildElements.Select(c => c.LocalName).ToList();
        Assert.True(names.IndexOf("rFonts") < names.IndexOf("b"), string.Join(", ", names));
        Assert.True(names.IndexOf("b") < names.IndexOf("color"), string.Join(", ", names));
        Assert.True(names.IndexOf("color") < names.IndexOf("sz"), string.Join(", ", names));
        Assert.True(names.IndexOf("sz") < names.IndexOf("u"), string.Join(", ", names));
    }

    [Fact]
    public void A_bordered_and_indented_paragraph_is_schema_valid()
    {
        var client = Client();

        var applied = Apply(client, Seeded(), new FormatOp
        {
            Target = new TextSpanAnchor { ParaId = "w14:00000002", Expect = string.Empty },
            Alignment = "center",
            IndentLeftTwips = 720,
            SpacingBeforeTwips = 120,
            BorderStyle = "single",
            StyleId = "Quote"
        });

        AssertValid(applied);

        using var stream = new MemoryStream(applied);
        using var document = WordprocessingDocument.Open(stream, isEditable: false);
        var pPr = document.MainDocumentPart!.Document.Body!
            .Descendants<ParagraphProperties>()
            .First(p => p.GetFirstChild<ParagraphBorders>() is not null);

        var names = pPr.ChildElements.Select(c => c.LocalName).ToList();
        Assert.True(names.IndexOf("pStyle") < names.IndexOf("pBdr"), string.Join(", ", names));
        Assert.True(names.IndexOf("pBdr") < names.IndexOf("spacing"), string.Join(", ", names));
        Assert.True(names.IndexOf("spacing") < names.IndexOf("ind"), string.Join(", ", names));
        Assert.True(names.IndexOf("ind") < names.IndexOf("jc"), string.Join(", ", names));
    }

    [Fact]
    public void A_bordered_cell_is_schema_valid()
    {
        var client = Client();
        var withTable = Apply(client, Blank(), new InsertTableOp
        {
            Target = new TextSpanAnchor { ParaId = "auto-0000", Expect = string.Empty },
            Position = InsertPosition.After,
            Table = new TableData { Headers = new[] { "A", "B" }, Rows = new[] { new[] { "1", "2" } } }
        });

        var applied = Apply(client, withTable, new FormatOp
        {
            Target = new NodeAnchor { Kind = "tableCell", Path = "table#0/cell#1/1" },
            BorderStyle = "single",
            Alignment = "right"
        });

        AssertValid(applied);
    }

    // ── helpers ───────────────────────────────────────────────────────────────

    [Fact]
    public void A_page_break_before_lands_ahead_of_the_indent_and_spacing()
    {
        var client = Client();

        // Every one of these writes into w:pPr, which is a sequence: w:pageBreakBefore has
        // to precede w:spacing and w:ind however late it is asked for.
        var document = Apply(client, Seeded(), new FormatOp
        {
            Target = new TextSpanAnchor { ParaId = "w14:00000002", Expect = string.Empty },
            SpacingBeforeTwips = 240,
            IndentLeftTwips = 480,
            PageBreakBefore = true
        });

        using var stream = new MemoryStream(document);
        using var opened = WordprocessingDocument.Open(stream, isEditable: false);
        var properties = opened.MainDocumentPart!.Document.Body!
            .Descendants<Paragraph>()
            .First(p => p.ParagraphId?.Value == "00000002")
            .ParagraphProperties!;

        var names = properties.ChildElements.Select(c => c.LocalName).ToList();
        Assert.Contains("pageBreakBefore", names);
        Assert.True(names.IndexOf("pageBreakBefore") < names.IndexOf("spacing"), string.Join(", ", names));
        Assert.True(names.IndexOf("pageBreakBefore") < names.IndexOf("ind"), string.Join(", ", names));
        AssertValid(document);
    }

    [Fact]
    public void A_page_break_before_can_be_taken_off_again()
    {
        var client = Client();

        var broken = Apply(client, Seeded(), new FormatOp
        {
            Target = new TextSpanAnchor { ParaId = "w14:00000002", Expect = string.Empty },
            PageBreakBefore = true
        });
        var healed = Apply(client, broken, new FormatOp
        {
            Target = new TextSpanAnchor { ParaId = "w14:00000002", Expect = string.Empty },
            PageBreakBefore = false
        });

        using var stream = new MemoryStream(healed);
        using var opened = WordprocessingDocument.Open(stream, isEditable: false);
        Assert.Empty(opened.MainDocumentPart!.Document.Body!.Descendants<PageBreakBefore>());
        AssertValid(healed);
    }

    [Fact]
    public void A_negative_first_line_indent_becomes_a_hanging_indent()
    {
        var client = Client();

        // The standard hanging bullet: text at 680, its first line pulled back to 340 so
        // the dash sits outside. w:firstLine is unsigned, so this has to become w:hanging.
        var document = Apply(client, Seeded(), new FormatOp
        {
            Target = new TextSpanAnchor { ParaId = "w14:00000002", Expect = string.Empty },
            IndentLeftTwips = 680,
            IndentFirstLineTwips = -340
        });

        using var stream = new MemoryStream(document);
        using var opened = WordprocessingDocument.Open(stream, isEditable: false);
        var indent = opened.MainDocumentPart!.Document.Body!
            .Descendants<Paragraph>()
            .First(p => p.ParagraphId?.Value == "00000002")
            .ParagraphProperties!.GetFirstChild<Indentation>()!;

        Assert.Equal("340", indent.Hanging!.Value);
        Assert.Equal("680", indent.Left!.Value);
        Assert.Null(indent.FirstLine);
        AssertValid(document);
    }

    [Fact]
    public void A_positive_first_line_indent_clears_a_hanging_one()
    {
        var client = Client();

        var hanging = Apply(client, Seeded(), new FormatOp
        {
            Target = new TextSpanAnchor { ParaId = "w14:00000002", Expect = string.Empty },
            IndentFirstLineTwips = -340
        });
        var indented = Apply(client, hanging, new FormatOp
        {
            Target = new TextSpanAnchor { ParaId = "w14:00000002", Expect = string.Empty },
            IndentFirstLineTwips = 240
        });

        using var stream = new MemoryStream(indented);
        using var opened = WordprocessingDocument.Open(stream, isEditable: false);
        var indent = opened.MainDocumentPart!.Document.Body!
            .Descendants<Paragraph>()
            .First(p => p.ParagraphId?.Value == "00000002")
            .ParagraphProperties!.GetFirstChild<Indentation>()!;

        // Both set at once is contradictory, and Word resolves it in a way the caller did
        // not ask for - so switching direction has to clear the other.
        Assert.Equal("240", indent.FirstLine!.Value);
        Assert.Null(indent.Hanging);
        AssertValid(indented);
    }

    [Fact]
    public void A_border_can_be_drawn_on_one_edge_only()
    {
        var client = Client();

        // A pull quote's rule. Drawn on all four edges it is a callout box, which is a
        // different thing entirely.
        var document = Apply(client, Seeded(), new FormatOp
        {
            Target = new TextSpanAnchor { ParaId = "w14:00000002", Expect = string.Empty },
            BorderStyle = "single",
            BorderColor = "C8632B",
            BorderSizeEighths = 12,
            BorderEdges = "left"
        });

        using var stream = new MemoryStream(document);
        using var opened = WordprocessingDocument.Open(stream, isEditable: false);
        var borders = opened.MainDocumentPart!.Document.Body!
            .Descendants<Paragraph>()
            .First(p => p.ParagraphId?.Value == "00000002")
            .ParagraphProperties!.GetFirstChild<ParagraphBorders>()!;

        var edge = Assert.Single(borders.ChildElements);
        Assert.IsType<LeftBorder>(edge);
        Assert.Equal("C8632B", ((LeftBorder)edge).Color!.Value);
        AssertValid(document);
    }

    [Fact]
    public void Named_edges_keep_the_order_the_schema_declares()
    {
        var client = Client();

        // Asked for out of order on purpose: w:pBdr is a sequence, so they must come back
        // top, left, bottom, right however they were named.
        var document = Apply(client, Seeded(), new FormatOp
        {
            Target = new TextSpanAnchor { ParaId = "w14:00000002", Expect = string.Empty },
            BorderStyle = "single",
            BorderEdges = "right, top, bottom"
        });

        using var stream = new MemoryStream(document);
        using var opened = WordprocessingDocument.Open(stream, isEditable: false);
        var names = opened.MainDocumentPart!.Document.Body!
            .Descendants<Paragraph>()
            .First(p => p.ParagraphId?.Value == "00000002")
            .ParagraphProperties!.GetFirstChild<ParagraphBorders>()!
            .ChildElements.Select(c => c.LocalName).ToList();

        Assert.Equal(new[] { "top", "bottom", "right" }, names);
        AssertValid(document);
    }

    [Fact]
    public void An_edge_that_is_not_a_side_is_refused()
    {
        var client = Client();

        using var report = new MemoryStream(Seeded());
        var result = new OfficeAgentClient(new WordModule()).Preview(
            new StreamHandle(report),
            new DocumentPlan
            {
                Operations = new PlanOperation[]
                {
                    new FormatOp
                    {
                        Target = new TextSpanAnchor { ParaId = "w14:00000002", Expect = string.Empty },
                        BorderStyle = "single",
                        BorderEdges = "left, middle"
                    }
                }
            });

        var error = Assert.Single(result.Errors);
        Assert.Equal(ValidationErrorCodes.InvalidOperation, error.Code);
        Assert.Contains("insideH", error.Message);
    }

    [Fact]
    public void Ruling_a_row_rules_every_cell_in_it()
    {
        var client = Client();

        var withTable = Apply(client, Blank(), new InsertTableOp
        {
            Target = new TextSpanAnchor { ParaId = "auto-0000", Expect = string.Empty },
            Position = InsertPosition.After,
            Table = new TableData
            {
                Headers = new[] { "Metric", "2025", "2026" },
                Rows = new List<IReadOnlyList<string>> { new[] { "ARR", "18.0", "28.4" } }
            }
        });

        var table = client.Inspect(withTable).Nodes.Single(n => n.Kind == "table");
        var ruled = Apply(client, withTable, new FormatOp
        {
            Target = new NodeAnchor { Kind = "tableRow", Path = table.Path + "/row#0" },
            BorderStyle = "single",
            BorderColor = "12161C",
            BorderSizeEighths = 8,
            BorderEdges = "bottom"
        });

        using var stream = new MemoryStream(ruled);
        using var opened = WordprocessingDocument.Open(stream, isEditable: false);
        var header = opened.MainDocumentPart!.Document.Body!
            .Descendants<Table>().Single()
            .Elements<TableRow>().First();

        // A row carries no borders of its own; three cells means three bottom borders.
        var borders = header.Descendants<TableCellBorders>().ToList();
        Assert.Equal(3, borders.Count);
        Assert.All(borders, b =>
        {
            var edge = Assert.Single(b.ChildElements);
            Assert.IsType<BottomBorder>(edge);
            Assert.Equal("12161C", ((BottomBorder)edge).Color!.Value);
        });

        // The body row is untouched.
        var body = opened.MainDocumentPart.Document.Body.Descendants<Table>().Single()
            .Elements<TableRow>().Last();
        Assert.Empty(body.Descendants<TableCellBorders>());
        AssertValid(ruled);
    }

    [Fact]
    public void Column_widths_reach_the_grid_the_cells_and_the_table()
    {
        var client = Client();
        var withTable = Apply(client, Blank(), new InsertTableOp
        {
            Target = new TextSpanAnchor { ParaId = "auto-0000", Expect = string.Empty },
            Position = InsertPosition.After,
            Table = new TableData
            {
                Headers = new[] { "Description", "Qty", "Amount" },
                Rows = new List<IReadOnlyList<string>> { new[] { "Workshop", "2", "3600.00" } }
            }
        });

        var sized = Apply(client, withTable, new FormatOp
        {
            Target = new NodeAnchor { Kind = "table", Path = "table#0" },
            ColumnWidthsPx = new[] { 400, 80, 120 }
        });

        using var stream = new MemoryStream(sized);
        using var opened = WordprocessingDocument.Open(stream, isEditable: false);
        var table = opened.MainDocumentPart!.Document.Body!.Elements<Table>().Single();

        // 15 twips to the pixel at 96 DPI.
        Assert.Equal(new[] { "6000", "1200", "1800" },
            table.GetFirstChild<TableGrid>()!.Elements<GridColumn>()
                .Select(c => c.Width!.Value).ToArray());

        // The grid alone is a hint Word abandons the moment a cell's content is wider, so
        // every cell has to agree with it and the layout has to be fixed.
        foreach (var row in table.Elements<TableRow>())
            Assert.Equal(new[] { "6000", "1200", "1800" },
                row.Elements<TableCell>()
                    .Select(c => c.TableCellProperties!.TableCellWidth!.Width!.Value).ToArray());

        var properties = table.GetFirstChild<TableProperties>()!;
        Assert.Equal(TableLayoutValues.Fixed, properties.TableLayout!.Type!.Value);
        Assert.Equal("9000", properties.TableWidth!.Width!.Value);
        AssertValid(sized);
    }

    [Fact]
    public void Column_widths_are_refused_anywhere_but_a_table()
    {
        var client = Client();

        var report = new OfficeAgentClient(new WordModule()).Preview(
            new StreamHandle(new MemoryStream(Seeded())),
            new DocumentPlan
            {
                Operations = new PlanOperation[]
                {
                    new FormatOp
                    {
                        Target = new TextSpanAnchor { ParaId = "w14:00000002", Expect = string.Empty },
                        ColumnWidthsPx = new[] { 100, 100 }
                    }
                }
            });

        Assert.Contains("belongs on a table target", Assert.Single(report.Errors).Message);
    }

    private static byte[] Blank() => new WordModule().CreateBlank();
    private static byte[] Seeded() => DocxFactory.Contract();

    private static byte[] Apply(OfficeAgentClient client, byte[] document, PlanOperation operation)
    {
        using var applied = client.Commit(
            new StreamHandle(new MemoryStream(document)),
            new DocumentPlan { Operations = new[] { operation } });

        Assert.True(applied.Committed,
            string.Join("; ", applied.Report.Errors.Select(e => $"{e.Code}: {e.Message}")));
        return applied.ToBytes();
    }

    private static void AssertValid(byte[] document)
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
