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
