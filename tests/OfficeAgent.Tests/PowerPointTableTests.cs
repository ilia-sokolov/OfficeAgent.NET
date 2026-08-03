using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.PowerPoint;
using DocFormat = OfficeAgent.Abstractions.DocumentFormat;

namespace OfficeAgent.Tests;

/// <summary>
/// Table verbs on a slide: tables surface as durable, shape-scoped nodes, and the shared
/// row/column vocabulary edits them while keeping the grid and the rows in agreement -
/// the invariant PowerPoint relies on to render a table at all.
/// </summary>
public class PowerPointTableTests
{
    private static OfficeAgentClient Client() => new(new PowerPointModule());

    [Fact]
    public void Tables_surface_as_shape_scoped_nodes()
    {
        var inspection = Client().Inspect(PptxFactory.DeckWithTable());

        var table = Assert.Single(inspection.Nodes, n => n.Kind == "table");
        // Shape-scoped rather than ordinal, so adding a table to an earlier slide cannot
        // silently retarget this path.
        Assert.Equal("table#256/3", table.Path);
        Assert.Contains("2×2", table.Summary);

        // Cell text is addressable as ordinary paragraphs.
        Assert.Contains(inspection.Paragraphs, p => p.Text == "Region" && p.ParaId.Contains("r0c0"));
        Assert.Contains(inspection.Paragraphs, p => p.Text == "41850" && p.ParaId.Contains("r1c1"));
    }

    [Fact]
    public void Insert_table_adds_one_to_a_slide()
    {
        var client = Client();
        var deck = PptxFactory.Deck();

        var applied = Apply(client, deck, new InsertTableOp
        {
            Target = new NodeAnchor { Kind = "slide", Path = "slide#256" },
            Table = new TableData
            {
                Headers = new[] { "Region", "Q1" },
                Rows = new[] { new[] { "EMEA", "41850" }, new[] { "APAC", "22100" } }
            }
        });

        var inspection = client.Inspect(applied);
        var table = Assert.Single(inspection.Nodes, n => n.Kind == "table");
        Assert.Contains("3×2", table.Summary);

        // The header must be the FIRST row, not merely present somewhere: a table whose
        // headers land at the bottom still reports 3×2 and still contains "EMEA".
        var (headers, _) = HeaderRowWithWidths(applied);
        Assert.Equal(new[] { "Region", "Q1" }, headers);
        Assert.Contains(inspection.Paragraphs, p => p.Text == "EMEA");
        AssertValid(applied);
    }

    [Fact]
    public void Insert_table_pads_short_rows_to_the_grid_width()
    {
        var client = Client();

        var applied = Apply(client, PptxFactory.Deck(), new InsertTableOp
        {
            Target = new NodeAnchor { Kind = "slide", Path = "slide#256" },
            Table = new TableData
            {
                Headers = new[] { "A", "B", "C" },
                // Deliberately short: every row must still carry three cells or
                // PowerPoint refuses to render the table.
                Rows = new[] { new[] { "1" } }
            }
        });

        // The widest row dictates the grid: narrowing it to the shortest would silently
        // drop headers B and C rather than pad the short row.
        var (headers, _) = HeaderRowWithWidths(applied);
        Assert.Equal(new[] { "A", "B", "C" }, headers);
        Assert.Equal(new[] { "1", string.Empty, string.Empty }, RowTexts(applied, 1));

        AssertValid(applied);
        AssertGridMatchesRows(applied);
    }

    [Fact]
    public void Rows_can_be_appended_and_inserted_at_a_position()
    {
        var client = Client();
        var anchor = new NodeAnchor { Kind = "table", Path = "table#256/3" };

        var appended = Apply(client, PptxFactory.DeckWithTable(), new InsertTableRowsOp
        {
            Target = anchor,
            Rows = new[] { new[] { "APAC", "22100" } }
        });
        Assert.Contains("3×2", Assert.Single(client.Inspect(appended).Nodes, n => n.Kind == "table").Summary);
        // Appended means last: a prepend also yields 3×2.
        Assert.Equal(new[] { "APAC", "22100" }, RowTexts(appended, 2));

        var atStart = Apply(client, PptxFactory.DeckWithTable(), new InsertTableRowsOp
        {
            Target = anchor,
            Rows = new[] { new[] { "First", "Row" } },
            Position = TablePosition.Start
        });

        // Position must be honoured, not merely accepted.
        var texts = client.Inspect(atStart).Paragraphs
            .Where(p => p.ParaId.Contains("r0c0")).Select(p => p.Text).ToList();
        Assert.Contains("First", texts);
        AssertValid(atStart);
        AssertGridMatchesRows(atStart);
    }

    [Fact]
    public void Rows_insert_after_an_index_in_the_order_supplied()
    {
        var client = Client();

        var applied = Apply(client, PptxFactory.DeckWithTable(), new InsertTableRowsOp
        {
            Target = new NodeAnchor { Kind = "table", Path = "table#256/3" },
            Rows = new[] { new[] { "One", "1" }, new[] { "Two", "2" } },
            Position = TablePosition.After,
            RowIndex = 0
        });

        // Both land after the header, and "One" precedes "Two".
        var inspection = client.Inspect(applied);
        Assert.Equal("One", Assert.Single(inspection.Paragraphs, p => p.ParaId.Contains("r1c0")).Text);
        Assert.Equal("Two", Assert.Single(inspection.Paragraphs, p => p.ParaId.Contains("r2c0")).Text);
        AssertGridMatchesRows(applied);
    }

    [Fact]
    public void Columns_can_be_added_and_removed_without_desyncing_the_grid()
    {
        var client = Client();
        var anchor = new NodeAnchor { Kind = "table", Path = "table#256/3" };

        var widened = Apply(client, PptxFactory.DeckWithTable(), new InsertTableColumnsOp
        {
            Target = anchor,
            Columns = new[] { new[] { "Growth", "12%" } }
        });
        Assert.Contains("2×3", Assert.Single(client.Inspect(widened).Nodes, n => n.Kind == "table").Summary);
        // Default position is rightmost: prepending instead also yields 2×3.
        Assert.Equal(new[] { "Region", "Revenue", "Growth" }, HeaderRowWithWidths(widened).Headers);
        AssertGridMatchesRows(widened);
        AssertValid(widened);

        var narrowed = Apply(client, PptxFactory.DeckWithTable(), new RemoveTableColumnsOp
        {
            Target = anchor,
            ColumnIndices = new[] { -1 }
        });
        Assert.Contains("2×1", Assert.Single(client.Inspect(narrowed).Nodes, n => n.Kind == "table").Summary);
        // -1 is the LAST column: an index that resolved to 0 would drop "Region" instead
        // and still leave a 2×1 table.
        Assert.Equal(new[] { "Region" }, HeaderRowWithWidths(narrowed).Headers);
        AssertGridMatchesRows(narrowed);
        AssertValid(narrowed);
    }

    [Fact]
    public void Inserting_several_columns_keeps_each_width_with_its_own_data()
    {
        var client = Client();

        // Two columns inserted after the first. The fixture's columns have deliberately
        // different widths, so a grid that drifts out of step with the cells shows up as
        // a width attached to the wrong column rather than passing unnoticed.
        var applied = Apply(client, PptxFactory.DeckWithTable(), new InsertTableColumnsOp
        {
            Target = new NodeAnchor { Kind = "table", Path = "table#256/3" },
            Columns = new[] { new[] { "N1h", "N1b" }, new[] { "N2h", "N2b" } },
            Position = TablePosition.After,
            ColumnIndex = 0
        });

        var (headers, widths) = HeaderRowWithWidths(applied);

        Assert.Equal(new[] { "Region", "N1h", "N2h", "Revenue" }, headers);
        // "Revenue" started in the narrow column and must still be in it.
        Assert.Equal(1000000L, widths[Array.IndexOf(headers, "Revenue")]);
        Assert.Equal(3000000L, widths[Array.IndexOf(headers, "Region")]);
        AssertGridMatchesRows(applied);
        AssertValid(applied);
    }

    [Fact]
    public void Empty_rows_are_removed_only_when_asked()
    {
        var client = Client();
        var anchor = new NodeAnchor { Kind = "table", Path = "table#256/3" };

        var withBlank = Apply(client, PptxFactory.DeckWithTable(), new InsertTableRowsOp
        {
            Target = anchor,
            Rows = new[] { new[] { "", "" } }
        });

        var cleaned = Apply(client, withBlank, new RemoveTableRowsOp { Target = anchor, OnlyIfEmpty = true });

        Assert.Contains("2×2", Assert.Single(client.Inspect(cleaned).Nodes, n => n.Kind == "table").Summary);
        AssertGridMatchesRows(cleaned);

        // …and without the flag, no indices means no removal. Ignoring the flag would
        // strip blank rows nobody asked to lose.
        var untouched = Apply(client, withBlank, new RemoveTableRowsOp { Target = anchor });
        Assert.Contains("3×2", Assert.Single(client.Inspect(untouched).Nodes, n => n.Kind == "table").Summary);
    }

    [Fact]
    public void Remove_table_takes_the_frame_with_it()
    {
        var client = Client();

        var applied = Apply(client, PptxFactory.DeckWithTable(), new RemoveTableOp
        {
            Target = new NodeAnchor { Kind = "table", Path = "table#256/3" }
        });

        var inspection = client.Inspect(applied);
        Assert.DoesNotContain(inspection.Nodes, n => n.Kind == "table");
        // Removing only the a:tbl would leave an empty p:graphicFrame behind, which still
        // reports no table node but leaves a stray shape on the slide.
        Assert.Empty(GraphicFrames(applied));
        // The title shape is untouched: only the table's own frame went.
        Assert.Contains(inspection.Paragraphs, p => p.Text == PptxFactory.TitleText);
        AssertValid(applied);
    }

    [Fact]
    public void Emptying_a_table_completely_is_refused_with_a_usable_alternative()
    {
        var client = Client();
        var anchor = new NodeAnchor { Kind = "table", Path = "table#256/3" };

        var rows = Preview(client, PptxFactory.DeckWithTable(), new RemoveTableRowsOp
        {
            Target = anchor,
            RowIndices = new[] { 0, 1 }
        });
        var columns = Preview(client, PptxFactory.DeckWithTable(), new RemoveTableColumnsOp
        {
            Target = anchor,
            ColumnIndices = new[] { 0, 1 }
        });

        Assert.False(rows.IsValid);
        Assert.Contains("removeTable", Assert.Single(rows.Errors).Message);
        Assert.False(columns.IsValid);
        Assert.Contains("removeTable", Assert.Single(columns.Errors).Message);
    }

    [Fact]
    public void Out_of_range_indices_and_missing_tables_are_reported_not_applied()
    {
        var client = Client();

        var badIndex = Preview(client, PptxFactory.DeckWithTable(), new RemoveTableRowsOp
        {
            Target = new NodeAnchor { Kind = "table", Path = "table#256/3" },
            RowIndices = new[] { 99 }
        });
        var missing = Preview(client, PptxFactory.DeckWithTable(), new RemoveTableOp
        {
            Target = new NodeAnchor { Kind = "table", Path = "table#256/999" }
        });

        Assert.Equal(ValidationErrorCodes.InvalidOperation, Assert.Single(badIndex.Errors).Code);
        Assert.Equal(ValidationErrorCodes.AnchorNotFound, Assert.Single(missing.Errors).Code);
    }

    // ── helpers ───────────────────────────────────────────────────────────────

    private static byte[] Apply(OfficeAgentClient client, byte[] deck, PlanOperation operation)
    {
        using var applied = client.Commit(new StreamHandle(new MemoryStream(deck)), Plan(operation));
        Assert.True(applied.Committed,
            string.Join("; ", applied.Report.Errors.Select(e => $"{e.Code}: {e.Message}")));
        return applied.ToBytes();
    }

    private static ChangeReport Preview(OfficeAgentClient client, byte[] deck, PlanOperation operation) =>
        client.Preview(new StreamHandle(new MemoryStream(deck)), Plan(operation));

    private static DocumentPlan Plan(PlanOperation operation) => new()
    {
        Format = DocFormat.PowerPoint,
        Operations = new[] { operation }
    };

    private static void AssertValid(byte[] deck)
    {
        using var stream = new MemoryStream(deck);
        using var document = PresentationDocument.Open(stream, isEditable: false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(document).ToList();
        Assert.True(errors.Count == 0,
            string.Join("; ", errors.Select(e => $"{e.Path?.XPath}: {e.Description}")));
    }

    /// <summary>The cell texts of one row, by index, for asserting on placement order.</summary>
    private static string[] RowTexts(byte[] deck, int rowIndex)
    {
        using var stream = new MemoryStream(deck);
        using var document = PresentationDocument.Open(stream, isEditable: false);

        return document.PresentationPart!.SlideParts
            .SelectMany(p => p.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Table>())
            .Single()
            .Elements<DocumentFormat.OpenXml.Drawing.TableRow>()
            .ElementAt(rowIndex)
            .Elements<DocumentFormat.OpenXml.Drawing.TableCell>()
            .Select(c => string.Concat(c.TextBody!.Descendants<DocumentFormat.OpenXml.Drawing.Text>()
                .Select(t => t.Text)))
            .ToArray();
    }

    /// <summary>The graphic frames left on the deck's slides.</summary>
    private static List<DocumentFormat.OpenXml.Presentation.GraphicFrame> GraphicFrames(byte[] deck)
    {
        using var stream = new MemoryStream(deck);
        using var document = PresentationDocument.Open(stream, isEditable: false);

        return document.PresentationPart!.SlideParts
            .SelectMany(p => p.Slide.Descendants<DocumentFormat.OpenXml.Presentation.GraphicFrame>())
            .ToList();
    }

    /// <summary>
    /// The header row's cell texts paired with the grid width sitting at each position,
    /// so a test can assert that a column's width still belongs to that column's data.
    /// </summary>
    private static (string[] Headers, long[] Widths) HeaderRowWithWidths(byte[] deck)
    {
        using var stream = new MemoryStream(deck);
        using var document = PresentationDocument.Open(stream, isEditable: false);

        var table = document.PresentationPart!.SlideParts
            .SelectMany(p => p.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Table>())
            .Single();

        var headers = table.Elements<DocumentFormat.OpenXml.Drawing.TableRow>().First()
            .Elements<DocumentFormat.OpenXml.Drawing.TableCell>()
            .Select(c => string.Concat(c.TextBody!.Descendants<DocumentFormat.OpenXml.Drawing.Text>()
                .Select(t => t.Text)))
            .ToArray();

        var widths = table.TableGrid!.Elements<DocumentFormat.OpenXml.Drawing.GridColumn>()
            .Select(c => c.Width!.Value)
            .ToArray();

        return (headers, widths);
    }

    /// <summary>
    /// Every row must carry exactly one cell per grid column. Schema validation does not
    /// catch a mismatch, but PowerPoint renders the table wrongly or not at all.
    /// </summary>
    private static void AssertGridMatchesRows(byte[] deck)
    {
        using var stream = new MemoryStream(deck);
        using var document = PresentationDocument.Open(stream, isEditable: false);

        foreach (var part in document.PresentationPart!.SlideParts)
            foreach (var table in part.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Table>())
            {
                var columns = table.TableGrid!.Elements<DocumentFormat.OpenXml.Drawing.GridColumn>().Count();
                foreach (var row in table.Elements<DocumentFormat.OpenXml.Drawing.TableRow>())
                    Assert.Equal(columns, row.Elements<DocumentFormat.OpenXml.Drawing.TableCell>().Count());
            }
    }
}
