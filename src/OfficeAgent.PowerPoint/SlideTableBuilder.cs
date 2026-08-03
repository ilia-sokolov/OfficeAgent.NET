using OfficeAgent.Abstractions;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeAgent.PowerPoint;

/// <summary>
/// Builds and edits the DrawingML tables PowerPoint uses, keeping the structural
/// invariants PowerPoint relies on: every row carries one cell per grid column, and the
/// grid's column count is the table's true width.
/// </summary>
internal static class SlideTableBuilder
{
    /// <summary>Default width of the frame a new table is placed in, in EMU (10 inches).</summary>
    private const long DefaultTableWidth = 9144000L;

    /// <summary>Height allowed per row when sizing a new frame, in EMU (~0.4 inch).</summary>
    private const long RowHeight = 370840L;

    /// <summary>
    /// Builds a <c>p:graphicFrame</c> holding a table with the supplied content. Headers,
    /// when present, become the first row and the table is marked <c>firstRow</c> so the
    /// deck's table style renders them as a header band.
    /// </summary>
    public static P.GraphicFrame BuildFrame(TableData data, uint shapeId, long offsetY)
    {
        var rows = new List<IReadOnlyList<string>>();
        if (data.Headers.Count > 0) rows.Add(data.Headers);
        rows.AddRange(data.Rows);
        if (rows.Count == 0) rows.Add(Array.Empty<string>());

        var columnCount = Math.Max(1, rows.Max(r => r.Count));
        var columnWidth = DefaultTableWidth / columnCount;

        var grid = new A.TableGrid();
        for (var i = 0; i < columnCount; i++)
            grid.Append(new A.GridColumn { Width = columnWidth });

        var table = new A.Table(
            new A.TableProperties { FirstRow = data.Headers.Count > 0, BandRow = true },
            grid);

        foreach (var row in rows)
            table.Append(BuildRow(row, columnCount));

        return new P.GraphicFrame(
            new P.NonVisualGraphicFrameProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = $"Table {shapeId}" },
                new P.NonVisualGraphicFrameDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.Transform(
                new A.Offset { X = 838200L, Y = offsetY },
                new A.Extents { Cx = DefaultTableWidth, Cy = RowHeight * rows.Count }),
            new A.Graphic(new A.GraphicData(table)
            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/table" }));
    }

    /// <summary>Builds one row, padded to the table's column count.</summary>
    public static A.TableRow BuildRow(IReadOnlyList<string> cells, int columnCount)
    {
        var row = new A.TableRow { Height = RowHeight };
        for (var i = 0; i < columnCount; i++)
            row.Append(BuildCell(i < cells.Count ? cells[i] : string.Empty));
        return row;
    }

    /// <summary>
    /// Builds one cell. The order matters: <c>a:tcPr</c> must follow <c>a:txBody</c>, and
    /// PowerPoint rejects the reverse.
    /// </summary>
    public static A.TableCell BuildCell(string text) => new(
        new A.TextBody(
            new A.BodyProperties(),
            new A.ListStyle(),
            new A.Paragraph(new A.Run(
                new A.RunProperties { Language = "en-US", Dirty = false },
                new A.Text(text ?? string.Empty)))),
        new A.TableCellProperties());

    /// <summary>The table's column count, taken from the grid that defines it.</summary>
    public static int ColumnCount(A.Table table) =>
        table.TableGrid?.Elements<A.GridColumn>().Count() ?? 0;

    /// <summary>The rows of a table, materialised so callers can index and mutate safely.</summary>
    public static List<A.TableRow> Rows(A.Table table) => table.Elements<A.TableRow>().ToList();

    /// <summary>The plain text of a cell, joined across its paragraphs.</summary>
    public static string TextOf(A.TableCell cell) =>
        cell.TextBody is null
            ? string.Empty
            : string.Join(" ", cell.TextBody.Elements<A.Paragraph>().Select(PowerPointModel.TextOf));

    /// <summary>
    /// Resolves an index that may count from the end, as the shared table verbs define
    /// (-1 is the last entry). Returns -1 when the index falls outside the collection.
    /// </summary>
    public static int Normalize(int index, int count)
    {
        var resolved = index < 0 ? count + index : index;
        return resolved >= 0 && resolved < count ? resolved : -1;
    }
}
