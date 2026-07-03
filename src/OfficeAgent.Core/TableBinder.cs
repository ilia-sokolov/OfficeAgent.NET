using DocumentFormat.OpenXml;
using OfficeAgent.Abstractions;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.Core;

/// <summary>Builds a clean WordprocessingML table from typed <see cref="TableData"/>.</summary>
internal sealed class TableBinder
{
    public Table Build(TableData data)
    {
        var table = new Table();

        var borders = new TableBorders(
            new TopBorder { Val = BorderValues.Single, Size = 4 },
            new BottomBorder { Val = BorderValues.Single, Size = 4 },
            new LeftBorder { Val = BorderValues.Single, Size = 4 },
            new RightBorder { Val = BorderValues.Single, Size = 4 },
            new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
            new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 });

        var properties = new TableProperties(
            new TableStyle { Val = data.StyleId ?? "TableGrid" },
            new TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" },
            borders);

        table.AppendChild(properties);

        int columnCount = Math.Max(
            data.Headers.Count,
            data.Rows.Count > 0 ? data.Rows.Max(r => r.Count) : 0);
        if (columnCount > 0)
        {
            const int textWidthTwips = 9360; // ~US Letter with 1" margins
            var columnWidth = (textWidthTwips / columnCount).ToString();
            var grid = new TableGrid();
            for (var i = 0; i < columnCount; i++)
                grid.AppendChild(new GridColumn { Width = columnWidth });
            table.AppendChild(grid);
        }

        if (data.Headers.Count > 0)
            table.AppendChild(BuildRow(data.Headers, header: true));

        foreach (var row in data.Rows)
            table.AppendChild(BuildRow(row, header: false));

        return table;
    }

    private static TableRow BuildRow(IReadOnlyList<string> cells, bool header)
    {
        var row = new TableRow();
        foreach (var value in cells)
            row.AppendChild(BuildCell(value, header));
        return row;
    }

    private static TableCell BuildCell(string value, bool header)
    {
        var runProperties = header ? new RunProperties(new Bold()) : new RunProperties();
        var run = new Run(runProperties, new Text(value) { Space = SpaceProcessingModeValues.Preserve });
        return new TableCell(new Paragraph(run));
    }
}
