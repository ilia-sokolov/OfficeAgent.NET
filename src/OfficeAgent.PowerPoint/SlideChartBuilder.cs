using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeAgent.PowerPoint;

internal static class SlideChartBuilder
{
    public const string FrameNamePrefix = "OfficeAgent Chart ";
    private const string ChartUri = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private const string WorkbookContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    private const long EmusPerPixel = 9525L;

    public static P.GraphicFrame Add(
        SlidePart slide, uint shapeId, InsertChartOp operation)
    {
        var chartPart = slide.AddNewPart<ChartPart>();
        Write(chartPart, operation.Kind, operation.Categories, operation.Series,
            operation.Title, operation.ShowLegend);
        var relationshipId = slide.GetIdOfPart(chartPart);
        return new P.GraphicFrame(
            new P.NonVisualGraphicFrameProperties(
                new P.NonVisualDrawingProperties
                {
                    Id = shapeId,
                    Name = FrameNamePrefix + shapeId,
                    Description = operation.Description
                },
                new P.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.Transform(
                new A.Offset { X = operation.XPx * EmusPerPixel, Y = operation.YPx * EmusPerPixel },
                new A.Extents { Cx = operation.WidthPx * EmusPerPixel, Cy = operation.HeightPx * EmusPerPixel }),
            new A.Graphic(new A.GraphicData(
                new C.ChartReference { Id = relationshipId }) { Uri = ChartUri }));
    }

    public static void Write(
        ChartPart chartPart,
        ChartKind kind,
        IReadOnlyList<string> categories,
        IReadOnlyList<ChartSeries> series,
        string? title,
        bool showLegend)
    {
        foreach (var embedded in chartPart.Parts.Select(pair => pair.OpenXmlPart).OfType<EmbeddedPackagePart>().ToList())
            chartPart.DeletePart(embedded);
        var workbook = chartPart.AddEmbeddedPackagePart(WorkbookContentType);
        using (var output = workbook.GetStream(FileMode.Create, FileAccess.Write))
        {
            var bytes = SpreadsheetPartUtility.CreateChartWorkbook(categories, series);
            output.Write(bytes, 0, bytes.Length);
        }

        var workbookId = chartPart.GetIdOfPart(workbook);
        var plotArea = new C.PlotArea(new C.Layout());
        if (kind == ChartKind.Pie)
            plotArea.Append(BuildPie(categories, series));
        else
        {
            const uint categoryAxisId = 48650112U;
            const uint valueAxisId = 48672768U;
            plotArea.Append(kind == ChartKind.Line
                ? BuildLine(categories, series, categoryAxisId, valueAxisId)
                : BuildBar(kind, categories, series, categoryAxisId, valueAxisId));
            plotArea.Append(BuildCategoryAxis(categoryAxisId, valueAxisId,
                kind == ChartKind.Bar ? C.AxisPositionValues.Left : C.AxisPositionValues.Bottom));
            plotArea.Append(BuildValueAxis(valueAxisId, categoryAxisId,
                kind == ChartKind.Bar ? C.AxisPositionValues.Bottom : C.AxisPositionValues.Left));
        }

        var chart = new C.Chart();
        if (!string.IsNullOrWhiteSpace(title)) chart.Append(BuildTitle(title!));
        chart.Append(plotArea);
        if (showLegend)
            chart.Append(new C.Legend(
                new C.LegendPosition { Val = C.LegendPositionValues.Right },
                new C.Layout(),
                new C.Overlay { Val = false }));
        chart.Append(new C.PlotVisibleOnly { Val = true }, new C.DisplayBlanksAs { Val = C.DisplayBlanksAsValues.Gap });
        chartPart.ChartSpace = new C.ChartSpace(
            new C.EditingLanguage { Val = "en-US" },
            chart,
            new C.ExternalData(new C.AutoUpdate { Val = false }) { Id = workbookId });
        chartPart.ChartSpace.Save();
    }

    private static C.BarChart BuildBar(
        ChartKind kind, IReadOnlyList<string> categories, IReadOnlyList<ChartSeries> series,
        uint categoryAxisId, uint valueAxisId)
    {
        var chart = new C.BarChart(
            new C.BarDirection { Val = kind == ChartKind.Bar ? C.BarDirectionValues.Bar : C.BarDirectionValues.Column },
            new C.BarGrouping { Val = C.BarGroupingValues.Clustered },
            new C.VaryColors { Val = false });
        for (var i = 0; i < series.Count; i++) chart.Append(BarSeries(i, categories, series[i]));
        chart.Append(new C.DataLabels(new C.ShowLegendKey { Val = false }, new C.ShowValue { Val = false },
            new C.ShowCategoryName { Val = false }, new C.ShowSeriesName { Val = false },
            new C.ShowPercent { Val = false }, new C.ShowBubbleSize { Val = false }));
        chart.Append(new C.GapWidth { Val = (ushort)150 }, new C.Overlap { Val = 0 },
            new C.AxisId { Val = categoryAxisId }, new C.AxisId { Val = valueAxisId });
        return chart;
    }

    private static C.LineChart BuildLine(
        IReadOnlyList<string> categories, IReadOnlyList<ChartSeries> series,
        uint categoryAxisId, uint valueAxisId)
    {
        var chart = new C.LineChart(new C.Grouping { Val = C.GroupingValues.Standard }, new C.VaryColors { Val = false });
        for (var i = 0; i < series.Count; i++) chart.Append(LineSeries(i, categories, series[i]));
        chart.Append(new C.DataLabels(new C.ShowLegendKey { Val = false }, new C.ShowValue { Val = false },
            new C.ShowCategoryName { Val = false }, new C.ShowSeriesName { Val = false },
            new C.ShowPercent { Val = false }, new C.ShowBubbleSize { Val = false }));
        chart.Append(new C.ShowMarker { Val = true }, new C.Smooth { Val = false },
            new C.AxisId { Val = categoryAxisId }, new C.AxisId { Val = valueAxisId });
        return chart;
    }

    private static C.PieChart BuildPie(IReadOnlyList<string> categories, IReadOnlyList<ChartSeries> series)
    {
        var chart = new C.PieChart(new C.VaryColors { Val = true });
        for (var i = 0; i < series.Count; i++) chart.Append(PieSeries(i, categories, series[i]));
        chart.Append(new C.DataLabels(new C.ShowLegendKey { Val = false }, new C.ShowValue { Val = false },
            new C.ShowCategoryName { Val = false }, new C.ShowSeriesName { Val = false },
            new C.ShowPercent { Val = true }, new C.ShowLeaderLines { Val = true }));
        chart.Append(new C.FirstSliceAngle { Val = (ushort)0 });
        return chart;
    }

    private static C.BarChartSeries BarSeries(int index, IReadOnlyList<string> categories, ChartSeries series) => new(
        new C.Index { Val = (uint)index }, new C.Order { Val = (uint)index },
        SeriesText(index, series.Name), new C.InvertIfNegative { Val = false },
        Categories(categories), Values(index, series.Values));

    private static C.LineChartSeries LineSeries(int index, IReadOnlyList<string> categories, ChartSeries series) => new(
        new C.Index { Val = (uint)index }, new C.Order { Val = (uint)index },
        SeriesText(index, series.Name),
        new C.Marker(new C.Symbol { Val = C.MarkerStyleValues.Circle }),
        Categories(categories), Values(index, series.Values), new C.Smooth { Val = false });

    private static C.PieChartSeries PieSeries(int index, IReadOnlyList<string> categories, ChartSeries series) => new(
        new C.Index { Val = (uint)index }, new C.Order { Val = (uint)index },
        SeriesText(index, series.Name), Categories(categories), Values(index, series.Values));

    private static C.SeriesText SeriesText(int index, string name)
    {
        var column = SpreadsheetPartUtility.ColumnName(index + 2);
        return new C.SeriesText(new C.StringReference(
            new C.Formula($"ChartData!${column}$1"),
            new C.StringCache(new C.PointCount { Val = 1U },
                new C.StringPoint(new C.NumericValue(name)) { Index = 0U })));
    }

    private static C.CategoryAxisData Categories(IReadOnlyList<string> categories)
    {
        var cache = new C.StringCache(new C.PointCount { Val = (uint)categories.Count });
        for (var i = 0; i < categories.Count; i++)
            cache.Append(new C.StringPoint(new C.NumericValue(categories[i])) { Index = (uint)i });
        return new C.CategoryAxisData(new C.StringReference(
            new C.Formula($"ChartData!$A$2:$A${categories.Count + 1}"), cache));
    }

    private static C.Values Values(int index, IReadOnlyList<double?> values)
    {
        var cache = new C.NumberingCache(new C.FormatCode("General"), new C.PointCount { Val = (uint)values.Count });
        for (var i = 0; i < values.Count; i++)
            if (values[i] is { } value)
                cache.Append(new C.NumericPoint(new C.NumericValue(value.ToString("R", CultureInfo.InvariantCulture))) { Index = (uint)i });
        var column = SpreadsheetPartUtility.ColumnName(index + 2);
        return new C.Values(new C.NumberReference(
            new C.Formula($"ChartData!${column}$2:${column}${values.Count + 1}"), cache));
    }

    private static C.CategoryAxis BuildCategoryAxis(uint id, uint crossingId, C.AxisPositionValues position) => new(
        new C.AxisId { Val = id },
        new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
        new C.Delete { Val = false }, new C.AxisPosition { Val = position },
        new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
        new C.CrossingAxis { Val = crossingId }, new C.Crosses { Val = C.CrossesValues.AutoZero },
        new C.AutoLabeled { Val = true }, new C.LabelAlignment { Val = C.LabelAlignmentValues.Center },
        new C.LabelOffset { Val = (ushort)100 });

    private static C.ValueAxis BuildValueAxis(uint id, uint crossingId, C.AxisPositionValues position) => new(
        new C.AxisId { Val = id },
        new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
        new C.Delete { Val = false }, new C.AxisPosition { Val = position },
        new C.MajorGridlines(), new C.NumberingFormat { FormatCode = "General", SourceLinked = true },
        new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
        new C.CrossingAxis { Val = crossingId }, new C.Crosses { Val = C.CrossesValues.AutoZero },
        new C.CrossBetween { Val = C.CrossBetweenValues.Between });

    private static C.Title BuildTitle(string title) => new(
        new C.ChartText(new C.RichText(
            new A.BodyProperties(), new A.ListStyle(),
            new A.Paragraph(new A.Run(new A.RunProperties { Language = "en-US" }, new A.Text(title))))),
        new C.Layout(), new C.Overlay { Val = false });
}
