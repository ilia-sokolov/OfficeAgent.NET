using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeAgent.PowerPoint;

/// <summary>Replaces every cell of a slide table, growing or shrinking it by rows, and keeps each cell's look.</summary>
internal sealed class SlideFillTableHandler : IOperationHandler
{
    /// <inheritdoc />
    public bool CanHandle(PlanOperation operation) =>
        operation is FillTableOp { Target: NodeAnchor { Kind: "table" } };

    /// <inheritdoc />
    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (FillTableOp)operation;
        var anchor = (NodeAnchor)op.Target;
        var table = SlideTableNodeProvider.Locate(anchor.Path, new PowerPointObjectMap(context.Package));
        if (table is null)
            return Fail(ValidationErrorCodes.AnchorNotFound, $"No table with path '{anchor.Path}'.", anchor);

        var rows = SlideTableBuilder.Rows(table.Table);
        var columns = SlideTableBuilder.ColumnCount(table.Table);
        if (TableFill.Check(op.Rows, columns, rows.Count > 0, Merged(table.Table)) is { } error)
            return Fail(ValidationErrorCodes.InvalidOperation, error, anchor);

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "fillTable",
            Before = $"[table {rows.Count}×{columns}]",
            After = TableFill.Describe(op.Rows),
            Context = $"slide {table.Slide.Number}",
            BlastRadius = op.Rows.Count * columns
        });
    }

    /// <inheritdoc />
    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (FillTableOp)operation;
        var anchor = (NodeAnchor)op.Target;
        var table = SlideTableNodeProvider.Locate(anchor.Path, new PowerPointObjectMap(context.Package))
            ?? throw new InvalidOperationException($"Table '{anchor.Path}' vanished before apply.");

        var rows = SlideTableBuilder.Rows(table.Table);
        while (rows.Count < op.Rows.Count)
        {
            var copy = (A.TableRow)rows[rows.Count - 1].CloneNode(deep: true);
            rows[rows.Count - 1].InsertAfterSelf(copy);
            rows.Add(copy);
        }

        while (rows.Count > op.Rows.Count)
        {
            rows[rows.Count - 1].Remove();
            rows.RemoveAt(rows.Count - 1);
        }

        for (var r = 0; r < rows.Count; r++)
        {
            var cells = rows[r].Elements<A.TableCell>().ToList();
            for (var c = 0; c < cells.Count; c++)
                SetText(cells[c], op.Rows[r][c]);
        }
    }

    /// <summary>
    /// Writes a cell's text into its first paragraph with the first run's formatting, and drops
    /// whatever else the cell held.
    /// </summary>
    private static void SetText(A.TableCell cell, string text)
    {
        var body = cell.TextBody ?? cell.PrependChild(new A.TextBody(new A.BodyProperties(), new A.ListStyle()));
        var paragraphs = body.Elements<A.Paragraph>().ToList();
        var first = paragraphs.FirstOrDefault() ?? body.AppendChild(new A.Paragraph());
        foreach (var extra in paragraphs.Skip(1)) extra.Remove();

        var pattern = first.Elements<A.Run>().FirstOrDefault();
        var properties = pattern?.RunProperties is { } rp
            ? (A.RunProperties)rp.CloneNode(deep: true)
            : new A.RunProperties { Language = "en-US" };
        foreach (var child in first.ChildElements.Where(child => child is A.Run or A.Field or A.Break).ToList())
            child.Remove();

        if (text.Length == 0) return;
        var run = new A.Run(properties, new A.Text(text));
        if (first.GetFirstChild<A.EndParagraphRunProperties>() is { } end) end.InsertBeforeSelf(run);
        else first.AppendChild(run);
    }

    private static bool Merged(A.Table table) =>
        table.Descendants<A.TableCell>().Any(cell =>
            (cell.GridSpan?.Value ?? 1) > 1 || (cell.RowSpan?.Value ?? 1) > 1 ||
            cell.HorizontalMerge?.Value == true || cell.VerticalMerge?.Value == true);

    private static OperationPreview Fail(string code, string message, Anchor anchor) =>
        OperationPreview.Fail(new ValidationError(code, message, anchor));
}

/// <summary>Replaces a chart's data while keeping its type, colours, labels and layout.</summary>
internal sealed class SlideSetChartDataHandler : IOperationHandler
{
    /// <inheritdoc />
    public bool CanHandle(PlanOperation operation) =>
        operation is SetChartDataOp { Target: NodeAnchor { Kind: "chart" } };

    /// <inheritdoc />
    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (SetChartDataOp)operation;
        var anchor = (NodeAnchor)op.Target;
        var located = SlideChartNodeProvider.Locate(anchor.Path, context.Package);
        if (located is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound, $"No chart with path '{anchor.Path}'.", anchor));

        if (ChartDataWriter.Check(located.Part, op.Categories, op.Series) is { } error)
            return OperationPreview.Fail(new ValidationError(ValidationErrorCodes.InvalidOperation, error, anchor));

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "setChartData",
            Before = $"[chart: {ChartDataWriter.SeriesCount(located.Part)} series]",
            After = $"[chart: {op.Categories.Count} categories, {op.Series.Count} series: {string.Join(", ", op.Series.Select(s => s.Name))}]",
            Context = $"slide {located.Slide.Number}",
            BlastRadius = op.Series.Count
        });
    }

    /// <inheritdoc />
    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (SetChartDataOp)operation;
        var anchor = (NodeAnchor)op.Target;
        var located = SlideChartNodeProvider.Locate(anchor.Path, context.Package)
            ?? throw new InvalidOperationException($"Chart '{anchor.Path}' vanished before apply.");
        ChartDataWriter.Write(located.Part, op.Categories, op.Series);
    }
}

/// <summary>Names a slide shape, table or chart with a data binding marker.</summary>
internal sealed class SlideBindDataHandler : IOperationHandler
{
    /// <inheritdoc />
    public bool CanHandle(PlanOperation operation) =>
        operation is BindDataOp { Target: NodeAnchor { Kind: "shape" or "table" or "chart" } };

    /// <inheritdoc />
    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (BindDataOp)operation;
        var anchor = (NodeAnchor)op.Target;
        if (!WorkbookReference.TryParse(op.Source, out _))
            return OperationPreview.Fail(new ValidationError(ValidationErrorCodes.InvalidOperation,
                $"'{op.Source}' is not a workbook range or name. Use a reference such as Summary!B2 or 'Q3 Sales'!A1:D6, " +
                "or a defined name or Excel table name.", anchor));

        var located = Locate(context.Package, anchor);
        if (located is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound, $"No {anchor.Kind} with path '{anchor.Path}'.", anchor));

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "bindData",
            Before = PowerPointModel.ShapeNameOf(located.Value.Element),
            After = DataBindingMarker.For(op.Source.Trim()),
            Context = $"slide {located.Value.Slide.Number}",
            BlastRadius = 1
        });
    }

    /// <inheritdoc />
    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (BindDataOp)operation;
        var anchor = (NodeAnchor)op.Target;
        var (slide, element) = Locate(context.Package, anchor)
            ?? throw new InvalidOperationException($"'{anchor.Path}' vanished before apply.");

        var marker = DataBindingMarker.For(op.Source.Trim());
        var current = PowerPointModel.ShapeNameOf(element);
        if (DataBindingMarker.TryParse(current, out var existing) &&
            string.Equals(existing, op.Source.Trim(), StringComparison.OrdinalIgnoreCase)) return;

        // Fill finds a shape by name, so two shapes named alike on one slide could not both
        // be refreshed. A suffix after '#' tells them apart and is ignored when read back.
        var taken = new HashSet<string>(
            slide.Part.SlideRoot().CommonSlideData?.ShapeTree?.ChildElements
                .Where(other => !ReferenceEquals(other, element))
                .Select(PowerPointModel.ShapeNameOf) ?? Enumerable.Empty<string>(),
            StringComparer.OrdinalIgnoreCase);
        var name = marker;
        for (var n = 2; taken.Contains(name); n++) name = $"{marker}#{n}";

        if (DrawingPropertiesOf(element) is { } properties) properties.Name = name;
    }

    private static (SlideRef Slide, OpenXmlElement Element)? Locate(IOpenXmlPackage package, NodeAnchor anchor)
    {
        switch (anchor.Kind)
        {
            case "shape":
                var shape = ShapeNodeProvider.Locate(anchor.Path, package);
                return shape is null ? null : (shape.Slide, shape.Element);
            case "table":
                var table = SlideTableNodeProvider.Locate(anchor.Path, new PowerPointObjectMap(package));
                return table is null ? null : (table.Slide, table.Frame);
            default:
                var chart = SlideChartNodeProvider.Locate(anchor.Path, package);
                return chart is null ? null : (chart.Slide, chart.Frame);
        }
    }

    private static P.NonVisualDrawingProperties? DrawingPropertiesOf(OpenXmlElement element) => element switch
    {
        P.Shape s => s.NonVisualShapeProperties?.NonVisualDrawingProperties,
        P.GraphicFrame f => f.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties,
        P.Picture p => p.NonVisualPictureProperties?.NonVisualDrawingProperties,
        P.GroupShape g => g.NonVisualGroupShapeProperties?.NonVisualDrawingProperties,
        P.ConnectionShape c => c.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties,
        _ => null
    };
}

/// <summary>
/// Rewrites the data inside an existing chart part: series names, categories and values, and
/// the embedded workbook behind them. Every element that is not data - series colours and
/// markers, per-point overrides, labels, axes, titles, layout - is left where it is.
/// </summary>
internal static class ChartDataWriter
{
    private const string WorkbookContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    private static bool IsSupported(OpenXmlElement group) =>
        group is C.BarChart or C.Bar3DChart or C.LineChart or C.Line3DChart or C.PieChart or C.Pie3DChart
            or C.DoughnutChart or C.OfPieChart or C.AreaChart or C.Area3DChart or C.RadarChart;

    private static bool IsUnsupported(OpenXmlElement group) =>
        group is C.ScatterChart or C.BubbleChart or C.SurfaceChart or C.Surface3DChart or C.StockChart;

    private static bool IsPie(OpenXmlElement group) => group is C.PieChart or C.Pie3DChart or C.OfPieChart;

    public static int SeriesCount(ChartPart part) => Series(part).Count;

    public static string? Check(ChartPart part, IReadOnlyList<string> categories, IReadOnlyList<ChartSeries> series)
    {
        var plotArea = part.ChartSpace?.GetFirstChild<C.Chart>()?.PlotArea;
        if (plotArea is null) return "The chart has no plot area to write data into.";
        var groups = plotArea.ChildElements.ToList();
        if (groups.Any(IsUnsupported))
            return "Scatter, bubble, surface and stock charts take numbers on both axes, not category labels, so setChartData does not write them.";
        if (!groups.Any(IsSupported)) return "The chart has no series setChartData can write.";
        if (categories.Count == 0) return "A chart needs at least one category.";
        if (series.Count == 0) return "A chart needs at least one series.";
        if (series.Any(item => item.Values.Count != categories.Count))
            return "Every series must have one value per category.";
        if (series.SelectMany(item => item.Values).Any(value => value is { } number && (double.IsNaN(number) || double.IsInfinity(number))))
            return "Chart values must be finite numbers.";
        if (groups.Where(IsSupported).All(IsPie) && series.Count != 1)
            return $"A pie chart shows exactly one series, and the data has {series.Count}.";
        return null;
    }

    public static void Write(ChartPart part, IReadOnlyList<string> categories, IReadOnlyList<ChartSeries> series)
    {
        var existing = Series(part);

        // Series beyond the new count go; missing ones copy the last series' formatting.
        for (var i = existing.Count - 1; i >= series.Count; i--)
        {
            existing[i].Remove();
            existing.RemoveAt(i);
        }

        var nextIndex = existing.Select(ser => Number(ser.GetFirstChild<C.Index>()?.Val)).DefaultIfEmpty(0U).Max() + 1;
        var nextOrder = existing.Select(ser => Number(ser.GetFirstChild<C.Order>()?.Val)).DefaultIfEmpty(0U).Max() + 1;
        while (existing.Count < series.Count)
        {
            var last = existing[existing.Count - 1];
            var copy = (OpenXmlCompositeElement)last.CloneNode(deep: true);
            // An extension list can carry a series' unique id; two series sharing one is a
            // chart PowerPoint repairs on open.
            foreach (var extensions in copy.Elements<C.ExtensionList>().ToList()) extensions.Remove();
            if (copy.GetFirstChild<C.Index>() is { } index) index.Val = nextIndex++;
            if (copy.GetFirstChild<C.Order>() is { } order) order.Val = nextOrder++;
            last.InsertAfterSelf(copy);
            existing.Add(copy);
        }

        for (var i = 0; i < series.Count; i++)
            WriteSeries(existing[i], i, categories, series[i]);

        ReplaceWorkbook(part, categories, series);
        part.ChartSpace!.Save();
    }

    /// <summary>Every series element in the plot area's chart groups, in document order.</summary>
    private static List<OpenXmlCompositeElement> Series(ChartPart part) =>
        part.ChartSpace?.GetFirstChild<C.Chart>()?.PlotArea?.ChildElements
            .Where(IsSupported)
            .SelectMany(group => group.ChildElements.OfType<OpenXmlCompositeElement>().Where(child => child.LocalName == "ser"))
            .ToList()
        ?? new List<OpenXmlCompositeElement>();

    private static void WriteSeries(OpenXmlCompositeElement ser, int index, IReadOnlyList<string> categories, ChartSeries data)
    {
        var column = SpreadsheetPartUtility.ColumnName(index + 2);
        var last = categories.Count + 1;

        var text = new C.SeriesText(new C.StringReference(
            new C.Formula($"ChartData!${column}$1"),
            new C.StringCache(new C.PointCount { Val = 1U }, new C.StringPoint(new C.NumericValue(data.Name)) { Index = 0U })));
        if (ser.GetFirstChild<C.SeriesText>() is { } oldText)
        {
            oldText.InsertAfterSelf(text);
            oldText.Remove();
        }
        else if (ser.GetFirstChild<C.Order>() is { } order) order.InsertAfterSelf(text);
        else ser.PrependChild(text);

        var categoryCache = new C.StringCache(new C.PointCount { Val = (uint)categories.Count });
        for (var i = 0; i < categories.Count; i++)
            categoryCache.Append(new C.StringPoint(new C.NumericValue(categories[i])) { Index = (uint)i });
        var categoryData = new C.CategoryAxisData(new C.StringReference(new C.Formula($"ChartData!$A$2:$A${last}"), categoryCache));

        var oldValues = ser.GetFirstChild<C.Values>();
        var formatCode = oldValues?.Descendants<C.FormatCode>().FirstOrDefault()?.Text is { Length: > 0 } code ? code : "General";
        var valueCache = new C.NumberingCache(new C.FormatCode(formatCode), new C.PointCount { Val = (uint)data.Values.Count });
        for (var i = 0; i < data.Values.Count; i++)
            if (data.Values[i] is { } value)
                valueCache.Append(new C.NumericPoint(new C.NumericValue(value.ToString("R", CultureInfo.InvariantCulture))) { Index = (uint)i });
        var values = new C.Values(new C.NumberReference(new C.Formula($"ChartData!${column}$2:${column}${last}"), valueCache));

        if (oldValues is not null) oldValues.InsertAfterSelf(values);
        else ser.AppendChild(values);
        oldValues?.Remove();

        // Categories sit immediately before values in every series type this writes.
        ser.GetFirstChild<C.CategoryAxisData>()?.Remove();
        values.InsertBeforeSelf(categoryData);
    }

    /// <summary>
    /// Replaces the chart's data source with an embedded workbook holding exactly the new
    /// values, so "Edit Data" in PowerPoint opens what the chart shows rather than last month's
    /// numbers or a link to a file that is not there.
    /// </summary>
    private static void ReplaceWorkbook(ChartPart part, IReadOnlyList<string> categories, IReadOnlyList<ChartSeries> series)
    {
        var space = part.ChartSpace!;
        var external = space.GetFirstChild<C.ExternalData>();
        if (external?.Id?.Value is { Length: > 0 } oldId &&
            part.ExternalRelationships.FirstOrDefault(link => link.Id == oldId) is { } link)
            part.DeleteExternalRelationship(link);
        foreach (var embedded in part.Parts.Select(pair => pair.OpenXmlPart).OfType<EmbeddedPackagePart>().ToList())
            part.DeletePart(embedded);

        var workbook = part.AddEmbeddedPackagePart(WorkbookContentType);
        using (var output = workbook.GetStream(FileMode.Create, FileAccess.Write))
        {
            var bytes = SpreadsheetPartUtility.CreateChartWorkbook(categories, series);
            output.Write(bytes, 0, bytes.Length);
        }

        var replacement = new C.ExternalData(new C.AutoUpdate { Val = false }) { Id = part.GetIdOfPart(workbook) };
        if (external is not null)
        {
            external.InsertAfterSelf(replacement);
            external.Remove();
            return;
        }

        // CT_ChartSpace puts externalData after the chart, its shape and text properties,
        // and before print settings, user shapes and extensions.
        var before = space.ChildElements.FirstOrDefault(child =>
            child is C.PrintSettings or C.UserShapesReference or C.ExtensionList);
        if (before is not null) before.InsertBeforeSelf(replacement);
        else space.AppendChild(replacement);
    }

    private static uint Number(UInt32Value? value) => value?.Value ?? 0U;
}
