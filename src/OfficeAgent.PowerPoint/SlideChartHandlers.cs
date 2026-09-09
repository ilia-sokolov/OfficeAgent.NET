using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeAgent.PowerPoint;

internal sealed class SlideInsertChartHandler : IOperationHandler
{
    public bool CanHandle(PlanOperation operation) => operation is InsertChartOp { Target: NodeAnchor { Kind: "slide" } };

    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (InsertChartOp)operation;
        var anchor = (NodeAnchor)op.Target;
        if (!SlideNodeProvider.TryParseSlideId(anchor.Path, out var slideId) ||
            PowerPointModel.Slide(context.Package, slideId) is not { } slide)
            return Fail($"No slide with path '{anchor.Path}'.", anchor, ValidationErrorCodes.AnchorNotFound);
        var error = Validate(op.Kind, op.Categories, op.Series, op.WidthPx, op.HeightPx);
        if (error is not null) return Fail(error, anchor);
        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "insertChart",
            After = $"[{op.Kind} chart: {op.Categories.Count} categories, {op.Series.Count} series]",
            Context = $"slide {slide.Number}"
        });
    }

    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (InsertChartOp)operation;
        var anchor = (NodeAnchor)op.Target;
        SlideNodeProvider.TryParseSlideId(anchor.Path, out var slideId);
        var slide = PowerPointModel.Slide(context.Package, slideId)
            ?? throw new InvalidOperationException($"Slide '{anchor.Path}' vanished before apply.");
        var id = PowerPointModel.NextShapeId(slide.Part);
        var frame = SlideChartBuilder.Add(slide.Part, id, op);
        slide.Part.Slide.CommonSlideData!.ShapeTree!.Append(frame);
        slide.Part.Slide.Save();
    }

    internal static string? Validate(ChartKind kind, IReadOnlyList<string> categories, IReadOnlyList<ChartSeries> series, int width = 1, int height = 1)
    {
        if (categories.Count == 0) return "A chart needs at least one category.";
        if (series.Count == 0) return "A chart needs at least one numeric series.";
        if (series.Any(item => item.Values.Count != categories.Count))
            return "Every chart series must contain one value per category.";
        if (series.SelectMany(item => item.Values).Any(value => value is { } number &&
                (double.IsNaN(number) || double.IsInfinity(number))))
            return "Chart values must be finite numbers.";
        if (kind == ChartKind.Pie && series.Count != 1)
            return "A pie chart needs exactly one series.";
        if (width <= 0 || height <= 0) return "Chart width and height must be positive.";
        return null;
    }

    private static OperationPreview Fail(string message, Anchor anchor, string code = ValidationErrorCodes.InvalidOperation) =>
        OperationPreview.Fail(new ValidationError(code, message, anchor));
}

internal sealed class SlideUpdateChartHandler : IOperationHandler
{
    public bool CanHandle(PlanOperation operation) => operation is UpdateChartOp { Target: NodeAnchor { Kind: "chart" } };

    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (UpdateChartOp)operation;
        var anchor = (NodeAnchor)op.Target;
        var located = SlideChartNodeProvider.Locate(anchor.Path, context.Package);
        if (located is null)
            return Fail($"No chart with path '{anchor.Path}'.", anchor, ValidationErrorCodes.AnchorNotFound);
        if (!PowerPointModel.ShapeNameOf(located.Frame).StartsWith(SlideChartBuilder.FrameNamePrefix, StringComparison.Ordinal))
            return Fail("Only charts created by OfficeAgent can be updated.", anchor);
        var error = SlideInsertChartHandler.Validate(op.Kind, op.Categories, op.Series);
        if (error is not null) return Fail(error, anchor);
        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "updateChart",
            Before = "[native chart]",
            After = $"[{op.Kind} chart: {op.Categories.Count} categories, {op.Series.Count} series]",
            Context = $"slide {located.Slide.Number}"
        });
    }

    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (UpdateChartOp)operation;
        var anchor = (NodeAnchor)op.Target;
        var located = SlideChartNodeProvider.Locate(anchor.Path, context.Package)
            ?? throw new InvalidOperationException($"Chart '{anchor.Path}' vanished before apply.");
        SlideChartBuilder.Write(located.Part, op.Kind, op.Categories, op.Series, op.Title, op.ShowLegend);
        located.Frame.NonVisualGraphicFrameProperties!.NonVisualDrawingProperties!.Description = op.Description;
        located.Slide.Part.Slide.Save();
    }

    private static OperationPreview Fail(string message, Anchor anchor, string code = ValidationErrorCodes.InvalidOperation) =>
        OperationPreview.Fail(new ValidationError(code, message, anchor));
}

internal sealed class SlideChartNodeProvider : IPowerPointNodeProvider
{
    public string Kind => "chart";

    public IEnumerable<NodeInfo> Enumerate(PowerPointObjectMap map)
    {
        foreach (var slide in PowerPointModel.Slides(map.Package))
            foreach (var frame in slide.Part.Slide.Descendants<P.GraphicFrame>())
            {
                var reference = frame.Descendants<C.ChartReference>().FirstOrDefault();
                var id = PowerPointModel.ShapeIdOf(frame);
                if (reference?.Id?.Value is not { Length: > 0 } || id is null) continue;
                var path = $"chart#{slide.SlideId}/{id}";
                yield return new NodeInfo
                {
                    Kind = Kind,
                    Path = path,
                    Summary = $"slide {slide.Number}: {PowerPointModel.ShapeNameOf(frame)}",
                    Anchor = new NodeAnchor { Id = path, Kind = Kind, Path = path }
                };
            }
    }

    public ResolvedNode? Resolve(NodeAnchor anchor, PowerPointObjectMap map)
    {
        var located = Locate(anchor.Path, map.Package);
        return located is null ? null : new ResolvedNode
        {
            Kind = Kind,
            Elements = new[] { (OpenXmlElement)located.Frame },
            Value = PowerPointModel.ShapeNameOf(located.Frame)
        };
    }

    public static LocatedChart? Locate(string path, IOpenXmlPackage package)
    {
        var value = path.StartsWith("chart#", StringComparison.OrdinalIgnoreCase) ? path.Substring(6) : path;
        var slash = value.IndexOf('/');
        if (slash <= 0 || !uint.TryParse(value.Substring(0, slash), out var slideId) ||
            !uint.TryParse(value.Substring(slash + 1), out var shapeId)) return null;
        var slide = PowerPointModel.Slide(package, slideId);
        var frame = slide?.Part.Slide.Descendants<P.GraphicFrame>()
            .FirstOrDefault(candidate => PowerPointModel.ShapeIdOf(candidate) == shapeId);
        var reference = frame?.Descendants<C.ChartReference>().FirstOrDefault();
        if (slide is null || frame is null || reference?.Id?.Value is not { Length: > 0 } relationshipId ||
            slide.Part.GetPartById(relationshipId) is not ChartPart chartPart) return null;
        return new LocatedChart(slide, frame, chartPart);
    }
}

internal sealed class LocatedChart
{
    public LocatedChart(SlideRef slide, P.GraphicFrame frame, ChartPart part)
    {
        Slide = slide;
        Frame = frame;
        Part = part;
    }
    public SlideRef Slide { get; }
    public P.GraphicFrame Frame { get; }
    public ChartPart Part { get; }
}
