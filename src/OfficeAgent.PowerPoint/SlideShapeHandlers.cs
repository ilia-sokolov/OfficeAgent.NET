using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeAgent.PowerPoint;

/// <summary>Pixels at 96 DPI to EMUs, the unit DrawingML positions everything in.</summary>
internal static class Emu
{
    public const long PerPixel = 9525L;

    public static long FromPixels(int pixels) => pixels * PerPixel;
}

/// <summary>
/// Inserts a paragraph - a bullet or a line - into an existing text body.
/// </summary>
/// <remarks>
/// This is the shared <c>insert</c> verb, targeting a paragraph the way it does in Word.
/// The new paragraph copies the anchor paragraph's properties, so a bullet inserted next
/// to a bullet is styled like it, and <c>level</c> overrides the depth when the caller
/// wants a sub-bullet. Inserting renumbers every later paragraph in the same body, which
/// is why <see cref="PowerPointModule.ValidatePlan"/> refuses a plan that would then
/// address that body positionally.
/// </remarks>
internal sealed class SlideInsertParagraphHandler : IOperationHandler
{
    /// <inheritdoc />
    public bool CanHandle(PlanOperation operation) =>
        operation is InsertOp { Target: TextSpanAnchor };

    /// <inheritdoc />
    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (InsertOp)operation;
        var anchor = (TextSpanAnchor)op.Target;

        var paragraph = PowerPointModel.ResolveParagraph(context, anchor.ParaId);
        if (paragraph is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No paragraph with id '{anchor.ParaId}'. Paragraph ids come from inspect_document.",
                anchor));

        if (op.StyleId is { Length: > 0 })
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "The PowerPoint module cannot apply styleId - a deck has no paragraph style " +
                "table. Use \"level\" for bullet depth, or the format verb for run properties.",
                anchor));

        if (op.Level is < 0 or > 8)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "level is the bullet depth and must be between 0 and 8.", anchor));

        var text = op.Text ?? string.Empty;
        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "insert",
            Before = string.Empty,
            After = op.Level is { } level ? $"{text} (level {level})" : text,
            Context = $"slide {paragraph.Slide.Number} ({paragraph.Location}), {op.Position} p{IndexOf(anchor.ParaId)}",
            BlastRadius = 1
        });
    }

    /// <inheritdoc />
    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (InsertOp)operation;
        var anchor = (TextSpanAnchor)op.Target;

        var paragraph = PowerPointModel.ResolveParagraph(context, anchor.ParaId)
            ?? throw new InvalidOperationException($"Paragraph '{anchor.ParaId}' vanished before apply.");

        var inserted = new A.Paragraph();

        // Carry the neighbour's paragraph properties so the new line is bulleted, indented
        // and aligned like the one it sits beside rather than reverting to the body default.
        if (paragraph.Paragraph.ParagraphProperties is { } source)
            inserted.ParagraphProperties = (A.ParagraphProperties)source.CloneNode(deep: true);
        if (op.Level is { } level)
            (inserted.ParagraphProperties ??= new A.ParagraphProperties()).Level = level;

        // Match the run properties of the text it joins, for the same reason.
        var template = paragraph.Paragraph.Elements<A.Run>().FirstOrDefault();
        var runProperties = template?.RunProperties is { } rp
            ? (A.RunProperties)rp.CloneNode(deep: true)
            : new A.RunProperties { Language = "en-US" };

        inserted.Append(new A.Run(runProperties, new A.Text(op.Text ?? string.Empty)));

        if (op.Position == InsertPosition.Before)
            paragraph.Paragraph.InsertBeforeSelf(inserted);
        else
            paragraph.Paragraph.InsertAfterSelf(inserted);
    }

    private static string IndexOf(string paraId)
    {
        var slash = paraId.LastIndexOf("/p", StringComparison.Ordinal);
        return slash < 0 ? paraId : paraId.Substring(slash + 2);
    }
}

/// <summary>Adds a free-standing text box to a slide.</summary>
internal sealed class SlideInsertShapeHandler : IOperationHandler
{
    private const long DefaultMargin = 838200L;

    /// <inheritdoc />
    public bool CanHandle(PlanOperation operation) =>
        operation is InsertShapeOp { Target: NodeAnchor { Kind: "slide" } };

    /// <inheritdoc />
    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (InsertShapeOp)operation;
        var anchor = (NodeAnchor)op.Target;

        var slide = SlideList.Target(context, anchor);
        if (slide is null) return OperationPreview.Fail(SlideList.NoSuchSlide(anchor));

        if (op.WidthPx <= 0 || op.HeightPx <= 0)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "widthPx and heightPx must be positive.", anchor));

        var text = string.Join(" / ", op.Text.Take(2));
        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "insertShape",
            Before = string.Empty,
            After = $"[text box {op.WidthPx}×{op.HeightPx}] {text}",
            Context = $"slide {slide.Number}",
            BlastRadius = 1
        });
    }

    /// <inheritdoc />
    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (InsertShapeOp)operation;
        var anchor = (NodeAnchor)op.Target;

        var slide = SlideList.Target(context, anchor)
            ?? throw new InvalidOperationException($"Slide '{anchor.Path}' vanished before apply.");

        var tree = slide.Part.Slide.CommonSlideData?.ShapeTree
            ?? throw new InvalidOperationException($"Slide '{anchor.Path}' has no shape tree.");

        var body = new TextBody(new A.BodyProperties(), new A.ListStyle());
        foreach (var line in op.Text.Count > 0 ? op.Text : new[] { string.Empty })
            body.Append(new A.Paragraph(
                new A.Run(new A.RunProperties { Language = "en-US" }, new A.Text(line))));

        var shape = new Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties
                {
                    Id = PowerPointModel.NextShapeId(slide.Part),
                    Name = "TextBox"
                },
                new NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties()),
            new ShapeProperties(
                new A.Transform2D(
                    new A.Offset
                    {
                        X = op.XPx is { } x ? Emu.FromPixels(x) : DefaultMargin,
                        Y = op.YPx is { } y ? Emu.FromPixels(y) : BelowEverything(slide)
                    },
                    new A.Extents
                    {
                        Cx = Emu.FromPixels(op.WidthPx),
                        Cy = Emu.FromPixels(op.HeightPx)
                    }),
                // A text box is a plain rectangle with no fill or outline, which is what
                // PowerPoint inserts; a preset geometry is required for it to render.
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle },
                new A.NoFill()),
            body);

        tree.Append(shape);
    }

    /// <summary>Below the lowest existing shape, so a new box does not land on the content.</summary>
    private static long BelowEverything(SlideRef slide)
    {
        long lowest = 0;
        foreach (var transform in slide.Part.Slide.Descendants<A.Transform2D>())
        {
            var bottom = (transform.Offset?.Y?.Value ?? 0L) + (transform.Extents?.Cy?.Value ?? 0L);
            if (bottom > lowest) lowest = bottom;
        }
        return lowest > 0 ? lowest + 91440L : DefaultMargin;
    }
}

/// <summary>Removes any shape a slide holds - a text box, a table frame, or a picture.</summary>
internal sealed class SlideRemoveShapeHandler : IOperationHandler
{
    /// <inheritdoc />
    public bool CanHandle(PlanOperation operation) =>
        operation is RemoveShapeOp { Target: NodeAnchor { Kind: "shape" } };

    /// <inheritdoc />
    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var anchor = (NodeAnchor)operation.Target;

        var located = ShapeNodeProvider.Locate(anchor.Path, context.Package);
        if (located is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No shape with path '{anchor.Path}'. Shape paths come from inspect_document.nodes.",
                anchor));

        if (ShapeNodeProvider.IsPlaceholder(located.Element))
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "That shape is a layout placeholder. Removing it leaves the layout re-offering " +
                "an empty prompt, so the slide looks unchanged while its content is gone. " +
                "Clear its text with changeText instead.",
                anchor));

        var text = string.Concat(located.Element.Descendants<A.Text>().Select(t => t.Text));
        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "removeShape",
            Before = text.Length > 60 ? text.Substring(0, 60) + "…" : text,
            After = string.Empty,
            Context = $"slide {located.Slide.Number}",
            BlastRadius = 1
        });
    }

    /// <inheritdoc />
    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var anchor = (NodeAnchor)operation.Target;
        var located = ShapeNodeProvider.Locate(anchor.Path, context.Package)
            ?? throw new InvalidOperationException($"Shape '{anchor.Path}' vanished before apply.");

        located.Element.Remove();
    }
}
