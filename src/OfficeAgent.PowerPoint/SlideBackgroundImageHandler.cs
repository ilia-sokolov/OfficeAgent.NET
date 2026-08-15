using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;

namespace OfficeAgent.PowerPoint;

/// <summary>
/// Puts an image behind a slide, or behind every slide.
/// </summary>
/// <remarks>
/// A background is not a picture at the back of the shape tree: it lives in <c>p:bg</c>, so
/// it cannot be selected, nudged or deleted by accident, and a shape added afterwards is
/// always in front of it. That is also why it is a verb of its own rather than an
/// <c>insertImage</c> with a low z-order.
/// </remarks>
internal sealed class SlideBackgroundImageHandler : IOperationHandler
{
    /// <inheritdoc />
    public bool CanHandle(PlanOperation operation) =>
        operation is BackgroundImageOp { Target: null or NodeAnchor { Kind: "slide" } };

    /// <inheritdoc />
    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (BackgroundImageOp)operation;
        var anchor = op.Target as NodeAnchor;

        if (!SlidePaint.IsOpacity(op.Opacity))
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"opacity must be between 0 and 1; got {op.Opacity}.", op.Target));

        if (op.Scope is { Length: > 0 })
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "scope is a Word page setting. A deck names the slides instead: target one, " +
                "or omit the target to paint every slide.", op.Target));

        if (SlideImages.Unsupported(op) is { } unsupported)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation, unsupported, op.Target));

        var slides = Targets(context, anchor);
        if (slides.Count == 0)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No slide with path '{anchor!.Path}'. Slide paths come from inspect_document.nodes.",
                anchor));

        var clearing = string.IsNullOrEmpty(op.Base64Bytes);
        var where = anchor is null ? $"{slides.Count} slide(s)" : $"slide {slides[0].Number}";

        return OperationPreview.Ok(new ProposedChange
        {
            Target = op.Target ?? new NodeAnchor { Kind = "slide", Path = "slide#*" },
            Verb = "backgroundImage",
            Before = string.Empty,
            After = clearing ? "[background cleared]" : $"[background image{Strength(op)}]",
            Context = where,
            BlastRadius = slides.Count
        });
    }

    /// <inheritdoc />
    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (BackgroundImageOp)operation;
        var slides = Targets(context, op.Target as NodeAnchor);

        if (string.IsNullOrEmpty(op.Base64Bytes))
        {
            foreach (var slide in slides) SlidePaint.ClearBackground(slide);
            return;
        }

        if (!SlideImages.TryDecode(op.Base64Bytes!, out var bytes))
            throw new InvalidOperationException("backgroundImage base64Bytes is not valid base64.");

        var partType = SlideImages.PartTypeFor(op.ImageType)
            ?? throw new InvalidOperationException($"Unsupported imageType '{op.ImageType}'.");

        // Each slide owns its relationship, so the bytes are added per slide part rather
        // than once for the deck.
        foreach (var slide in slides)
        {
            var part = slide.Part.AddImagePart(partType);
            using (var stream = new MemoryStream(bytes))
                part.FeedData(stream);

            SlidePaint.SetBackgroundImage(slide, slide.Part.GetIdOfPart(part), op.Opacity);
        }
    }

    /// <summary>
    /// The slides the operation covers: the one named, or every slide when none is - which
    /// is what "apply to all" means for a background.
    /// </summary>
    private static List<SlideRef> Targets(ApplyContext context, NodeAnchor? anchor)
    {
        if (anchor is null) return PowerPointModel.Slides(context.Package).ToList();

        var slide = SlideList.Target(context, anchor);
        return slide is null ? new List<SlideRef>() : new List<SlideRef> { slide };
    }

    private static string Strength(BackgroundImageOp op) =>
        op.Opacity is { } o && o < 1 ? $" at {o:P0}" : string.Empty;
}
