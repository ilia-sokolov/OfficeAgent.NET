using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeAgent.PowerPoint;

/// <summary>
/// Replaces the image of a picture on a slide in place, so a template's photo or logo slot
/// takes a new image and keeps its frame, position, effects, z-order, hyperlink and alt text.
/// </summary>
/// <remarks>
/// <para>
/// The target is what inspection lists for the picture: its <c>image#{slideId}/{shapeId}</c>
/// node, or the same picture's <c>shape#…</c> node. A picture placeholder that already holds a
/// picture is a picture like any other; one that is still an empty prompt holds nothing to
/// replace and is refused.
/// </para>
/// <para>
/// A placeholder picture often has no size of its own and takes its frame from the layout, so
/// the fit is computed against the inherited frame; the picture is left inheriting it.
/// The old image part is released only when nothing else on the slide still references it.
/// A deck is always <see cref="ChangeMode.Direct"/>: the module refuses a tracked plan.
/// </para>
/// </remarks>
internal sealed class SlideReplaceImageHandler : IOperationHandler
{
    /// <inheritdoc />
    public bool CanHandle(PlanOperation operation) =>
        operation is ReplaceImageOp { Target: NodeAnchor { Kind: "image" or "shape" } };

    /// <inheritdoc />
    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (ReplaceImageOp)operation;
        var anchor = (NodeAnchor)op.Target;

        if (PictureFill.Read(op, out var image) is { } unusable)
            return OperationPreview.Fail(new ValidationError(ValidationErrorCodes.InvalidOperation, unusable, anchor));

        var (located, refusal) = Locate(context, anchor);
        if (refusal is not null) return OperationPreview.Fail(refusal);

        var (width, height) = Frame(located!.Slide, located.Picture);
        if (width <= 0 || height <= 0)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"The picture at '{anchor.Path}' has no size, on the slide or in its layout, so a new image cannot be fitted to it.",
                anchor));

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "replaceImage",
            Before = PowerPointModel.ShapeNameOf(located.Picture),
            After = $"[{op.ImageType} {image!.Width}x{image.Height}px, {op.Fit.ToString().ToLowerInvariant()} in a " +
                    $"{width / 9525}x{height / 9525}px frame{(op.Description is null ? "" : $", alt='{op.Description}'")}]",
            Context = $"slide {located.Slide.Number}",
            BlastRadius = 1
        });
    }

    /// <inheritdoc />
    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (ReplaceImageOp)operation;
        var anchor = (NodeAnchor)op.Target;
        if (PictureFill.Read(op, out var image) is { } unusable)
            throw new InvalidOperationException(unusable);
        var (located, refusal) = Locate(context, anchor);
        if (refusal is not null)
            throw new InvalidOperationException($"Image '{anchor.Path}' vanished before apply.");

        var slide = located!.Slide;
        var picture = located.Picture;
        var (width, height) = Frame(slide, picture);
        var relationshipId = PictureFill.ReferenceOrAddImagePart(slide.Part, image!.PartType, image.Bytes);
        var previous = PictureFill.Refill(
            picture.BlipFill!, relationshipId, PictureFill.Crop(width, height, image.Width, image.Height, op.Fit));

        if (op.Description is not null && picture.NonVisualPictureProperties?.NonVisualDrawingProperties is { } properties)
            properties.Description = op.Description;

        PictureFill.ReleaseUnreferenced(slide.Part, previous);
    }

    /// <summary>The picture the anchor names, or why it names none that can be replaced.</summary>
    private static (PictureRef? Picture, ValidationError? Refusal) Locate(ApplyContext context, NodeAnchor anchor)
    {
        PictureRef? located = null;
        if (anchor.Kind == "image")
        {
            located = SlideImageNodeProvider.Locate(anchor.Path, new PowerPointObjectMap(context.Package));
        }
        else if (ShapeNodeProvider.Locate(anchor.Path, context.Package) is { } shape)
        {
            if (shape.Element is not P.Picture shapePicture)
                return (null, new ValidationError(
                    ValidationErrorCodes.InvalidOperation,
                    IsEmptyPicturePlaceholder(shape.Element)
                        ? $"'{anchor.Path}' is an empty picture placeholder: it holds no picture to replace."
                        : $"'{anchor.Path}' is not a picture, so it has no image to replace. Image paths come from inspect_document.nodes.",
                    anchor));
            located = new PictureRef(anchor.Path, shapePicture, shape.Slide);
        }

        if (located is null)
            return (null, new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No image with path '{anchor.Path}'. Image paths come from inspect_document.nodes.", anchor));

        if (located.Picture.BlipFill?.Blip is null)
            return (null, new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"The picture at '{anchor.Path}' has no image to replace.", anchor));

        // A video or audio object is a p:pic too, whose image is only its poster frame.
        var media = located.Picture.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties;
        if (media is not null && (media.GetFirstChild<A.VideoFromFile>() is not null ||
                                  media.GetFirstChild<A.AudioFromFile>() is not null ||
                                  media.Descendants().Any(element => element.LocalName == "media")))
            return (null, new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"'{anchor.Path}' is a video or audio object, not a picture; replaceImage does not change its poster frame.",
                anchor));

        return (located, null);
    }

    private static bool IsEmptyPicturePlaceholder(OpenXmlElement element) =>
        element is P.Shape shape &&
        shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value
            is { } type && type == P.PlaceholderValues.Picture;

    /// <summary>
    /// The frame the image fills, in EMUs: the picture's own extents, or, for a placeholder
    /// that inherits its size, the matching placeholder's on the layout and then the master.
    /// </summary>
    private static (long Width, long Height) Frame(SlideRef slide, P.Picture picture)
    {
        var own = picture.ShapeProperties?.Transform2D?.Extents;
        if (own?.Cx?.Value is > 0 && own.Cy?.Value is > 0)
            return (own.Cx.Value, own.Cy.Value);

        var placeholder = picture.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape;
        if (placeholder is null) return (0, 0);

        // PowerPoint's own order: the layout placeholder with the same index, then the layout's
        // and the master's of the same type.
        var layout = slide.Part.SlideLayoutPart;
        var layoutTree = layout?.SlideLayout?.CommonSlideData?.ShapeTree;
        var masterTree = layout?.SlideMasterPart?.SlideMaster?.CommonSlideData?.ShapeTree;
        var type = placeholder.Type?.Value ?? P.PlaceholderValues.Object;
        bool SameType(P.PlaceholderShape candidate) => (candidate.Type?.Value ?? P.PlaceholderValues.Object) == type;

        var inherited = (placeholder.Index?.Value is { } index
                            ? Extents(layoutTree, candidate => candidate.Index?.Value == index)
                            : null)
                        ?? Extents(layoutTree, SameType)
                        ?? Extents(masterTree, SameType);
        return inherited ?? (0, 0);
    }

    /// <summary>The extents of the first placeholder in <paramref name="tree"/> that <paramref name="matches"/> accepts.</summary>
    private static (long, long)? Extents(P.ShapeTree? tree, Func<P.PlaceholderShape, bool> matches)
    {
        if (tree is null) return null;
        foreach (var element in tree.Descendants())
        {
            var (properties, placeholder) = element switch
            {
                P.Shape s => (s.ShapeProperties, s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape),
                P.Picture p => (p.ShapeProperties, p.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape),
                _ => (null, null)
            };
            if (placeholder is null || !matches(placeholder)) continue;
            var extents = properties?.Transform2D?.Extents;
            if (extents?.Cx?.Value is > 0 && extents.Cy?.Value is > 0)
                return (extents.Cx.Value, extents.Cy.Value);
        }
        return null;
    }
}
