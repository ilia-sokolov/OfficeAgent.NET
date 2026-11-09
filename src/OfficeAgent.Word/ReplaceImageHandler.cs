using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using DocumentFormat.OpenXml.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace OfficeAgent.Word;

/// <summary>
/// Replaces the image of a picture addressed as <c>image#N</c>, inline or floating, keeping
/// everything the template set around it: the drawing's position and wrapping, its frame size,
/// effects, z-order, hyperlink and alt text.
/// </summary>
/// <remarks>
/// <para>
/// <see cref="ChangeMode.Direct"/> rewrites the picture's fill in place and releases the old
/// image part once nothing else in the part references it.
/// </para>
/// <para>
/// WordprocessingML has no revision for a picture's image, so a tracked replacement is written
/// the way Word itself records one: the run holding the old drawing is marked deleted and an
/// identical run holding the new image is marked inserted beside it. Accepting leaves the new
/// picture where the old one was; rejecting restores the old one, whose bytes are therefore
/// kept, as a tracked <c>removeImage</c> keeps them.
/// </para>
/// </remarks>
internal sealed class ReplaceImageHandler : IOperationHandler
{
    private const string Word2010DrawingNamespace = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing";

    private readonly ImageNodeProvider _images = new();

    public bool CanHandle(PlanOperation operation) =>
        operation is ReplaceImageOp { Target: NodeAnchor { Kind: "image" } };

    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (ReplaceImageOp)operation;
        var anchor = (NodeAnchor)op.Target;

        if (PictureFill.Read(op, out var image) is { } unusable)
            return OperationPreview.Fail(new ValidationError(ValidationErrorCodes.InvalidOperation, unusable, anchor));

        var located = _images.ResolveWithHost(anchor, new WordObjectMap(context.Package));
        if (located is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No image at path '{anchor.Path}'. Image paths come from inspect_document.nodes.", anchor));

        var drawing = located.Value.Drawing;
        if (Refusal(drawing, op, anchor) is { } refusal)
            return OperationPreview.Fail(refusal);

        var (width, height) = Frame(drawing);
        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "replaceImage",
            Before = $"[{NameOf(drawing)}]",
            After = $"[{op.ImageType} {image!.Width}x{image.Height}px, {op.Fit.ToString().ToLowerInvariant()} in a " +
                    $"{width / 9525}x{height / 9525}px frame{(op.Description is null ? "" : $", alt='{op.Description}'")}]",
            Context = anchor.Path,
            BlastRadius = 1
        });
    }

    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (ReplaceImageOp)operation;
        var anchor = (NodeAnchor)op.Target;
        if (PictureFill.Read(op, out var image) is { } unusable)
            throw new InvalidOperationException(unusable);
        var (drawing, host) = _images.ResolveWithHost(anchor, new WordObjectMap(context.Package))
            ?? throw new InvalidOperationException($"Image '{anchor.Path}' vanished before apply.");

        if (!WordRevisionMarker.IsTracked(op.Mode))
        {
            var previous = Refill(drawing, host, image!, op);
            PictureFill.ReleaseUnreferenced(host, previous);
            return;
        }

        // The old drawing's run is isolated first, so marking it deleted strikes only the
        // picture and not text that shared its run.
        var run = Isolate((Run)drawing.Parent!, drawing);
        var replacement = (Run)run.CloneNode(true);
        run.InsertAfterSelf(replacement);

        var copy = replacement.GetFirstChild<Drawing>()!;
        if (copy.Descendants<DW.DocProperties>().FirstOrDefault() is { } properties)
            properties.Id = InsertImageHandler.NextDrawingId(context.Package);
        // Word's own drawing ids must stay unique; it assigns new ones when they are absent.
        foreach (var frame in new OpenXmlElement?[] { copy.Inline, copy.Anchor })
        {
            frame?.RemoveAttribute("anchorId", Word2010DrawingNamespace);
            frame?.RemoveAttribute("editId", Word2010DrawingNamespace);
        }
        Refill(copy, host, image!, op);

        var marker = new WordRevisionMarker(context.Package, context.Revision);
        marker.MarkRunDeleted(run);
        marker.WrapInserted(replacement);
    }

    /// <summary>Why this drawing's image cannot be replaced, or null when it can.</summary>
    private static ValidationError? Refusal(Drawing drawing, ReplaceImageOp op, NodeAnchor anchor)
    {
        var picture = drawing.Descendants<PIC.Picture>().FirstOrDefault();
        if (picture?.BlipFill?.Blip is null)
            return new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"'{anchor.Path}' is a chart, diagram or shape, not a picture, so it has no image to replace.", anchor);

        var (width, height) = Frame(drawing);
        if (width <= 0 || height <= 0)
            return new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"The picture at '{anchor.Path}' has no size, so a new image cannot be fitted to it.", anchor);

        if (drawing.Ancestors<DeletedRun>().Any())
            return new ValidationError(
                ValidationErrorCodes.RevisionOverlap,
                $"The picture at '{anchor.Path}' is a pending tracked deletion. Accept or reject that revision first.", anchor);

        if (WordRevisionMarker.IsTracked(op.Mode))
        {
            if (drawing.Parent is not Run)
                return new ValidationError(
                    ValidationErrorCodes.InvalidOperation,
                    $"The picture at '{anchor.Path}' is not inside a run, so its replacement cannot be tracked. Use mode 'Direct'.", anchor);
            if (drawing.Ancestors<InsertedRun>().Any())
                return new ValidationError(
                    ValidationErrorCodes.RevisionOverlap,
                    $"The picture at '{anchor.Path}' is a pending tracked insertion. Accept or reject that revision first, " +
                    "or replace it with mode 'Direct'.", anchor);
        }

        return null;
    }

    /// <summary>
    /// Puts the new image in the drawing's picture, cropped to its frame, and returns the
    /// relationship ids the old image used.
    /// </summary>
    private static IReadOnlyList<string> Refill(Drawing drawing, OpenXmlPart host, PictureFill.Image image, ReplaceImageOp op)
    {
        var picture = drawing.Descendants<PIC.Picture>().First();
        var relationshipId = PictureFill.ReferenceOrAddImagePart(host, image.PartType, image.Bytes);
        var (width, height) = Frame(drawing);
        var previous = PictureFill.Refill(
            picture.BlipFill!, relationshipId, PictureFill.Crop(width, height, image.Width, image.Height, op.Fit));

        if (op.Description is not null)
        {
            if (drawing.Descendants<DW.DocProperties>().FirstOrDefault() is { } properties)
                properties.Description = op.Description;
            if (picture.NonVisualPictureProperties?.NonVisualDrawingProperties is { } nonVisual)
                nonVisual.Description = op.Description;
        }
        return previous;
    }

    /// <summary>
    /// The frame the image fills, in EMUs: the picture's own extents, or the drawing's when the
    /// picture leaves them out.
    /// </summary>
    private static (long Width, long Height) Frame(Drawing drawing)
    {
        var extents = drawing.Descendants<PIC.Picture>().FirstOrDefault()?.ShapeProperties?.Transform2D?.Extents;
        if (extents?.Cx?.Value is > 0 && extents.Cy?.Value is > 0)
            return (extents.Cx.Value, extents.Cy.Value);

        var extent = drawing.Inline?.Extent ?? drawing.Anchor?.GetFirstChild<DW.Extent>();
        return (extent?.Cx?.Value ?? 0, extent?.Cy?.Value ?? 0);
    }

    /// <summary>
    /// Leaves <paramref name="drawing"/> alone in its run, moving anything before or after it
    /// into runs of their own with the same formatting.
    /// </summary>
    private static Run Isolate(Run run, Drawing drawing)
    {
        var before = drawing.ElementsBefore().Where(element => element is not RunProperties).ToList();
        var after = drawing.ElementsAfter().ToList();
        if (before.Count > 0) run.InsertBeforeSelf(Split(run, before));
        if (after.Count > 0) run.InsertAfterSelf(Split(run, after));
        return run;
    }

    private static Run Split(Run run, IReadOnlyList<OpenXmlElement> content)
    {
        var split = new Run();
        if (run.RunProperties is { } properties) split.Append(properties.CloneNode(true));
        foreach (var element in content)
        {
            element.Remove();
            split.Append(element);
        }
        return split;
    }

    private static string NameOf(Drawing drawing) =>
        drawing.Descendants<DW.DocProperties>().FirstOrDefault()?.Name?.Value ?? "picture";
}
