using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeAgent.PowerPoint;

/// <summary>
/// Surfaces each picture on a slide as an addressable node, so an image can be removed
/// or reported without the agent inventing a path.
/// </summary>
internal sealed class SlideImageNodeProvider : IPowerPointNodeProvider
{
    /// <inheritdoc />
    public string Kind => "image";

    /// <inheritdoc />
    public IEnumerable<NodeInfo> Enumerate(PowerPointObjectMap map)
    {
        foreach (var picture in Pictures(map))
        {
            var name = PowerPointModel.ShapeNameOf(picture.Picture);
            var alt = picture.Picture.NonVisualPictureProperties?
                .NonVisualDrawingProperties?.Description?.Value;

            yield return new NodeInfo
            {
                Kind = Kind,
                Path = picture.Path,
                Summary = $"slide {picture.Slide.Number}: {(name.Length > 0 ? name : "picture")}" +
                          (string.IsNullOrEmpty(alt) ? string.Empty : $" - {alt}"),
                Anchor = new NodeAnchor { Id = picture.Path, Kind = Kind, Path = picture.Path }
            };
        }
    }

    /// <inheritdoc />
    public ResolvedNode? Resolve(NodeAnchor anchor, PowerPointObjectMap map)
    {
        var located = Locate(anchor.Path, map);
        return located is null
            ? null
            : new ResolvedNode
            {
                Kind = Kind,
                Elements = new OpenXmlElement[] { located.Picture },
                Value = located.Path
            };
    }

    /// <summary>Finds one picture by its path, or null when it is gone.</summary>
    internal static PictureRef? Locate(string path, PowerPointObjectMap map)
    {
        foreach (var picture in Pictures(map))
            if (string.Equals(picture.Path, path, StringComparison.OrdinalIgnoreCase))
                return picture;
        return null;
    }

    /// <summary>Enumerates every picture in the deck, in slide order.</summary>
    internal static IEnumerable<PictureRef> Pictures(PowerPointObjectMap map)
    {
        foreach (var slide in PowerPointModel.Slides(map.Package))
            foreach (var picture in slide.Part.Slide.Descendants<P.Picture>())
            {
                var shapeId = PowerPointModel.ShapeIdOf(picture);
                if (shapeId is null) continue;
                yield return new PictureRef($"image#{slide.SlideId}/{shapeId}", picture, slide);
            }
    }
}

/// <summary>One picture and the slide it sits on.</summary>
internal sealed class PictureRef
{
    public PictureRef(string path, P.Picture picture, SlideRef slide)
    {
        Path = path;
        Picture = picture;
        Slide = slide;
    }

    public string Path { get; }
    public P.Picture Picture { get; }
    public SlideRef Slide { get; }
}

/// <summary>
/// Places an image on a slide as a <c>p:pic</c> backed by a new image part.
/// </summary>
/// <remarks>
/// The target is the slide, addressed as <c>{ "kind": "slide", "path": "slide#256" }</c>,
/// for the same reason tables are: a slide has no text flow to anchor a picture to. The
/// bytes must arrive as base64 - a provider-backed <c>imageDocumentId</c> is resolved to
/// base64 by <see cref="OfficeAgentClient"/> before the plan reaches this handler.
/// </remarks>
internal sealed class SlideInsertImageHandler : IOperationHandler
{
    /// <summary>Pixels-to-EMU at 96 DPI, the unit DrawingML positions shapes in.</summary>
    private const long EmuPerPixel = 9525L;

    /// <inheritdoc />
    public bool CanHandle(PlanOperation operation) =>
        operation is InsertImageOp { Target: NodeAnchor { Kind: "slide" } };

    /// <inheritdoc />
    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (InsertImageOp)operation;
        var anchor = (NodeAnchor)op.Target!;

        var slide = ResolveSlide(context, anchor);
        if (slide is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No slide with path '{anchor.Path}'. Slide paths come from inspect_document.nodes.",
                anchor));

        if (string.IsNullOrEmpty(op.Base64Bytes))
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "insertImage needs either base64Bytes or an imageConnectionId/imageDocumentId pair.",
                anchor));

        if (!TryDecode(op.Base64Bytes!, out _))
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "insertImage base64Bytes is not valid base64.", anchor));

        if (PartTypeFor(op.ImageType) is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"Unsupported imageType '{op.ImageType}'. Use png, jpeg, gif, bmp, or tiff.", anchor));

        if (op.WidthPx <= 0 || op.HeightPx <= 0)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "insertImage needs a positive widthPx and heightPx.", anchor));

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "insertImage",
            Before = string.Empty,
            After = $"[image {op.WidthPx}×{op.HeightPx}]",
            Context = $"slide {slide.Number}",
            BlastRadius = 1
        });
    }

    /// <inheritdoc />
    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (InsertImageOp)operation;
        var anchor = (NodeAnchor)op.Target!;

        var slide = ResolveSlide(context, anchor)
            ?? throw new InvalidOperationException($"Slide '{anchor.Path}' vanished before apply.");
        if (!TryDecode(op.Base64Bytes!, out var bytes))
            throw new InvalidOperationException("insertImage base64Bytes is not valid base64.");

        var partType = PartTypeFor(op.ImageType)
            ?? throw new InvalidOperationException($"Unsupported imageType '{op.ImageType}'.");

        var imagePart = slide.Part.AddImagePart(partType);
        using (var stream = new MemoryStream(bytes))
            imagePart.FeedData(stream);
        var relationshipId = slide.Part.GetIdOfPart(imagePart);

        var shapeId = PowerPointModel.NextShapeId(slide.Part);
        var tree = slide.Part.Slide.CommonSlideData?.ShapeTree
            ?? throw new InvalidOperationException($"Slide '{anchor.Path}' has no shape tree.");

        tree.Append(new P.Picture(
            new P.NonVisualPictureProperties(
                new P.NonVisualDrawingProperties
                {
                    Id = shapeId,
                    Name = $"Picture {shapeId}",
                    // Alt text is the accessibility contract; carry it when supplied.
                    Description = string.IsNullOrEmpty(op.AltText) ? null : op.AltText
                },
                new P.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true }),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.BlipFill(
                new A.Blip { Embed = relationshipId },
                new A.Stretch(new A.FillRectangle())),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = 914400L, Y = LowestEdge(slide) },
                    new A.Extents
                    {
                        Cx = op.WidthPx * EmuPerPixel,
                        Cy = op.HeightPx * EmuPerPixel
                    }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })));
    }

    private static SlideRef? ResolveSlide(ApplyContext context, NodeAnchor anchor) =>
        SlideNodeProvider.TryParseSlideId(anchor.Path, out var slideId)
            ? PowerPointModel.Slide(context.Package, slideId)
            : null;

    private static bool TryDecode(string base64, out byte[] bytes)
    {
        try
        {
            bytes = Convert.FromBase64String(base64);
            return bytes.Length > 0;
        }
        catch (FormatException)
        {
            bytes = Array.Empty<byte>();
            return false;
        }
    }

    /// <summary>
    /// Maps the plan's image-type vocabulary onto Open XML part types, matching the set
    /// the Word module accepts so one plan vocabulary serves both formats.
    /// </summary>
    private static PartTypeInfo? PartTypeFor(string? imageType)
    {
        if (string.IsNullOrEmpty(imageType)) return ImagePartType.Png;
        return ImagePartTypes.TryGetValue(imageType!, out var type) ? type : null;
    }

    private static readonly Dictionary<string, PartTypeInfo> ImagePartTypes =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["png"] = ImagePartType.Png,
            ["jpeg"] = ImagePartType.Jpeg,
            ["jpg"] = ImagePartType.Jpeg,
            ["gif"] = ImagePartType.Gif,
            ["bmp"] = ImagePartType.Bmp,
            ["tiff"] = ImagePartType.Tiff,
            ["tif"] = ImagePartType.Tiff
        };

    /// <summary>Places the picture below existing content rather than on top of it.</summary>
    private static long LowestEdge(SlideRef slide)
    {
        long lowest = 0;
        foreach (var transform in slide.Part.Slide.Descendants<A.Transform2D>())
        {
            var y = transform.Offset?.Y?.Value ?? 0L;
            var height = transform.Extents?.Cy?.Value ?? 0L;
            if (y + height > lowest) lowest = y + height;
        }
        foreach (var transform in slide.Part.Slide.Descendants<P.Transform>())
        {
            var y = transform.Offset?.Y?.Value ?? 0L;
            var height = transform.Extents?.Cy?.Value ?? 0L;
            if (y + height > lowest) lowest = y + height;
        }
        return lowest > 0 ? lowest + 228600L : 914400L;
    }
}

/// <summary>
/// Removes a picture from its slide. The underlying image part is left in the package:
/// other slides may share it, and Open XML tolerates an unreferenced part.
/// </summary>
internal sealed class SlideRemoveImageHandler : IOperationHandler
{
    /// <inheritdoc />
    public bool CanHandle(PlanOperation operation) =>
        operation is RemoveImageOp { Target: NodeAnchor { Kind: "image" } };

    /// <inheritdoc />
    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var anchor = (NodeAnchor)operation.Target!;
        var picture = SlideImageNodeProvider.Locate(anchor.Path, new PowerPointObjectMap(context.Package));
        if (picture is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No image with path '{anchor.Path}'. Image paths come from inspect_document.nodes.",
                anchor));

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "removeImage",
            Before = PowerPointModel.ShapeNameOf(picture.Picture),
            After = string.Empty,
            Context = $"slide {picture.Slide.Number}",
            BlastRadius = 1
        });
    }

    /// <inheritdoc />
    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var anchor = (NodeAnchor)operation.Target!;
        var picture = SlideImageNodeProvider.Locate(anchor.Path, new PowerPointObjectMap(context.Package))
            ?? throw new InvalidOperationException($"Image '{anchor.Path}' vanished before apply.");

        picture.Picture.Remove();
    }
}
