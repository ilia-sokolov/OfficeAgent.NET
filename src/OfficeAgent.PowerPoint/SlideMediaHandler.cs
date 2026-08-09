using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using A = DocumentFormat.OpenXml.Drawing;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace OfficeAgent.PowerPoint;

/// <summary>
/// Embeds video or audio in a slide.
/// </summary>
/// <remarks>
/// An embedded clip is a <c>p:pic</c> wearing three relationships at once: the media part
/// itself, a video or audio reference to the same part, and a poster image for the frame.
/// PowerPoint needs all three - the <c>p14:media</c> extension is what makes the clip
/// <em>embedded</em> rather than linked, and without the poster the picture has no fill to
/// draw. The bytes travel inside the package, so the deck still plays when it is mailed on.
/// </remarks>
internal sealed class SlideInsertMediaHandler : IOperationHandler
{
    /// <summary>The uri PowerPoint stores the embedded-media reference under.</summary>
    private const string MediaExtensionUri = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}";

    /// <summary>
    /// A 1x1 transparent PNG. A picture must have a fill to render at all, so a clip with
    /// no poster still needs one; PowerPoint does not generate a frame from the media.
    /// </summary>
    private const string BlankPoster =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";

    private static readonly Dictionary<string, (string Media, MediaKind Kind)> Types =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["mp4"] = ("video/mp4", MediaKind.Video),
            ["m4v"] = ("video/mp4", MediaKind.Video),
            ["mov"] = ("video/quicktime", MediaKind.Video),
            ["wmv"] = ("video/x-ms-wmv", MediaKind.Video),
            ["avi"] = ("video/avi", MediaKind.Video),
            ["mp3"] = ("audio/mpeg", MediaKind.Audio),
            ["m4a"] = ("audio/mp4", MediaKind.Audio),
            ["wav"] = ("audio/wav", MediaKind.Audio),
            ["wma"] = ("audio/x-ms-wma", MediaKind.Audio)
        };

    /// <inheritdoc />
    public bool CanHandle(PlanOperation operation) =>
        operation is InsertMediaOp { Target: NodeAnchor { Kind: "slide" } };

    /// <inheritdoc />
    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (InsertMediaOp)operation;
        var anchor = (NodeAnchor)op.Target;

        var slide = SlideList.Target(context, anchor);
        if (slide is null) return OperationPreview.Fail(SlideList.NoSuchSlide(anchor));

        if (string.IsNullOrEmpty(op.Base64Bytes) == string.IsNullOrEmpty(op.MediaDocumentId))
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "insertMedia needs exactly one of base64Bytes or mediaDocumentId.", anchor));

        if (!Types.TryGetValue(op.MediaType.TrimStart('.'), out var descriptor))
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"Unknown mediaType '{op.MediaType}'. Expected one of: {string.Join(", ", Types.Keys)}.",
                anchor));

        // Saying "audio" over an .mp4 would produce a deck that plays nothing, because the
        // element PowerPoint reads is chosen by the declared kind, not by the bytes.
        if (descriptor.Kind != op.Kind)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"mediaType '{op.MediaType}' is {descriptor.Kind.ToString().ToLowerInvariant()}, " +
                $"but kind says {op.Kind.ToString().ToLowerInvariant()}.",
                anchor));

        if (op.WidthPx <= 0 || op.HeightPx <= 0)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "widthPx and heightPx must be positive.", anchor));

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "insertMedia",
            Before = string.Empty,
            After = $"[{op.Kind.ToString().ToLowerInvariant()} {op.MediaType} {op.WidthPx}×{op.HeightPx}]",
            Context = $"slide {slide.Number}",
            BlastRadius = 1
        });
    }

    /// <inheritdoc />
    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (InsertMediaOp)operation;
        var anchor = (NodeAnchor)op.Target;

        var slide = SlideList.Target(context, anchor)
            ?? throw new InvalidOperationException($"Slide '{anchor.Path}' vanished before apply.");
        var tree = slide.Part.Slide.CommonSlideData?.ShapeTree
            ?? throw new InvalidOperationException($"Slide '{anchor.Path}' has no shape tree.");

        var descriptor = Types[op.MediaType.TrimStart('.')];
        var bytes = Convert.FromBase64String(op.Base64Bytes!);

        var mediaPart = PowerPointModel.Doc(context.Package).CreateMediaDataPart(descriptor.Media);
        using (var stream = mediaPart.GetStream())
            stream.Write(bytes, 0, bytes.Length);

        // The same part, referenced twice: once as the linked media PowerPoint plays, once
        // through the p14 extension that marks it embedded. Both are required - a clip with
        // only the reference is treated as a link to a file that is not there.
        var referenceId = op.Kind == MediaKind.Video
            ? slide.Part.AddVideoReferenceRelationship(mediaPart).Id
            : slide.Part.AddAudioReferenceRelationship(mediaPart).Id;
        var embedId = slide.Part.AddMediaReferenceRelationship(mediaPart).Id;

        var poster = slide.Part.AddImagePart(ImagePartType.Png);
        var posterBytes = Convert.FromBase64String(op.PosterBase64 ?? BlankPoster);
        using (var stream = poster.GetStream())
            stream.Write(posterBytes, 0, posterBytes.Length);
        var posterId = slide.Part.GetIdOfPart(poster);

        var nonVisual = new ApplicationNonVisualDrawingProperties();
        nonVisual.Append(op.Kind == MediaKind.Video
            ? new A.VideoFromFile { Link = referenceId }
            : (OpenXmlElement)new A.AudioFromFile { Link = referenceId });
        nonVisual.Append(new ApplicationNonVisualDrawingPropertiesExtensionList(
            new ApplicationNonVisualDrawingPropertiesExtension(
                new P14.Media { Embed = embedId })
            { Uri = MediaExtensionUri }));

        var drawing = new NonVisualDrawingProperties
        {
            Id = PowerPointModel.NextShapeId(slide.Part),
            Name = $"{op.Kind} {op.MediaType}"
        };
        if (op.AltText is { Length: > 0 } alt) drawing.Description = alt;
        // The click action is what gives the shape PowerPoint's play controls.
        drawing.Append(new A.HyperlinkOnClick { Id = string.Empty, Action = "ppaction://media" });

        tree.Append(new Picture(
            new NonVisualPictureProperties(
                drawing,
                new NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true }),
                nonVisual),
            new BlipFill(
                new A.Blip { Embed = posterId },
                new A.Stretch(new A.FillRectangle())),
            new ShapeProperties(
                new A.Transform2D(
                    new A.Offset
                    {
                        X = op.XPx is { } x ? Emu.FromPixels(x) : 838200L,
                        Y = op.YPx is { } y ? Emu.FromPixels(y) : 1825625L
                    },
                    new A.Extents
                    {
                        Cx = Emu.FromPixels(op.WidthPx),
                        Cy = Emu.FromPixels(op.HeightPx)
                    }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })));
    }
}

/// <summary>Surfaces embedded clips as addressable nodes so a plan can remove one.</summary>
internal sealed class SlideMediaNodeProvider : IPowerPointNodeProvider
{
    /// <inheritdoc />
    public string Kind => "media";

    /// <inheritdoc />
    public IEnumerable<NodeInfo> Enumerate(PowerPointObjectMap map)
    {
        foreach (var slide in PowerPointModel.Slides(map.Package))
        {
            var tree = slide.Part.Slide.CommonSlideData?.ShapeTree;
            if (tree is null) continue;

            foreach (var picture in tree.Elements<Picture>())
            {
                var nonVisual = picture.NonVisualPictureProperties?
                    .ApplicationNonVisualDrawingProperties;
                var kind = nonVisual?.GetFirstChild<A.VideoFromFile>() is not null ? "video"
                    : nonVisual?.GetFirstChild<A.AudioFromFile>() is not null ? "audio"
                    : null;
                if (kind is null) continue;

                var shapeId = PowerPointModel.ShapeIdOf(picture);
                var path = $"media#{slide.SlideId}/{shapeId}";
                yield return new NodeInfo
                {
                    Kind = Kind,
                    Path = path,
                    Summary = $"slide {slide.Number}: embedded {kind} '{PowerPointModel.ShapeNameOf(picture)}'",
                    Anchor = new NodeAnchor { Id = path, Kind = Kind, Path = path }
                };
            }
        }
    }

    /// <inheritdoc />
    public ResolvedNode? Resolve(NodeAnchor anchor, PowerPointObjectMap map)
    {
        var located = Enumerate(map).FirstOrDefault(n =>
            string.Equals(n.Path, anchor.Path, StringComparison.OrdinalIgnoreCase));
        if (located is null) return null;

        // Media shares the shape addressing space, so removal goes through removeShape.
        var shapePath = "shape#" + anchor.Path.Substring("media#".Length);
        var shape = ShapeNodeProvider.Locate(shapePath, map.Package);
        return shape is null
            ? null
            : new ResolvedNode { Kind = Kind, Elements = new[] { shape.Element }, Value = located.Summary };
    }
}
