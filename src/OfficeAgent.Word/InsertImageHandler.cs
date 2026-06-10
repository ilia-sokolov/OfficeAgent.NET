using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace OfficeAgent.Word;

/// <summary>
/// Inserts an inline image anchored to a <see cref="TextSpanAnchor"/>. The image
/// bytes come from base64-encoded <see cref="InsertImageOp.Base64Bytes"/>. When the
/// caller supplied an <see cref="InsertImageOp.ImageDocumentId"/> instead,
/// <see cref="OfficeAgentClient"/> resolves it through the provider registry and
/// substitutes the base64 before the plan reaches this handler. The drawing is
/// wrapped in a new paragraph placed before or after the anchor paragraph
/// according to <see cref="InsertImageOp.Position"/>. The drawing is a reference to a
/// shared <see cref="DocumentFormat.OpenXml.Packaging.ImagePart"/>: identical bytes
/// already embedded in the document are reused rather than duplicated.
/// </summary>
internal sealed class InsertImageHandler : IOperationHandler
{
    // 9525 EMU per pixel at 96 DPI.
    private const long EmuPerPixel = 9525;

    public bool CanHandle(PlanOperation operation) =>
        operation is InsertImageOp { Target: TextSpanAnchor };

    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (InsertImageOp)operation;
        var anchor = (TextSpanAnchor)op.Target;

        var hasBytes = !string.IsNullOrEmpty(op.Base64Bytes);
        var hasDocumentId = !string.IsNullOrEmpty(op.ImageDocumentId);

        if (!hasBytes && !hasDocumentId)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "insertImage requires either Base64Bytes or ImageConnectionId+ImageDocumentId.", anchor));

        if (hasBytes && hasDocumentId)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "insertImage cannot mix Base64Bytes with ImageDocumentId; choose one.", anchor));

        if (hasDocumentId && string.IsNullOrEmpty(op.ImageConnectionId))
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "insertImage with ImageDocumentId also requires ImageConnectionId.", anchor));

        if (hasDocumentId)
            // The provider-id path is resolved upstream in OfficeAgentClient; if it
            // reached this handler unresolved, the client did not register a
            // provider for this connection (or wasn't routed through the provider
            // path at all).
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "insertImage with ImageDocumentId must be applied via the provider-backed OfficeAgentClient overload so the image bytes can be resolved.",
                anchor));

        if (!ImagePartTypes.TryGetValue(op.ImageType.ToLowerInvariant(), out _))
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"insertImage ImageType must be one of png, jpeg, gif, bmp, tiff; got '{op.ImageType}'.", anchor));

        if (WordModel.ResolveParagraph(context, anchor.ParaId) is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No paragraph with id '{anchor.ParaId}'.", anchor));

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "insertImage",
            Before = string.Empty,
            After = $"[image {op.WidthPx}x{op.HeightPx}px {op.ImageType}{(op.AltText is null ? "" : $" alt='{op.AltText}'")}]",
            Context = $"{op.Position.ToString().ToLowerInvariant()} paragraph '{anchor.ParaId}'",
            BlastRadius = 1
        });
    }

    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (InsertImageOp)operation;
        var anchor = (TextSpanAnchor)op.Target;

        var paragraph = WordModel.ResolveParagraph(context, anchor.ParaId)
            ?? throw new InvalidOperationException($"Paragraph '{anchor.ParaId}' vanished before apply.");

        var doc = WordModel.Doc(context.Package);
        var main = doc.MainDocumentPart!;
        var imageType = ImagePartTypes[op.ImageType.ToLowerInvariant()];

        var bytes = Convert.FromBase64String(
            op.Base64Bytes
            ?? throw new InvalidOperationException("insertImage Base64Bytes is null at apply time; the ImageDocumentId was not resolved by the client."));

        // An image is a reference to a shared resource: when identical bytes are already
        // embedded in the document, point the new drawing at that existing ImagePart
        // instead of duplicating the bytes in the package.
        var relationshipId = ReferenceOrAddImagePart(main, imageType, bytes);

        // Each drawing needs its own DocProperties id; deriving it from the (now shared)
        // relationship id would collide when the same image is referenced more than once.
        var drawing = BuildInlineDrawing(
            relationshipId,
            op.WidthPx * EmuPerPixel,
            op.HeightPx * EmuPerPixel,
            altText: op.AltText,
            uniqueId: NextDrawingId(context.Package));

        var imageParagraph = new Paragraph(new Run(drawing));

        if (op.Position == InsertPosition.Before)
            paragraph.InsertBeforeSelf(imageParagraph);
        else
            paragraph.InsertAfterSelf(imageParagraph);
    }

    /// <summary>
    /// Returns the relationship id for the image bytes, reusing an already-embedded
    /// <see cref="ImagePart"/> whose content is byte-identical or adding a fresh part
    /// when none matches. Deduplicating keeps a repeated image a single shared resource.
    /// </summary>
    private static string ReferenceOrAddImagePart(MainDocumentPart main, PartTypeInfo imageType, byte[] bytes)
    {
        foreach (var existing in main.ImageParts)
            if (PartContentEquals(existing, bytes))
                return main.GetIdOfPart(existing);

        var imagePart = main.AddImagePart(imageType);
        using (var source = new MemoryStream(bytes, writable: false))
            imagePart.FeedData(source);
        return main.GetIdOfPart(imagePart);
    }

    /// <summary>Returns a DocProperties id one greater than any drawing already present.</summary>
    private static uint NextDrawingId(IOpenXmlPackage package)
    {
        uint max = 0;
        foreach (var (drawing, _) in ImageNodeProvider.EnumerateDrawingsWithHost(package))
        {
            var id = drawing.Descendants<DW.DocProperties>().FirstOrDefault()?.Id?.Value ?? 0U;
            if (id > max) max = id;
        }
        return max + 1U;
    }

    private static bool PartContentEquals(ImagePart part, byte[] candidate)
    {
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        if (stream.CanSeek && stream.Length != candidate.LongLength) return false;

        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return buffer.ToArray().AsSpan().SequenceEqual(candidate);
    }

    private static readonly Dictionary<string, PartTypeInfo> ImagePartTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        ["png"] = ImagePartType.Png,
        ["jpeg"] = ImagePartType.Jpeg,
        ["jpg"] = ImagePartType.Jpeg,
        ["gif"] = ImagePartType.Gif,
        ["bmp"] = ImagePartType.Bmp,
        ["tiff"] = ImagePartType.Tiff
    };

    private static Drawing BuildInlineDrawing(string relationshipId, long widthEmu, long heightEmu, string? altText, uint uniqueId)
    {
        var description = altText ?? string.Empty;
        return new Drawing(
            new DW.Inline(
                new DW.Extent { Cx = widthEmu, Cy = heightEmu },
                new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                new DW.DocProperties { Id = uniqueId, Name = $"Picture {uniqueId}", Description = description },
                new DW.NonVisualGraphicFrameDrawingProperties(
                    new A.GraphicFrameLocks { NoChangeAspect = true }),
                new A.Graphic(
                    new A.GraphicData(
                        new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                                new PIC.NonVisualDrawingProperties { Id = 0U, Name = $"Picture {uniqueId}", Description = description },
                                new PIC.NonVisualPictureDrawingProperties()),
                            new PIC.BlipFill(
                                new A.Blip { Embed = relationshipId, CompressionState = A.BlipCompressionValues.Print },
                                new A.Stretch(new A.FillRectangle())),
                            new PIC.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset { X = 0L, Y = 0L },
                                    new A.Extents { Cx = widthEmu, Cy = heightEmu }),
                                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })))
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }))
            { DistanceFromTop = 0U, DistanceFromBottom = 0U, DistanceFromLeft = 0U, DistanceFromRight = 0U });
    }
}
