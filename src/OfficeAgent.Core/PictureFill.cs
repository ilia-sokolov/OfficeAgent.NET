using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeAgent.Abstractions;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeAgent.Core;

/// <summary>
/// Puts a new image into an existing picture's <c>blipFill</c>: the DrawingML both Word and
/// PowerPoint use for the bytes a picture shows and how they meet its frame.
/// </summary>
/// <remarks>
/// <para>
/// Everything about the picture other than the image stays: its frame, position, rotation,
/// outline and shape effects live outside the fill, and the blip's own colour effects
/// (transparency, recolouring) are kept on it. What belongs to the old image goes with it:
/// the crop rectangle, which was cut for the old image's proportions and is recomputed, the
/// SVG reference an SVG picture carries beside its raster fallback, which Office would
/// otherwise go on drawing, and the artistic-effect layers Office caches from the old image.
/// </para>
/// <para>
/// The fit is made with the crop rectangle alone, as Office's own Crop to Fill and Crop to Fit
/// do: a positive crop trims the image to the frame's proportions, and a negative one pads it.
/// The frame is never resized, so the picture keeps the place the template gave it.
/// </para>
/// </remarks>
internal static class PictureFill
{
    private const string RelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// <summary>The blip extension holding <c>asvg:svgBlip</c>, an SVG picture's vector image.</summary>
    private const string SvgExtension = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";

    /// <summary>The blip extension holding <c>a14:imgProps</c>, artistic effects cached from the image.</summary>
    private const string ImagePropertiesExtension = "{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B}";

    private static readonly Dictionary<string, PartTypeInfo> PartTypes =
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

    /// <summary>A replacement image, decoded and measured.</summary>
    internal sealed class Image
    {
        public Image(byte[] bytes, PartTypeInfo partType, int width, int height)
        {
            Bytes = bytes;
            PartType = partType;
            Width = width;
            Height = height;
        }

        public byte[] Bytes { get; }
        public PartTypeInfo PartType { get; }
        public int Width { get; }
        public int Height { get; }
    }

    /// <summary>
    /// Decodes and measures the image a <c>replaceImage</c> carries, or says why it cannot be
    /// used. The fit needs the image's proportions, so bytes whose size cannot be read, or
    /// that are not the declared type, are refused rather than stretched blind.
    /// </summary>
    public static string? Read(ReplaceImageOp op, out Image? image)
    {
        image = null;
        var hasBytes = !string.IsNullOrEmpty(op.Base64Bytes);
        var hasDocumentId = !string.IsNullOrEmpty(op.ImageDocumentId);

        if (hasBytes && hasDocumentId)
            return "replaceImage cannot mix base64Bytes with imageDocumentId; choose one.";
        if (!hasBytes && !hasDocumentId)
            return "replaceImage needs either base64Bytes or an imageConnectionId/imageDocumentId pair.";
        if (hasDocumentId && string.IsNullOrEmpty(op.ImageConnectionId))
            return "replaceImage with imageDocumentId also requires imageConnectionId.";
        if (hasDocumentId)
            return "replaceImage with imageDocumentId must be applied through the provider-backed " +
                   "OfficeAgentClient overload so the image bytes can be resolved.";

        if (!PartTypes.TryGetValue(op.ImageType ?? "png", out var partType))
            return $"Unsupported imageType '{op.ImageType}'. Use png, jpeg, gif, bmp, or tiff.";

        byte[] bytes;
        try
        {
            bytes = Convert.FromBase64String(op.Base64Bytes!);
        }
        catch (FormatException)
        {
            return "replaceImage base64Bytes is not valid base64.";
        }

        if (!TryMeasure(bytes, out var actual, out var width, out var height))
            return "replaceImage could not read the image's width and height. Send a png, jpeg, gif, bmp or tiff file.";
        if (!string.Equals(actual, partType.ContentType, StringComparison.Ordinal))
            return $"The bytes are a {actual.Substring("image/".Length)} image, not the declared imageType '{op.ImageType}'. " +
                   "Declare the type the bytes actually are.";

        image = new Image(bytes, partType, width, height);
        return null;
    }

    /// <summary>
    /// The crop rectangle that makes an image of <paramref name="width"/> by
    /// <paramref name="height"/> pixels meet a frame of the given size, or null when the two
    /// already share proportions. Values are thousandths of a percent of the image, as
    /// <c>a:srcRect</c> counts.
    /// </summary>
    public static A.SourceRectangle? Crop(long frameWidth, long frameHeight, int width, int height, ImageFit fit)
    {
        if (frameWidth <= 0 || frameHeight <= 0 || width <= 0 || height <= 0) return null;

        double frame = (double)frameWidth / frameHeight;
        double image = (double)width / height;
        bool wider = image > frame;

        // Cover trims the long axis to the frame's proportions; contain pads the short one.
        double each = fit == ImageFit.Contain
            ? -((wider ? image / frame : frame / image) - 1) / 2
            : (1 - (wider ? frame / image : image / frame)) / 2;
        var value = (int)Math.Round(each * 100000, MidpointRounding.AwayFromZero);
        if (value == 0) return null;

        // Cover crops the sides of a wider image; contain pads the top and bottom of one.
        bool horizontal = wider == (fit == ImageFit.Cover);
        return horizontal
            ? new A.SourceRectangle { Left = value, Right = value }
            : new A.SourceRectangle { Top = value, Bottom = value };
    }

    /// <summary>
    /// Points <paramref name="blipFill"/> at the image behind <paramref name="relationshipId"/>
    /// with the given crop, stretched to the frame. Returns every relationship id the old fill
    /// referenced, for the caller to release once nothing else uses them.
    /// </summary>
    public static IReadOnlyList<string> Refill(OpenXmlCompositeElement blipFill, string relationshipId, A.SourceRectangle? crop)
    {
        var blip = blipFill.GetFirstChild<A.Blip>()
            ?? throw new InvalidOperationException("The picture has no image to replace.");
        var previous = References(blip);

        blip.Embed = relationshipId;
        blip.Link = null;
        if (blip.GetFirstChild<A.BlipExtensionList>() is { } extensions)
        {
            foreach (var extension in extensions.Elements<A.BlipExtension>().ToList())
                if (extension.Uri?.Value is SvgExtension or ImagePropertiesExtension)
                    extension.Remove();
            if (!extensions.HasChildren) extensions.Remove();
        }

        // CT_BlipFillProperties: blip, srcRect, then the fill mode.
        blipFill.RemoveAllChildren<A.SourceRectangle>();
        if (crop is not null) blip.InsertAfterSelf(crop);

        blipFill.RemoveAllChildren<A.Tile>();
        var stretch = blipFill.GetFirstChild<A.Stretch>();
        if (stretch is null)
        {
            stretch = new A.Stretch();
            ((OpenXmlElement?)crop ?? blip).InsertAfterSelf(stretch);
        }
        stretch.RemoveAllChildren();
        stretch.Append(new A.FillRectangle());

        return previous;
    }

    /// <summary>
    /// Returns the relationship id for the image bytes in <paramref name="owner"/>, reusing a
    /// byte-identical image part already there rather than embedding the bytes twice.
    /// </summary>
    public static string ReferenceOrAddImagePart(OpenXmlPartContainer owner, PartTypeInfo partType, byte[] bytes)
    {
        foreach (var existing in owner.GetPartsOfType<ImagePart>())
            if (ContentEquals(existing, bytes))
                return owner.GetIdOfPart(existing);

        var part = owner.AddNewPart<ImagePart>(partType.ContentType);
        using (var source = new MemoryStream(bytes, writable: false))
            part.FeedData(source);
        return owner.GetIdOfPart(part);
    }

    /// <summary>
    /// Deletes each image relationship of <paramref name="host"/> in <paramref name="relationshipIds"/>
    /// that nothing in the host's XML references any more. Another part's reference to the same
    /// image is its own relationship, which the package keeps.
    /// </summary>
    public static void ReleaseUnreferenced(OpenXmlPart host, IEnumerable<string> relationshipIds)
    {
        var root = host.RootElement;
        if (root is null) return;

        var used = new HashSet<string>(
            new[] { root }.Concat(root.Descendants())
                .SelectMany(element => element.GetAttributes())
                .Where(attribute => attribute.NamespaceUri == RelationshipNamespace && attribute.Value is not null)
                .Select(attribute => attribute.Value!),
            StringComparer.Ordinal);

        foreach (var id in relationshipIds.Distinct(StringComparer.Ordinal))
        {
            if (used.Contains(id)) continue;
            if (host.Parts.Any(pair => pair.RelationshipId == id && pair.OpenXmlPart is ImagePart))
                host.DeletePart(id);
        }
    }

    /// <summary>Every relationship id an element or its descendants reference.</summary>
    private static IReadOnlyList<string> References(OpenXmlElement element) =>
        new[] { element }.Concat(element.Descendants())
            .SelectMany(e => e.GetAttributes())
            .Where(attribute => attribute.NamespaceUri == RelationshipNamespace && !string.IsNullOrEmpty(attribute.Value))
            .Select(attribute => attribute.Value!)
            .Distinct(StringComparer.Ordinal)
            .ToList();

    private static bool ContentEquals(ImagePart part, byte[] candidate)
    {
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        if (stream.CanSeek && stream.Length != candidate.LongLength) return false;

        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return buffer.ToArray().AsSpan().SequenceEqual(candidate);
    }

    /// <summary>
    /// Reads the pixel size from an image header, and the content type the header shows the
    /// bytes to be. Only the header is read; the image is not decoded.
    /// </summary>
    internal static bool TryMeasure(byte[] bytes, out string contentType, out int width, out int height)
    {
        contentType = string.Empty;
        width = height = 0;
        try
        {
            if (bytes.Length >= 24 && bytes[0] == 0x89 && bytes[1] == 0x50 && bytes[2] == 0x4E && bytes[3] == 0x47)
            {
                contentType = ImagePartType.Png.ContentType;
                width = BigEndian32(bytes, 16);
                height = BigEndian32(bytes, 20);
            }
            else if (bytes.Length >= 10 && bytes[0] == 0x47 && bytes[1] == 0x49 && bytes[2] == 0x46)
            {
                contentType = ImagePartType.Gif.ContentType;
                width = bytes[6] | bytes[7] << 8;
                height = bytes[8] | bytes[9] << 8;
            }
            else if (bytes.Length >= 26 && bytes[0] == 0x42 && bytes[1] == 0x4D)
            {
                contentType = ImagePartType.Bmp.ContentType;
                if (LittleEndian32(bytes, 14) == 12)
                {
                    width = bytes[18] | bytes[19] << 8;
                    height = bytes[20] | bytes[21] << 8;
                }
                else
                {
                    width = Math.Abs(LittleEndian32(bytes, 18));
                    height = Math.Abs(LittleEndian32(bytes, 22));
                }
            }
            else if (bytes.Length >= 4 && bytes[0] == 0xFF && bytes[1] == 0xD8 && bytes[2] == 0xFF)
            {
                contentType = ImagePartType.Jpeg.ContentType;
                MeasureJpeg(bytes, out width, out height);
            }
            else if (bytes.Length >= 8 && ((bytes[0] == 0x49 && bytes[1] == 0x49) || (bytes[0] == 0x4D && bytes[1] == 0x4D)))
            {
                contentType = ImagePartType.Tiff.ContentType;
                MeasureTiff(bytes, out width, out height);
            }
        }
        catch (IndexOutOfRangeException)
        {
            return false;
        }

        return contentType.Length > 0 && width > 0 && height > 0;
    }

    private static void MeasureJpeg(byte[] bytes, out int width, out int height)
    {
        width = height = 0;
        var position = 2;
        while (position + 9 < bytes.Length)
        {
            if (bytes[position] != 0xFF) return;
            var marker = bytes[position + 1];
            if (marker == 0xFF) { position++; continue; }

            // Markers without a length: the restart markers, start of image, and TEM.
            if (marker is >= 0xD0 and <= 0xD8 or 0x01) { position += 2; continue; }

            var length = bytes[position + 2] << 8 | bytes[position + 3];
            // Every start-of-frame marker carries the size; C4, C8 and CC are other segments.
            if (marker is >= 0xC0 and <= 0xCF and not 0xC4 and not 0xC8 and not 0xCC)
            {
                height = bytes[position + 5] << 8 | bytes[position + 6];
                width = bytes[position + 7] << 8 | bytes[position + 8];
                return;
            }
            if (length < 2) return;
            position += 2 + length;
        }
    }

    private static void MeasureTiff(byte[] bytes, out int width, out int height)
    {
        width = height = 0;
        var little = bytes[0] == 0x49;
        int Read16(int at) => little ? bytes[at] | bytes[at + 1] << 8 : bytes[at] << 8 | bytes[at + 1];
        int Read32(int at) => little ? LittleEndian32(bytes, at) : BigEndian32(bytes, at);

        var directory = Read32(4);
        if (directory < 8 || directory + 2 > bytes.Length) return;
        var entries = Read16(directory);
        for (var index = 0; index < entries; index++)
        {
            var entry = directory + 2 + index * 12;
            if (entry + 12 > bytes.Length) return;
            var tag = Read16(entry);
            var type = Read16(entry + 2);
            var value = type == 3 ? Read16(entry + 8) : Read32(entry + 8);
            if (tag == 256) width = value;
            else if (tag == 257) height = value;
        }
    }

    private static int BigEndian32(byte[] bytes, int at) =>
        bytes[at] << 24 | bytes[at + 1] << 16 | bytes[at + 2] << 8 | bytes[at + 3];

    private static int LittleEndian32(byte[] bytes, int at) =>
        bytes[at] | bytes[at + 1] << 8 | bytes[at + 2] << 16 | bytes[at + 3] << 24;
}
