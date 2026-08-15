using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeAgent.Abstractions;

namespace OfficeAgent.PowerPoint;

/// <summary>
/// The bits every handler that takes image bytes needs: decoding them, naming their part
/// type, and refusing the ways a caller can supply them wrongly.
/// </summary>
/// <remarks>
/// The image-type vocabulary matches the Word module's, so one plan vocabulary serves both
/// formats and an agent does not have to learn which words each accepts.
/// </remarks>
internal static class SlideImages
{
    public static bool TryDecode(string base64, out byte[] bytes)
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

    /// <summary>Maps the plan's image-type vocabulary onto Open XML part types.</summary>
    public static PartTypeInfo? PartTypeFor(string? imageType)
    {
        if (string.IsNullOrEmpty(imageType)) return ImagePartType.Png;
        return PartTypes.TryGetValue(imageType!, out var type) ? type : null;
    }

    /// <summary>
    /// Says why the image on a <c>backgroundImage</c> cannot be used, or null when it can.
    /// The provider-id form is resolved to bytes upstream by the client, so a plan that
    /// still carries one here was never routed through a provider-backed client.
    /// </summary>
    public static string? Unsupported(BackgroundImageOp op)
    {
        var hasBytes = !string.IsNullOrEmpty(op.Base64Bytes);
        var hasDocumentId = !string.IsNullOrEmpty(op.ImageDocumentId);

        if (hasBytes && hasDocumentId)
            return "backgroundImage cannot mix base64Bytes with imageDocumentId; choose one.";

        if (hasDocumentId && string.IsNullOrEmpty(op.ImageConnectionId))
            return "backgroundImage with imageDocumentId also requires imageConnectionId.";

        if (hasDocumentId)
            return "backgroundImage with imageDocumentId must be applied through the provider-backed " +
                   "OfficeAgentClient overload so the image bytes can be resolved.";

        if (!hasBytes) return null;   // clearing the background

        if (!TryDecode(op.Base64Bytes!, out _))
            return "backgroundImage base64Bytes is not valid base64.";

        if (PartTypeFor(op.ImageType) is null)
            return $"Unsupported imageType '{op.ImageType}'. Use png, jpeg, gif, bmp, or tiff.";

        return null;
    }

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
}
