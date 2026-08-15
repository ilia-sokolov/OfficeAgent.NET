using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeAgent.Abstractions;

namespace OfficeAgent.Word;

/// <summary>
/// Decoding image bytes and naming their part type, shared by the handlers that take them.
/// </summary>
/// <remarks>
/// The accepted type names match the PowerPoint module's, so one plan vocabulary serves both
/// formats.
/// </remarks>
internal static class WordImages
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

    public static PartTypeInfo? PartTypeFor(string? imageType)
    {
        if (string.IsNullOrEmpty(imageType)) return ImagePartType.Png;
        return PartTypes.TryGetValue(imageType!, out var type) ? type : null;
    }

    /// <summary>
    /// Converts an opacity of 0-1 to the thousandths of a percent DrawingML counts in. Full
    /// strength returns null, because writing <c>100%</c> says no more than writing nothing.
    /// </summary>
    public static int? Alpha(double? opacity) =>
        opacity is null or >= 1 ? null : (int)Math.Round(Math.Max(0, opacity!.Value) * 100000);

    /// <summary>Says why the image cannot be used, or null when it can.</summary>
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
