namespace OfficeAgent.Core.DocumentProviders;

/// <summary>
/// Maps a document's file extension to the media type providers report on a
/// <see cref="OfficeAgent.Abstractions.DocumentReference"/>.
/// </summary>
/// <remarks>
/// Hosts serve downloads with this value, so an unknown extension falling back to
/// <c>application/octet-stream</c> is not cosmetic - a browser prompts to save an
/// unnamed binary rather than opening the deck in Office. Shared by every provider so a
/// format added to one is not missing from the next.
/// </remarks>
internal static class OfficeContentTypes
{
    private const string Fallback = "application/octet-stream";

    private static readonly Dictionary<string, string> ByExtension =
        new(StringComparer.OrdinalIgnoreCase)
        {
            [".docx"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            [".docm"] = "application/vnd.ms-word.document.macroEnabled.12",
            [".dotx"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.template",
            [".pptx"] = "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            [".pptm"] = "application/vnd.ms-powerpoint.presentation.macroEnabled.12",
            [".potx"] = "application/vnd.openxmlformats-officedocument.presentationml.template",
            [".ppsx"] = "application/vnd.openxmlformats-officedocument.presentationml.slideshow",
            [".xlsx"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            [".xlsm"] = "application/vnd.ms-excel.sheet.macroEnabled.12",
            [".xltx"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.template"
        };

    /// <summary>
    /// Returns the media type for a document name, or <c>application/octet-stream</c>
    /// when the extension is not a known Office format.
    /// </summary>
    public static string ForName(string? name)
    {
        if (string.IsNullOrEmpty(name)) return Fallback;
        var extension = Path.GetExtension(name);
        return extension.Length > 0 && ByExtension.TryGetValue(extension, out var contentType)
            ? contentType
            : Fallback;
    }
}
