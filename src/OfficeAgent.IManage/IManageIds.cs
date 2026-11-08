using System.Globalization;
using System.Text.RegularExpressions;

namespace OfficeAgent.IManage;

/// <summary>Parses iManage document and folder ids and the URLs that carry them.</summary>
internal static class IManageIds
{
    // iManage ids are LIBRARY!number.version for a document and LIBRARY!number for a container.
    private const string LibraryPattern = "[A-Za-z0-9_-]+";

    private static readonly Regex DocumentId = new(
        "^(?<lib>" + LibraryPattern + ")!(?<num>[0-9]{1,18})\\.(?<ver>[0-9]{1,9})$",
        RegexOptions.CultureInvariant);

    private static readonly Regex ContainerId = new(
        "^(?<lib>" + LibraryPattern + ")!(?<num>[0-9]{1,18})$",
        RegexOptions.CultureInvariant);

    private static readonly Regex Library = new("^" + LibraryPattern + "$", RegexOptions.CultureInvariant);

    public static bool IsLibrary(string? value) => value is not null && Library.IsMatch(value);

    public static bool TryParseDocumentId(string value, out IManageDocumentRef document)
    {
        document = default;
        var match = DocumentId.Match(value.Trim());
        if (!match.Success) return false;
        if (!long.TryParse(match.Groups["num"].Value, NumberStyles.None, CultureInfo.InvariantCulture, out var number) ||
            !int.TryParse(match.Groups["ver"].Value, NumberStyles.None, CultureInfo.InvariantCulture, out var version) ||
            number <= 0 || version <= 0)
            return false;
        document = new IManageDocumentRef(match.Groups["lib"].Value, number, version);
        return true;
    }

    public static bool TryParseContainerId(string value, out string library)
    {
        library = string.Empty;
        var match = ContainerId.Match(value.Trim());
        if (!match.Success) return false;
        library = match.Groups["lib"].Value;
        return true;
    }

    public static bool LooksLikeContainerId(string value) => ContainerId.IsMatch(value.Trim());

    /// <summary>
    /// Finds a document id in an iManage Work URL: a path segment such as
    /// <c>/work/web/r/libraries/ACTIVE/documents/ACTIVE!123.1</c> or
    /// <c>/work/link/d/ACTIVE!123.1</c>, a query value, or <c>lib</c>/<c>num</c>/<c>ver</c>
    /// query parameters.
    /// </summary>
    public static bool TryParseDocumentUrl(Uri url, out IManageDocumentRef document)
    {
        document = default;
        foreach (var segment in url.AbsolutePath.Split('/'))
        {
            if (segment.Length == 0) continue;
            if (TryParseDocumentId(Uri.UnescapeDataString(segment), out document))
                return true;
        }

        var query = ParseQuery(url.Query);
        foreach (var value in query.Values)
            if (TryParseDocumentId(value, out document))
                return true;

        if (query.TryGetValue("lib", out var lib) &&
            query.TryGetValue("num", out var num) &&
            query.TryGetValue("ver", out var ver))
            return TryParseDocumentId($"{lib}!{num}.{ver}", out document);
        return false;
    }

    private static Dictionary<string, string> ParseQuery(string query)
    {
        var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var pair in query.TrimStart('?').Split('&'))
        {
            if (pair.Length == 0) continue;
            var equals = pair.IndexOf('=');
            var key = Uri.UnescapeDataString((equals < 0 ? pair : pair.Substring(0, equals)).Replace('+', ' '));
            var value = equals < 0 ? string.Empty : Uri.UnescapeDataString(pair.Substring(equals + 1).Replace('+', ' '));
            if (key.Length > 0 && !values.ContainsKey(key))
                values[key] = value;
        }
        return values;
    }
}
