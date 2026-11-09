using System.Security.Cryptography;
using System.Text;

namespace OfficeAgent.Core.DocumentProviders;

/// <summary>
/// The listing rules every provider applies the same way: the page-size clamp, which
/// extensions a query may list, what counts as a template, and how a continuation token is
/// tied to the query that produced it.
/// </summary>
internal static class DocumentListingRules
{
    /// <summary>The largest page a provider returns, whatever the query asks for.</summary>
    public const int MaximumPageSize = 500;

    private static readonly HashSet<string> TemplateExtensions =
        new(StringComparer.OrdinalIgnoreCase) { ".dotx", ".potx", ".xltx" };

    // Macro-enabled packages are never listed, even when a host allowed one: an agent that
    // finds a document through a listing should not be led to a file that carries code.
    private static readonly HashSet<string> MacroEnabledExtensions =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ".docm", ".dotm", ".pptm", ".potm", ".ppsm", ".sldm", ".ppam", ".xlsm", ".xltm", ".xlsb", ".xlam"
        };

    /// <summary>Clamps a requested page size to 1..<see cref="MaximumPageSize"/>.</summary>
    public static int PageSize(DocumentListQuery query) =>
        Math.Max(1, Math.Min(MaximumPageSize, query.MaximumResults));

    /// <summary>Whether a file name has a template extension.</summary>
    public static bool IsTemplateName(string name) => TemplateExtensions.Contains(Path.GetExtension(name));

    /// <summary>
    /// The extensions a query lists: the ones it asks for that the connection allows, or every
    /// allowed one when it asks for none. Macro-enabled extensions are always removed.
    /// </summary>
    public static HashSet<string> Extensions(DocumentListQuery query, IEnumerable<string> allowed)
    {
        var permitted = new HashSet<string>(
            allowed.Select(Normalize).Where(extension => !MacroEnabledExtensions.Contains(extension)),
            StringComparer.OrdinalIgnoreCase);
        if (query.Extensions is null || query.Extensions.Count == 0) return permitted;

        var requested = new HashSet<string>(
            query.Extensions.Where(extension => !string.IsNullOrWhiteSpace(extension)).Select(Normalize),
            StringComparer.OrdinalIgnoreCase);
        // A list of blanks narrows nothing; it is read as no filter rather than as "list nothing".
        if (requested.Count == 0) return permitted;
        permitted.IntersectWith(requested);
        return permitted;
    }

    /// <summary>Whether a file name passes the query's extension set and name filter.</summary>
    public static bool Matches(string name, HashSet<string> extensions, DocumentListQuery query) =>
        extensions.Contains(Path.GetExtension(name)) &&
        (string.IsNullOrEmpty(query.NameContains) ||
         name.IndexOf(query.NameContains!.Trim(), StringComparison.OrdinalIgnoreCase) >= 0);

    /// <summary>
    /// A short fingerprint of everything in a query except its page size and token. A token
    /// carries it, so a token reused with a different folder or filter is refused rather than
    /// resuming somewhere that query never reached.
    /// </summary>
    public static string Fingerprint(DocumentListQuery query, string scope)
    {
        var extensions = query.Extensions is null
            ? string.Empty
            : string.Join(",", query.Extensions.Select(e => Normalize(e ?? string.Empty).ToLowerInvariant()).OrderBy(e => e, StringComparer.Ordinal));
        var text = string.Join("\n", scope, query.Templates ? "t" : "f", query.Recursive ? "r" : "-",
            (query.Folder ?? string.Empty).Trim(), (query.NameContains ?? string.Empty).Trim().ToLowerInvariant(), extensions);
        using var sha = SHA256.Create();
        return Convert.ToBase64String(sha.ComputeHash(Encoding.UTF8.GetBytes(text)), 0, 12);
    }

    /// <summary>Encodes a token payload as URL-safe base64.</summary>
    public static string EncodeToken(string payload) =>
        Convert.ToBase64String(Encoding.UTF8.GetBytes(payload)).TrimEnd('=').Replace('+', '-').Replace('/', '_');

    /// <summary>Decodes a token from <see cref="EncodeToken"/>, or <see langword="null"/> when it is not one.</summary>
    public static string? DecodeToken(string token)
    {
        var text = token.Trim().Replace('-', '+').Replace('_', '/');
        switch (text.Length % 4)
        {
            case 2: text += "=="; break;
            case 3: text += "="; break;
            case 1: return null;
        }

        try
        {
            return Encoding.UTF8.GetString(Convert.FromBase64String(text));
        }
        catch (FormatException)
        {
            return null;
        }
    }

    private static string Normalize(string extension)
    {
        var trimmed = extension.Trim();
        return trimmed.StartsWith(".", StringComparison.Ordinal) ? trimmed : "." + trimmed;
    }
}
