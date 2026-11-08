// Compiled into OfficeAgent.AzureBlob and OfficeAgent.AmazonS3 as linked source, so the two
// object-store providers apply one set of naming rules without a shared package dependency.

namespace OfficeAgent.ObjectStorage;

/// <summary>Why a name relative to a connection's prefix cannot be used.</summary>
internal enum ObjectKeyProblem
{
    None,

    /// <summary>Empty, too long, an empty segment, a trailing slash, or a control character.</summary>
    Invalid,

    /// <summary>A leading slash, a backslash, or a <c>.</c>/<c>..</c> segment: a way out of the prefix.</summary>
    Escapes
}

/// <summary>
/// Naming rules for object keys and blob names. A connection is one container or bucket plus an
/// optional prefix; every name a caller supplies is relative to that prefix and must stay inside it.
/// </summary>
internal static class ObjectKeys
{
    /// <summary>Blob names and S3 keys are limited to 1024 characters (bytes for S3).</summary>
    internal const int MaximumKeyLength = 1024;

    /// <summary>
    /// Normalizes a configured prefix to <c>""</c> or <c>"segment/…/"</c>. A prefix that is not a
    /// plain relative path is a host configuration error, so it throws rather than reporting.
    /// </summary>
    internal static string NormalizePrefix(string? prefix, string parameterName)
    {
        if (string.IsNullOrWhiteSpace(prefix)) return string.Empty;
        var trimmed = prefix!.Trim().Trim('/');
        if (trimmed.Length == 0) return string.Empty;
        if (Check(trimmed) != ObjectKeyProblem.None)
            throw new ArgumentException(
                "The prefix must be a relative path without '.', '..', empty segments, backslashes, or control characters.",
                parameterName);
        return trimmed + "/";
    }

    /// <summary>Checks a name relative to the connection's prefix.</summary>
    internal static ObjectKeyProblem Check(string relative)
    {
        if (string.IsNullOrEmpty(relative) || relative.Length > MaximumKeyLength)
            return ObjectKeyProblem.Invalid;
        if (relative[0] == '/' || relative.IndexOf('\\') >= 0)
            return ObjectKeyProblem.Escapes;
        foreach (var segment in relative.Split('/'))
        {
            if (segment == "." || segment == "..") return ObjectKeyProblem.Escapes;
            if (segment.Length == 0) return ObjectKeyProblem.Invalid;
        }
        foreach (var character in relative)
            if (char.IsControl(character)) return ObjectKeyProblem.Invalid;
        return ObjectKeyProblem.None;
    }

    /// <summary>The last segment of a key: the document's file name.</summary>
    internal static string Leaf(string key)
    {
        var slash = key.LastIndexOf('/');
        return slash < 0 ? key : key.Substring(slash + 1);
    }

    /// <summary>Everything up to and including the last slash, or <c>""</c>.</summary>
    internal static string Parent(string key)
    {
        var slash = key.LastIndexOf('/');
        return slash < 0 ? string.Empty : key.Substring(0, slash + 1);
    }

    /// <summary>
    /// The versioned sibling names for a file, in order, matching the filesystem and SharePoint
    /// providers: <c>contract.docx</c> → <c>contract.v2.docx</c>, <c>contract.v3.docx</c>, …; and
    /// <c>contract.v4.docx</c> → <c>contract.v5.docx</c>, ….
    /// </summary>
    internal static IEnumerable<string> VersionedNames(string fileName)
    {
        var extension = Path.GetExtension(fileName);
        var stem = Path.GetFileNameWithoutExtension(fileName);
        var (baseStem, startVersion) = ParseVersionedStem(stem);
        for (var n = Math.Max(2, startVersion + 1); n < startVersion + 1000; n++)
            yield return $"{baseStem}.v{n}{extension}";
    }

    private static (string Stem, int Version) ParseVersionedStem(string stem)
    {
        var dot = stem.LastIndexOf('.');
        if (dot <= 0 || dot == stem.Length - 1) return (stem, 1);
        var tail = stem.Substring(dot + 1);
        if (tail.Length < 2 || tail[0] is not ('v' or 'V')) return (stem, 1);
        if (int.TryParse(tail.Substring(1), System.Globalization.NumberStyles.None,
                System.Globalization.CultureInfo.InvariantCulture, out var version) && version >= 1)
            return (stem.Substring(0, dot), version);
        return (stem, 1);
    }

    /// <summary>Whether a registration id is an opaque token rather than something path-like.</summary>
    internal static bool IsOpaqueId(string itemId) =>
        itemId.IndexOf('/') < 0 && itemId.IndexOf('\\') < 0 && itemId.IndexOf("..", StringComparison.Ordinal) < 0;

    internal static string NormalizeExtension(string extension) =>
        extension.StartsWith(".", StringComparison.Ordinal) ? extension : "." + extension;

    /// <summary>Lowercase hex SHA-256, the content hash a write-recovery locator carries.</summary>
    internal static string Sha256(byte[] bytes)
    {
        using var sha = System.Security.Cryptography.SHA256.Create();
        return string.Concat(sha.ComputeHash(bytes).Select(b => b.ToString("x2", System.Globalization.CultureInfo.InvariantCulture)));
    }

    /// <summary>
    /// Reads a stream into memory and stops as soon as it exceeds <paramref name="maximumBytes"/>,
    /// so an object larger than the connection allows is never held in full.
    /// </summary>
    internal static async Task<byte[]?> ReadBoundedAsync(Stream stream, long maximumBytes, CancellationToken cancellationToken)
    {
        using var copy = new MemoryStream();
        var buffer = new byte[81920];
        while (true)
        {
            var read = await stream.ReadAsync(buffer, 0, buffer.Length, cancellationToken).ConfigureAwait(false);
            if (read == 0) break;
            if (copy.Length + read > maximumBytes) return null;
            copy.Write(buffer, 0, read);
        }
        return copy.ToArray();
    }
}
