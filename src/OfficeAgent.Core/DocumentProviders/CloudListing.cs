using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace OfficeAgent.Core.DocumentProviders;

/// <summary>Why a listing folder cannot be used.</summary>
internal enum ListingFolderProblem
{
    None,

    /// <summary>An empty segment, a control character, or more than <see cref="CloudListing.MaximumFolderSegments"/> segments.</summary>
    Invalid,

    /// <summary>A backslash or a <c>.</c>/<c>..</c> segment: a way out of the listing root.</summary>
    Escapes
}

/// <summary>One item a cloud folder listing returned: a document, a subfolder, or something to pass over.</summary>
internal sealed class CloudListingChild
{
    /// <summary>For a document, the source the provider registers it by; for a folder, the id or prefix that lists it.</summary>
    public string Id { get; init; } = string.Empty;

    /// <summary>The file or folder name.</summary>
    public string Name { get; init; } = string.Empty;

    /// <summary>Whether the item is a folder, which a recursive listing descends into.</summary>
    public bool IsFolder { get; init; }

    /// <summary>
    /// Whether the item can never be registered on this connection, such as a native Google
    /// file, a shortcut, a trashed file, or one over the size cap. It is passed over silently.
    /// </summary>
    public bool Unusable { get; init; }

    /// <summary>The size in bytes, when the service reports it.</summary>
    public long? Length { get; init; }

    /// <summary>When the item last changed, when the service reports it.</summary>
    public DateTimeOffset? LastModified { get; init; }
}

/// <summary>One page of a folder's children, in the order the service returned them.</summary>
internal sealed class CloudListingPage
{
    public IReadOnlyList<CloudListingChild> Items { get; init; } = Array.Empty<CloudListingChild>();

    /// <summary>The service's token for the folder's next page, or <see langword="null"/> after the last.</summary>
    public string? NextToken { get; init; }
}

/// <summary>Lists one page of a folder's children. The same folder and token must return the same page.</summary>
internal delegate Task<CloudListingPage> CloudChildrenFetcher(string folderId, string? pageToken, CancellationToken cancellationToken);

/// <summary>A folder a listing will read: its service id or prefix, its display path, and its depth below the listing root.</summary>
internal sealed class CloudListingFolder
{
    public CloudListingFolder(string id, string? path, int depth)
    {
        Id = id;
        Path = path;
        Depth = depth;
    }

    public string Id { get; }

    /// <summary>The display folder relative to the listing root, or <see langword="null"/> at the root.</summary>
    public string? Path { get; }

    public int Depth { get; }
}

/// <summary>
/// A <see cref="DocumentListQuery"/> checked against one connection: the clamped page size, the
/// extensions that may be listed, the folder segments, and a fingerprint that binds a continuation
/// token to the query that produced it.
/// </summary>
internal sealed class CloudListingRequest
{
    private CloudListingRequest(
        int maximumResults, HashSet<string> extensions, string? nameContains, bool recursive, bool templates,
        IReadOnlyList<string> folderSegments)
    {
        MaximumResults = maximumResults;
        Extensions = extensions;
        NameContains = nameContains;
        Recursive = recursive;
        Templates = templates;
        FolderSegments = folderSegments;
        Fingerprint = string.Join("\n",
            templates ? "t" : "c",
            recursive ? "r" : "f",
            string.Join("/", folderSegments),
            nameContains?.ToUpperInvariant() ?? string.Empty,
            string.Join(",", extensions.Select(e => e.ToLowerInvariant()).OrderBy(e => e, StringComparer.Ordinal)));
    }

    public int MaximumResults { get; }

    /// <summary>The extensions to list: the connection's allow-list, narrowed by the query, never macro-enabled.</summary>
    public HashSet<string> Extensions { get; }

    public string? NameContains { get; }

    public bool Recursive { get; }

    public bool Templates { get; }

    /// <summary>The query's folder, split into segments below the listing root; empty for the root.</summary>
    public IReadOnlyList<string> FolderSegments { get; }

    /// <summary>The display path of the listed folder, or <see langword="null"/> for the listing root.</summary>
    public string? FolderPath => FolderSegments.Count == 0 ? null : string.Join("/", FolderSegments);

    public string Fingerprint { get; }

    /// <summary>
    /// Checks a query against a connection's allow-list. Returns <see langword="null"/> and the
    /// problem when the folder cannot be listed.
    /// </summary>
    public static CloudListingRequest? Create(
        DocumentListQuery query, IEnumerable<string> allowedExtensions, out ListingFolderProblem problem)
    {
        if (query is null) throw new ArgumentNullException(nameof(query));
        problem = CloudListing.SplitFolder(query.Folder, out var segments);
        if (problem != ListingFolderProblem.None) return null;

        var allowed = new HashSet<string>(
            allowedExtensions.Select(CloudListing.NormalizeExtension).Where(e => !CloudListing.IsMacroEnabled(e)),
            StringComparer.OrdinalIgnoreCase);
        if (query.Extensions is { Count: > 0 } requested)
            allowed.IntersectWith(requested.Where(e => !string.IsNullOrWhiteSpace(e)).Select(e => CloudListing.NormalizeExtension(e.Trim())));

        var name = string.IsNullOrWhiteSpace(query.NameContains) ? null : query.NameContains!.Trim();
        return new CloudListingRequest(
            CloudListing.ClampResults(query.MaximumResults), allowed, name, query.Recursive, query.Templates, segments);
    }

    /// <summary>Whether a document child belongs in the listing.</summary>
    public bool Accepts(CloudListingChild child) =>
        !child.IsFolder &&
        !child.Unusable &&
        Extensions.Contains(CloudListing.NormalizeExtension(Path.GetExtension(child.Name))) &&
        (NameContains is null || child.Name.IndexOf(NameContains, StringComparison.OrdinalIgnoreCase) >= 0);
}

/// <summary>
/// Where a cloud listing stands between pages: the folder being read, the service's page token and
/// how many of that page's items were already consumed, and the subfolders still to read. A caller
/// only ever sees it encrypted and authenticated, so a token cannot be read or edited to point the
/// listing outside the connection.
/// </summary>
internal sealed class CloudListingCursor
{
    // A per-process key: a token is valid only in the process that issued it, and a restart
    // invalidates every outstanding token rather than trusting one it cannot verify.
    private static readonly byte[] EncryptionKey = RandomBytes(32);
    private static readonly byte[] SigningKey = RandomBytes(32);

    public CloudListingFolder? Current { get; set; }

    public string? PageToken { get; set; }

    public int Skip { get; set; }

    public List<CloudListingFolder> Pending { get; } = new();

    /// <summary>A provider value carried between pages, such as the shared drive a Drive folder is in.</summary>
    public string? Context { get; set; }

    /// <summary>Encrypts and signs the cursor, bound to one connection and query.</summary>
    public string Protect(string connectionId, string fingerprint)
    {
        var plain = Serialize();
        byte[] iv;
        byte[] cipher;
        using (var aes = Aes.Create())
        {
            aes.Key = EncryptionKey;
            aes.GenerateIV();
            iv = aes.IV;
            using var encryptor = aes.CreateEncryptor();
            cipher = encryptor.TransformFinalBlock(plain, 0, plain.Length);
        }
        var body = new byte[iv.Length + cipher.Length];
        Buffer.BlockCopy(iv, 0, body, 0, iv.Length);
        Buffer.BlockCopy(cipher, 0, body, iv.Length, cipher.Length);
        return Base64Url(body) + "." + Base64Url(Sign(body, connectionId, fingerprint));
    }

    /// <summary>
    /// Reads a token issued by <see cref="Protect"/> for the same connection and query. Returns
    /// <see langword="false"/> for anything else: another query's token, an edited one, or one
    /// issued before the process restarted.
    /// </summary>
    public static bool TryUnprotect(string token, string connectionId, string fingerprint, out CloudListingCursor cursor)
    {
        cursor = new CloudListingCursor();
        if (string.IsNullOrWhiteSpace(token) || token.Length > CloudListing.MaximumTokenLength) return false;
        var dot = token.IndexOf('.');
        if (dot <= 0 || dot == token.Length - 1) return false;

        var body = FromBase64Url(token.Substring(0, dot));
        var mac = FromBase64Url(token.Substring(dot + 1));
        if (body is null || mac is null || body.Length <= 16) return false;
        if (!FixedTimeEquals(mac, Sign(body, connectionId, fingerprint))) return false;

        try
        {
            byte[] plain;
            using (var aes = Aes.Create())
            {
                aes.Key = EncryptionKey;
                var iv = new byte[16];
                Buffer.BlockCopy(body, 0, iv, 0, iv.Length);
                aes.IV = iv;
                using var decryptor = aes.CreateDecryptor();
                plain = decryptor.TransformFinalBlock(body, iv.Length, body.Length - iv.Length);
            }
            return Deserialize(plain, cursor);
        }
        catch (Exception ex) when (ex is CryptographicException or JsonException or InvalidOperationException or FormatException)
        {
            return false;
        }
    }

    private byte[] Serialize()
    {
        using var stream = new MemoryStream();
        using (var writer = new Utf8JsonWriter(stream))
        {
            writer.WriteStartObject();
            if (Current is not null)
            {
                writer.WritePropertyName("c");
                WriteFolder(writer, Current);
            }
            if (PageToken is not null) writer.WriteString("t", PageToken);
            writer.WriteNumber("s", Skip);
            writer.WriteStartArray("q");
            foreach (var folder in Pending) WriteFolder(writer, folder);
            writer.WriteEndArray();
            if (Context is not null) writer.WriteString("x", Context);
            writer.WriteEndObject();
        }
        return stream.ToArray();
    }

    private static void WriteFolder(Utf8JsonWriter writer, CloudListingFolder folder)
    {
        writer.WriteStartObject();
        writer.WriteString("i", folder.Id);
        if (folder.Path is not null) writer.WriteString("p", folder.Path);
        writer.WriteNumber("d", folder.Depth);
        writer.WriteEndObject();
    }

    private static bool Deserialize(byte[] plain, CloudListingCursor cursor)
    {
        using var json = JsonDocument.Parse(plain);
        var root = json.RootElement;
        if (root.ValueKind != JsonValueKind.Object) return false;
        if (root.TryGetProperty("c", out var current)) cursor.Current = ReadFolder(current);
        if (root.TryGetProperty("t", out var page)) cursor.PageToken = page.GetString();
        cursor.Skip = root.GetProperty("s").GetInt32();
        foreach (var folder in root.GetProperty("q").EnumerateArray())
            cursor.Pending.Add(ReadFolder(folder));
        if (root.TryGetProperty("x", out var context)) cursor.Context = context.GetString();
        return cursor.Skip >= 0;
    }

    private static CloudListingFolder ReadFolder(JsonElement element) => new(
        element.GetProperty("i").GetString() ?? string.Empty,
        element.TryGetProperty("p", out var path) ? path.GetString() : null,
        element.GetProperty("d").GetInt32());

    private static byte[] Sign(byte[] body, string connectionId, string fingerprint)
    {
        using var hmac = new HMACSHA256(SigningKey);
        var binding = Encoding.UTF8.GetBytes(connectionId + "\n" + fingerprint + "\n");
        var signed = new byte[binding.Length + body.Length];
        Buffer.BlockCopy(binding, 0, signed, 0, binding.Length);
        Buffer.BlockCopy(body, 0, signed, binding.Length, body.Length);
        return hmac.ComputeHash(signed);
    }

    private static bool FixedTimeEquals(byte[] left, byte[] right)
    {
        if (left.Length != right.Length) return false;
        var difference = 0;
        for (var i = 0; i < left.Length; i++) difference |= left[i] ^ right[i];
        return difference == 0;
    }

    private static byte[] RandomBytes(int count)
    {
        var bytes = new byte[count];
        using var random = RandomNumberGenerator.Create();
        random.GetBytes(bytes);
        return bytes;
    }

    private static string Base64Url(byte[] bytes) =>
        Convert.ToBase64String(bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_');

    private static byte[]? FromBase64Url(string text)
    {
        var base64 = text.Replace('-', '+').Replace('_', '/');
        base64 = base64.PadRight(base64.Length + (4 - base64.Length % 4) % 4, '=');
        try
        {
            return Convert.FromBase64String(base64);
        }
        catch (FormatException)
        {
            return null;
        }
    }
}

/// <summary>
/// The listing rules and bounds the Google Drive, Box, Azure Blob, Amazon S3 and iManage
/// providers share, and the folder walk they all run over their own children call.
/// </summary>
internal static class CloudListing
{
    /// <summary>The most entries one page returns.</summary>
    internal const int MaximumPageSize = 500;

    /// <summary>The most service requests one <c>ListAsync</c> call makes, folder lookups included.</summary>
    internal const int MaximumServiceCalls = 10;

    /// <summary>How many folder levels below the listed folder a recursive listing descends.</summary>
    internal const int MaximumDepth = 5;

    /// <summary>
    /// The most subfolders a recursive listing holds to read later. Further subfolders found while
    /// the queue is full are not listed, which keeps a continuation token small.
    /// </summary>
    internal const int MaximumPendingFolders = 50;

    /// <summary>The most segments a query's folder may have.</summary>
    internal const int MaximumFolderSegments = 16;

    /// <summary>The page size asked of the service on every request, so a token's page can be read again.</summary>
    internal const int ServicePageSize = 200;

    /// <summary>The longest continuation token accepted back.</summary>
    internal const int MaximumTokenLength = 64 * 1024;

    private static readonly HashSet<string> MacroEnabled = new(StringComparer.OrdinalIgnoreCase)
    {
        ".docm", ".dotm", ".xlsm", ".xltm", ".xlam", ".pptm", ".potm", ".ppsm", ".ppam", ".sldm"
    };

    private static readonly HashSet<string> TemplateExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".dotx", ".potx", ".xltx"
    };

    /// <summary>What a provider reports for a continuation token it cannot use.</summary>
    internal const string InvalidTokenMessage =
        "The continuation token does not belong to this listing, or was issued before the server restarted. List again without it.";

    /// <summary>What a provider reports when resolving the query's folder would spend the page's request budget.</summary>
    internal const string TooDeepMessage =
        "The listing folder takes more requests to find than one page may make. List a shallower folder.";

    internal static int ClampResults(int requested) => Math.Min(Math.Max(requested, 1), MaximumPageSize);

    /// <summary>The error code and message a provider reports for a folder <see cref="SplitFolder"/> refused.</summary>
    internal static (ProviderErrorCode Code, string Message) FolderError(ListingFolderProblem problem) =>
        problem == ListingFolderProblem.Escapes
            ? (ProviderErrorCode.AccessDenied,
                "A listing folder must stay inside the listed folder: no '.', '..' or backslash segments.")
            : (ProviderErrorCode.InvalidArgument,
                $"A listing folder is a relative path of at most {MaximumFolderSegments} non-empty names, such as 'Clients/Acme'.");

    /// <summary>
    /// The template location a connection reports: the configured one, or
    /// <see cref="TemplateLocationKind.Connection"/> when a templates folder is set without one. A
    /// location without a folder, or <see cref="TemplateLocationKind.None"/> with one, is a host
    /// configuration error, so it throws rather than reporting.
    /// </summary>
    internal static TemplateLocationKind TemplateLocationFor(bool hasFolder, TemplateLocationKind? configured, string parameterName)
    {
        if (configured is { } kind && !Enum.IsDefined(typeof(TemplateLocationKind), kind))
            throw new ArgumentException($"TemplateLocation {(int)kind} is not a defined template location.", parameterName);
        if (!hasFolder)
        {
            if (configured is { } set && set != TemplateLocationKind.None)
                throw new ArgumentException("TemplateLocation is set but TemplatesFolder is empty; set both or neither.", parameterName);
            return TemplateLocationKind.None;
        }
        if (configured == TemplateLocationKind.None)
            throw new ArgumentException(
                "TemplateLocation is None but TemplatesFolder is set; leave TemplateLocation unset or name where the folder sits.",
                parameterName);
        return configured ?? TemplateLocationKind.Connection;
    }

    internal static string NormalizeExtension(string extension) =>
        extension.StartsWith(".", StringComparison.Ordinal) ? extension : "." + extension;

    /// <summary>Whether an extension names a macro-enabled Office format, which is never listed.</summary>
    internal static bool IsMacroEnabled(string extension) => MacroEnabled.Contains(NormalizeExtension(extension));

    /// <summary>Whether a file name has a template extension.</summary>
    internal static bool IsTemplateName(string name) => TemplateExtensions.Contains(Path.GetExtension(name));

    /// <summary>
    /// Splits a query folder into segments. Leading and trailing slashes are ignored, so
    /// <c>/</c> and an empty folder both name the listing root.
    /// </summary>
    internal static ListingFolderProblem SplitFolder(string? folder, out IReadOnlyList<string> segments)
    {
        segments = Array.Empty<string>();
        if (string.IsNullOrWhiteSpace(folder)) return ListingFolderProblem.None;
        var trimmed = folder!.Trim();
        if (trimmed.IndexOf('\\') >= 0) return ListingFolderProblem.Escapes;
        trimmed = trimmed.Trim('/');
        if (trimmed.Length == 0) return ListingFolderProblem.None;

        var parts = trimmed.Split('/');
        foreach (var part in parts)
        {
            if (part == "." || part == "..") return ListingFolderProblem.Escapes;
            if (part.Length == 0 || part.Any(char.IsControl)) return ListingFolderProblem.Invalid;
        }
        if (parts.Length > MaximumFolderSegments) return ListingFolderProblem.Invalid;
        segments = parts;
        return ListingFolderProblem.None;
    }

    /// <summary>
    /// Reads folders breadth first from where <paramref name="cursor"/> stands until the page is
    /// full, the service-call budget is spent, or nothing is left. A service page is never skipped
    /// part-read: when the page fills mid-way, the cursor keeps the service token and the count of
    /// items consumed, and the next call reads the same service page again from there.
    /// </summary>
    internal static async Task<DocumentListPage> WalkAsync(
        CloudListingRequest request,
        CloudListingCursor cursor,
        CloudChildrenFetcher fetch,
        int callsLeft,
        string connectionId,
        CancellationToken cancellationToken)
    {
        var items = new List<DocumentListEntry>();
        if (request.Extensions.Count > 0)
        {
            while (items.Count < request.MaximumResults && callsLeft > 0)
            {
                if (cursor.Current is null)
                {
                    if (cursor.Pending.Count == 0) break;
                    cursor.Current = cursor.Pending[0];
                    cursor.Pending.RemoveAt(0);
                    cursor.PageToken = null;
                    cursor.Skip = 0;
                }

                var folder = cursor.Current;
                var page = await fetch(folder.Id, cursor.PageToken, cancellationToken).ConfigureAwait(false);
                callsLeft--;

                var full = false;
                for (var index = cursor.Skip; index < page.Items.Count; index++)
                {
                    if (items.Count == request.MaximumResults)
                    {
                        cursor.Skip = index;
                        full = true;
                        break;
                    }

                    var child = page.Items[index];
                    if (child.IsFolder)
                    {
                        if (request.Recursive && folder.Depth < MaximumDepth && cursor.Pending.Count < MaximumPendingFolders)
                            cursor.Pending.Add(new CloudListingFolder(child.Id, Join(folder.Path, child.Name), folder.Depth + 1));
                        continue;
                    }
                    if (request.Accepts(child))
                        items.Add(Entry(child, folder.Path, request.Templates));
                }
                if (full) break;

                cursor.Skip = 0;
                if (page.NextToken is null)
                    cursor.Current = null;
                else
                    cursor.PageToken = page.NextToken;
            }
        }
        else
        {
            // Nothing the connection allows was asked for: the listing is empty without a request.
            cursor.Current = null;
            cursor.Pending.Clear();
        }

        var done = cursor.Current is null && cursor.Pending.Count == 0;
        return new DocumentListPage
        {
            Items = items,
            ContinuationToken = done ? null : cursor.Protect(connectionId, request.Fingerprint)
        };
    }

    /// <summary>The cursor a first page starts from: the listed folder, nothing read yet.</summary>
    internal static CloudListingCursor Start(string folderId, CloudListingRequest request, string? context = null) => new()
    {
        Current = new CloudListingFolder(folderId, request.FolderPath, depth: 0),
        Context = context
    };

    private static DocumentListEntry Entry(CloudListingChild child, string? folder, bool templates) => new()
    {
        Source = child.Id,
        Name = child.Name,
        Folder = folder,
        ContentType = OfficeContentTypes.ForName(child.Name),
        Length = child.Length,
        LastModified = child.LastModified,
        IsTemplate = templates || IsTemplateName(child.Name)
    };

    private static string Join(string? parent, string name) =>
        string.IsNullOrEmpty(parent) ? name : parent + "/" + name;
}
