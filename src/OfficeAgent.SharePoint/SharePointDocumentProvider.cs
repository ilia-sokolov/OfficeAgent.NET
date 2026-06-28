using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using OfficeAgent.Abstractions;
using OfficeAgent.Core.DocumentProviders;

namespace OfficeAgent.SharePoint;

/// <summary>Configures one SharePoint provider connection.</summary>
public sealed class SharePointDocumentProviderOptions
{
    /// <summary>Gets or sets the provider connection identifier.</summary>
    public string ConnectionId { get; set; } = string.Empty;

    /// <summary>Gets or sets the Graph endpoint base. Override for sovereign clouds.</summary>
    public string GraphBaseUrl { get; set; } = "https://graph.microsoft.com/v1.0";

    /// <summary>Gets or sets the maximum accepted document size in bytes.</summary>
    public long MaximumBytes { get; set; } = 100 * 1024 * 1024;

    /// <summary>Gets or sets the allowed document extensions.</summary>
    public IReadOnlyCollection<string> AllowedExtensions { get; set; } = new[] { ".docx" };
}

/// <summary>
/// SharePoint document provider over the Microsoft Graph drive-item REST endpoints:
/// a registry of references to drive items, addressed by an opaque item id assigned
/// at registration time. A document is registered by a SharePoint/OneDrive URL or by
/// a <c>driveId/itemId</c> pair, so the connection is not pinned to a single drive -
/// any drive the configured identity can reach is registrable. The provider stores
/// only the registration id → (drive id, item id) mapping (via
/// <see cref="ISharePointRegistrationStore"/>) and routes open/save back to Graph.
/// The agent only sees opaque ids - never site URLs, drive ids, or tokens. Saves use
/// the drive item's ETag for optimistic concurrency, and removing a registration
/// never deletes SharePoint content.
/// </summary>
public sealed class SharePointDocumentProvider : IDocumentProvider
{
    /// <summary>Gets the provider discriminator used in document references.</summary>
    public const string ProviderName = "sharepoint";

    private readonly SharePointDocumentProviderOptions _options;
    private readonly HttpClient _http;
    private readonly IAccessTokenProvider _tokens;
    private readonly ISharePointRegistrationStore _store;
    private readonly string _graphBaseUrl;
    private readonly HashSet<string> _allowedExtensions;

    /// <summary>Initializes a provider over one tenant's Graph drives.</summary>
    public SharePointDocumentProvider(
        SharePointDocumentProviderOptions options,
        HttpClient httpClient,
        IAccessTokenProvider tokenProvider,
        ISharePointRegistrationStore? registrationStore = null)
    {
        _options = options ?? throw new ArgumentNullException(nameof(options));
        _http = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
        _tokens = tokenProvider ?? throw new ArgumentNullException(nameof(tokenProvider));
        _store = registrationStore ?? new InMemoryRegistrationStore();

        if (string.IsNullOrWhiteSpace(options.ConnectionId))
            throw new ArgumentException("A SharePoint provider connection id is required.", nameof(options));
        if (string.IsNullOrWhiteSpace(options.GraphBaseUrl))
            throw new ArgumentException("A Graph base URL is required.", nameof(options));
        if (options.MaximumBytes <= 0)
            throw new ArgumentOutOfRangeException(nameof(options), "MaximumBytes must be positive.");
        if (options.AllowedExtensions is null || options.AllowedExtensions.Count == 0)
            throw new ArgumentException("At least one allowed document extension is required.", nameof(options));

        _allowedExtensions = new HashSet<string>(
            options.AllowedExtensions.Select(NormalizeExtension),
            StringComparer.OrdinalIgnoreCase);
        _graphBaseUrl = options.GraphBaseUrl.TrimEnd('/');
    }

    /// <inheritdoc />
    public string Provider => ProviderName;

    /// <inheritdoc />
    public string ConnectionId => _options.ConnectionId;

    /// <inheritdoc />
    public async Task<DocumentReference> RegisterAsync(
        string source,
        CancellationToken cancellationToken = default)
    {
        var item = await ResolveSourceAsync(source, cancellationToken).ConfigureAwait(false);

        ValidateExtension(item.Name, itemId: null);
        if (item.Size > _options.MaximumBytes)
            throw Error(ProviderErrorCode.ContentTooLarge,
                $"Document '{item.Name}' exceeds the configured maximum of {_options.MaximumBytes} bytes.");
        if (string.IsNullOrEmpty(item.DriveId))
            throw Error(ProviderErrorCode.IO,
                "Graph did not report which drive the item belongs to; the document cannot be registered.");

        var id = await _store.AddAsync(new SharePointItemRef(item.DriveId, item.Id), cancellationToken).ConfigureAwait(false);
        return CreateReference(id, item);
    }

    /// <summary>
    /// Resolves a registration source - a SharePoint/OneDrive URL or a
    /// <c>driveId/itemId</c> pair - to the Graph drive item it names.
    /// </summary>
    private async Task<GraphItem> ResolveSourceAsync(string source, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(source))
            throw Error(ProviderErrorCode.InvalidArgument,
                "A SharePoint registration requires a source: a document URL " +
                "(https://contoso.sharepoint.com/...) or a 'driveId/itemId' pair.");

        var trimmed = source.Trim();
        if (trimmed.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
            trimmed.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
        {
            var shareId = EncodeShareUrl(trimmed);
            return await GetItemAsync(
                $"{_graphBaseUrl}/shares/{shareId}/driveItem?$select=id,name,eTag,size,file,parentReference",
                cancellationToken).ConfigureAwait(false);
        }

        var (driveId, itemId) = ParseDriveAndItemId(trimmed);
        return await GetItemAsync(ItemUrl(driveId, itemId), cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Parses a <c>driveId/itemId</c> source. Accepts the bare pair
    /// (<c>b!abc.../01XYZ</c>) and the Graph resource form
    /// (<c>drives/{driveId}/items/{itemId}</c>).
    /// </summary>
    private (string DriveId, string ItemId) ParseDriveAndItemId(string source)
    {
        const string itemsMarker = "/items/";
        string driveId, itemId;

        var marker = source.IndexOf(itemsMarker, StringComparison.OrdinalIgnoreCase);
        if (marker > 0)
        {
            var drivePart = source.Substring(0, marker);
            const string drivesPrefix = "drives/";
            if (drivePart.StartsWith(drivesPrefix, StringComparison.OrdinalIgnoreCase))
                drivePart = drivePart.Substring(drivesPrefix.Length);
            driveId = drivePart;
            itemId = source.Substring(marker + itemsMarker.Length);
        }
        else
        {
            var slash = source.LastIndexOf('/');
            if (slash <= 0 || slash >= source.Length - 1)
                throw Error(ProviderErrorCode.InvalidArgument,
                    "Register a SharePoint document by a document URL or a 'driveId/itemId' pair " +
                    "(for example 'b!9a3f.../01ABCDEF'), not by a bare path or name.");
            driveId = source.Substring(0, slash);
            itemId = source.Substring(slash + 1);
        }

        driveId = driveId.Trim().Trim('/');
        itemId = itemId.Trim().Trim('/');
        if (driveId.Length == 0 || itemId.Length == 0 || itemId.Contains('/'))
            throw Error(ProviderErrorCode.InvalidArgument,
                "A 'driveId/itemId' source needs both a non-empty drive id and a single item id.");
        return (driveId, itemId);
    }

    /// <summary>Encodes a sharing or web URL into a Graph <c>shares</c> id (the <c>u!</c> form).</summary>
    private static string EncodeShareUrl(string url)
    {
        var base64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(url));
        return "u!" + base64.TrimEnd('=').Replace('/', '_').Replace('+', '-');
    }

    /// <inheritdoc />
    public async Task<DocumentContent> OpenReadAsync(
        DocumentReference reference,
        CancellationToken cancellationToken = default)
    {
        ValidateReference(reference);
        var registered = await ResolveRegistrationAsync(reference.ItemId, cancellationToken).ConfigureAwait(false);
        var item = await GetItemAsync(ItemUrl(registered.DriveId, registered.ItemId), cancellationToken, reference.ItemId).ConfigureAwait(false);

        if (item.Size > _options.MaximumBytes)
            throw Error(ProviderErrorCode.ContentTooLarge,
                $"Document exceeds the configured maximum of {_options.MaximumBytes} bytes.", reference.ItemId);
        if (!string.IsNullOrEmpty(reference.Version) &&
            !string.Equals(reference.Version, item.ETag, StringComparison.Ordinal))
            throw new DocumentVersionConflictException(reference.Version!, item.ETag, ProviderName, ConnectionId, reference.ItemId);

        var bytes = await DownloadAsync(item, reference.ItemId, cancellationToken).ConfigureAwait(false);
        return new DocumentContent(CreateReference(reference.ItemId, item), new MemoryStream(bytes, writable: false));
    }

    /// <inheritdoc />
    public async Task<DocumentReference> SaveAsync(
        DocumentReference source,
        Stream content,
        SaveDocumentOptions options,
        CancellationToken cancellationToken = default)
    {
        ValidateReference(source);
        if (content is null) throw new ArgumentNullException(nameof(content));
        if (options is null) throw new ArgumentNullException(nameof(options));
        if (!string.IsNullOrEmpty(options.DestinationItemId))
            throw Error(ProviderErrorCode.InvalidArgument,
                "SharePoint item ids are opaque and provider-assigned; a destination id cannot be chosen by the caller. Use Mode=NewVersion/NewDocument (optionally with NewName) to mint a new document, or Mode=Replace to overwrite the source.",
                source.ItemId);
        if (options.Mode == SaveMode.Replace && !string.IsNullOrWhiteSpace(options.NewName))
            throw Error(ProviderErrorCode.InvalidArgument,
                "Replace mode overwrites the source in place and cannot rename. Use Mode=NewVersion/NewDocument with NewName to mint a renamed document.",
                source.ItemId);

        var registered = await ResolveRegistrationAsync(source.ItemId, cancellationToken).ConfigureAwait(false);
        var driveId = registered.DriveId;
        var item = await GetItemAsync(ItemUrl(driveId, registered.ItemId), cancellationToken, source.ItemId).ConfigureAwait(false);

        var expectedVersion = options.ExpectedVersion ?? source.Version;
        if (!string.IsNullOrEmpty(expectedVersion) &&
            !string.Equals(expectedVersion, item.ETag, StringComparison.Ordinal))
            throw new DocumentVersionConflictException(expectedVersion!, item.ETag, ProviderName, ConnectionId, source.ItemId);

        var bytes = await ReadStreamAsync(content, source.ItemId, cancellationToken).ConfigureAwait(false);

        if (options.Mode == SaveMode.Replace)
        {
            var replaced = await UploadAsync(
                $"{ItemUrl(driveId, registered.ItemId)}/content",
                bytes, ifMatch: item.ETag, itemId: source.ItemId, cancellationToken).ConfigureAwait(false);
            return CreateReference(source.ItemId, replaced);
        }

        var destinationName = string.IsNullOrWhiteSpace(options.NewName)
            ? await NextVersionedNameAsync(driveId, item, cancellationToken).ConfigureAwait(false)
            : ValidateName(options.NewName!);
        var uploadUrl =
            $"{ItemUrl(driveId, item.ParentId)}:/{Uri.EscapeDataString(destinationName)}:/content?@microsoft.graph.conflictBehavior=fail";
        var saved = await UploadAsync(uploadUrl, bytes, ifMatch: null, itemId: source.ItemId, cancellationToken).ConfigureAwait(false);

        var savedDriveId = string.IsNullOrEmpty(saved.DriveId) ? driveId : saved.DriveId;
        var newId = await _store.AddAsync(new SharePointItemRef(savedDriveId, saved.Id), cancellationToken).ConfigureAwait(false);
        return CreateReference(newId, saved);
    }

    /// <inheritdoc />
    public async Task RemoveAsync(
        DocumentReference reference,
        CancellationToken cancellationToken = default)
    {
        ValidateReference(reference);
        if (!await _store.RemoveAsync(reference.ItemId, cancellationToken).ConfigureAwait(false))
            throw Error(ProviderErrorCode.NotFound,
                $"The provider document '{reference.ItemId}' is not registered.", reference.ItemId);
    }

    // ── Graph plumbing ────────────────────────────────────────────────────────

    private string ItemUrl(string driveId, string driveItemId) =>
        $"{_graphBaseUrl}/drives/{Uri.EscapeDataString(driveId)}/items/{Uri.EscapeDataString(driveItemId)}";

    private async Task<GraphItem> GetItemAsync(string url, CancellationToken cancellationToken, string? itemId = null)
    {
        using var response = await SendAsync(HttpMethod.Get, url, content: null, ifMatch: null, cancellationToken).ConfigureAwait(false);
        await ThrowOnFailureAsync(response, itemId).ConfigureAwait(false);
        return ParseItem(await response.Content.ReadAsStringAsync().ConfigureAwait(false), itemId);
    }

    private async Task<GraphItem> UploadAsync(string url, byte[] bytes, string? ifMatch, string? itemId, CancellationToken cancellationToken)
    {
        using var body = new ByteArrayContent(bytes);
        body.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
        using var response = await SendAsync(HttpMethod.Put, url, body, ifMatch, cancellationToken).ConfigureAwait(false);
        if (response.StatusCode == HttpStatusCode.Conflict)
            throw Error(ProviderErrorCode.IO,
                "A file already exists at the destination path; refusing to overwrite without Replace mode.", itemId);
        await ThrowOnFailureAsync(response, itemId).ConfigureAwait(false);
        return ParseItem(await response.Content.ReadAsStringAsync().ConfigureAwait(false), itemId);
    }

    private async Task<HttpResponseMessage> SendAsync(
        HttpMethod method, string url, HttpContent? content, string? ifMatch, CancellationToken cancellationToken)
    {
        var token = await _tokens.GetAccessTokenAsync(cancellationToken).ConfigureAwait(false);
        using var request = new HttpRequestMessage(method, url);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        if (ifMatch is not null)
            request.Headers.TryAddWithoutValidation("If-Match", ifMatch);
        if (content is not null)
            request.Content = content;
        return await _http.SendAsync(request, HttpCompletionOption.ResponseContentRead, cancellationToken).ConfigureAwait(false);
    }

    private async Task<byte[]> DownloadAsync(GraphItem item, string itemId, CancellationToken cancellationToken)
    {
        // The downloadUrl is pre-authenticated and short-lived; it is fetched without
        // the bearer header so credentials never travel to the storage endpoint.
        if (string.IsNullOrEmpty(item.DownloadUrl))
            throw Error(ProviderErrorCode.IO, "Graph did not return a download URL for the item.", itemId);

        using var response = await _http.GetAsync(item.DownloadUrl, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(false);
        if (!response.IsSuccessStatusCode)
            throw Error(ProviderErrorCode.IO,
                $"Content download failed with status {(int)response.StatusCode}.", itemId);

        var stream = await response.Content.ReadAsStreamAsync().ConfigureAwait(false);
        using (stream)
        {
            return await ReadBoundedAsync(stream, itemId, cancellationToken).ConfigureAwait(false);
        }
    }

    private async Task ThrowOnFailureAsync(HttpResponseMessage response, string? itemId)
    {
        if (response.IsSuccessStatusCode) return;

        var detail = await ReadGraphErrorAsync(response).ConfigureAwait(false);
        throw response.StatusCode switch
        {
            HttpStatusCode.NotFound => Error(ProviderErrorCode.NotFound,
                $"The referenced SharePoint item no longer exists. {detail}".TrimEnd(), itemId),
            HttpStatusCode.Unauthorized or HttpStatusCode.Forbidden => Error(ProviderErrorCode.AccessDenied,
                $"Graph refused access to the item. {detail}".TrimEnd(), itemId),
            HttpStatusCode.PreconditionFailed => new DocumentVersionConflictException(
                "(If-Match)", "(changed)", ProviderName, ConnectionId, itemId),
            HttpStatusCode.RequestEntityTooLarge => Error(ProviderErrorCode.ContentTooLarge,
                $"Graph rejected the content as too large. {detail}".TrimEnd(), itemId),
            _ => Error(ProviderErrorCode.IO,
                $"Graph request failed with status {(int)response.StatusCode}. {detail}".TrimEnd(), itemId)
        };
    }

    private static async Task<string> ReadGraphErrorAsync(HttpResponseMessage response)
    {
        try
        {
            var payload = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            if (string.IsNullOrWhiteSpace(payload)) return string.Empty;
            using var json = JsonDocument.Parse(payload);
            return json.RootElement.TryGetProperty("error", out var error) &&
                   error.TryGetProperty("message", out var message)
                ? message.GetString() ?? string.Empty
                : string.Empty;
        }
        catch (JsonException)
        {
            return string.Empty;
        }
    }

    private GraphItem ParseItem(string payload, string? itemId)
    {
        try
        {
            using var json = JsonDocument.Parse(payload);
            var root = json.RootElement;
            return new GraphItem
            {
                Id = root.GetProperty("id").GetString() ?? string.Empty,
                Name = root.TryGetProperty("name", out var name) ? name.GetString() ?? string.Empty : string.Empty,
                ETag = root.TryGetProperty("eTag", out var etag) ? etag.GetString() ?? string.Empty : string.Empty,
                Size = root.TryGetProperty("size", out var size) ? size.GetInt64() : 0,
                ParentId = root.TryGetProperty("parentReference", out var parent) &&
                           parent.TryGetProperty("id", out var parentId)
                    ? parentId.GetString() ?? string.Empty
                    : string.Empty,
                DriveId = root.TryGetProperty("parentReference", out var driveParent) &&
                          driveParent.TryGetProperty("driveId", out var driveId)
                    ? driveId.GetString() ?? string.Empty
                    : string.Empty,
                MimeType = root.TryGetProperty("file", out var file) &&
                           file.TryGetProperty("mimeType", out var mime)
                    ? mime.GetString()
                    : null,
                DownloadUrl = root.TryGetProperty("@microsoft.graph.downloadUrl", out var download)
                    ? download.GetString()
                    : null
            };
        }
        catch (Exception ex) when (ex is JsonException or KeyNotFoundException or InvalidOperationException)
        {
            throw Error(ProviderErrorCode.IO, $"Graph returned an unreadable drive item: {ex.Message}", itemId, ex);
        }
    }

    // ── Registration & validation ─────────────────────────────────────────────

    private async Task<SharePointItemRef> ResolveRegistrationAsync(string registrationId, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(registrationId))
            throw Error(ProviderErrorCode.InvalidArgument, "An item id is required.");
        return await _store.ResolveAsync(registrationId, cancellationToken).ConfigureAwait(false)
            ?? throw Error(ProviderErrorCode.NotFound,
                $"The provider document '{registrationId}' is not registered.", registrationId);
    }

    private void ValidateExtension(string name, string? itemId)
    {
        var extension = NormalizeExtension(Path.GetExtension(name));
        if (!_allowedExtensions.Contains(extension))
            throw Error(ProviderErrorCode.ExtensionNotAllowed,
                $"Documents with extension '{extension}' are not allowed by this connection. Allowed: {string.Join(", ", _allowedExtensions)}.",
                itemId);
    }

    private string ValidateName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw Error(ProviderErrorCode.InvalidArgument, "A document name is required.");
        if (name.IndexOfAny(new[] { '/', '\\' }) >= 0 || name == "." || name == "..")
            throw Error(ProviderErrorCode.InvalidArgument,
                "A document name must be a bare file name without path segments.");
        ValidateExtension(name, itemId: null);
        return name;
    }

    private void ValidateReference(DocumentReference reference)
    {
        if (reference is null) throw new ArgumentNullException(nameof(reference));
        if (!string.Equals(reference.Provider, ProviderName, StringComparison.OrdinalIgnoreCase) ||
            !string.Equals(reference.ConnectionId, ConnectionId, StringComparison.Ordinal))
            throw new DocumentProviderException(
                ProviderErrorCode.InvalidArgument,
                $"The document reference (provider='{reference.Provider}', connectionId='{reference.ConnectionId}') does not belong to this SharePoint connection (provider='{ProviderName}', connectionId='{ConnectionId}').",
                ProviderName, ConnectionId, reference.ItemId);
    }

    /// <summary>
    /// Picks the next unused versioned sibling name in the item's folder, matching the
    /// filesystem provider's convention: <c>contract.docx</c> → <c>contract.v2.docx</c>, ….
    /// </summary>
    private async Task<string> NextVersionedNameAsync(string driveId, GraphItem source, CancellationToken cancellationToken)
    {
        var extension = Path.GetExtension(source.Name);
        var stem = Path.GetFileNameWithoutExtension(source.Name);
        var (baseStem, startVersion) = ParseVersionedStem(stem);

        for (var n = Math.Max(2, startVersion + 1); n < startVersion + 1000; n++)
        {
            var candidate = $"{baseStem}.v{n}{extension}";
            var probeUrl = $"{ItemUrl(driveId, source.ParentId)}:/{Uri.EscapeDataString(candidate)}";
            using var response = await SendAsync(HttpMethod.Get, probeUrl, content: null, ifMatch: null, cancellationToken).ConfigureAwait(false);
            if (response.StatusCode == HttpStatusCode.NotFound)
                return candidate;
            await ThrowOnFailureAsync(response, source.Id).ConfigureAwait(false);
        }
        throw Error(ProviderErrorCode.IO, "Exhausted version suffixes for the source file.", source.Id);
    }

    private static (string Stem, int Version) ParseVersionedStem(string stem)
    {
        var dot = stem.LastIndexOf('.');
        if (dot <= 0 || dot == stem.Length - 1) return (stem, 1);
        var tail = stem.Substring(dot + 1);
        if (tail.Length < 2 || tail[0] is not ('v' or 'V')) return (stem, 1);
        if (int.TryParse(tail.Substring(1), out var version) && version >= 1)
            return (stem.Substring(0, dot), version);
        return (stem, 1);
    }

    private static string NormalizeExtension(string extension) =>
        extension.StartsWith(".", StringComparison.Ordinal) ? extension : "." + extension;

    private DocumentReference CreateReference(string registrationId, GraphItem item) => new()
    {
        Provider = ProviderName,
        ConnectionId = ConnectionId,
        ItemId = registrationId,
        Version = item.ETag,
        Name = item.Name,
        ContentType = item.MimeType ?? ContentTypeFor(item.Name)
    };

    private static string ContentTypeFor(string name) =>
        string.Equals(Path.GetExtension(name), ".docx", StringComparison.OrdinalIgnoreCase)
            ? "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            : "application/octet-stream";

    private async Task<byte[]> ReadStreamAsync(Stream content, string? itemId, CancellationToken cancellationToken)
    {
        if (content.CanSeek) content.Position = 0;
        return await ReadBoundedAsync(content, itemId, cancellationToken).ConfigureAwait(false);
    }

    private async Task<byte[]> ReadBoundedAsync(Stream stream, string? itemId, CancellationToken cancellationToken)
    {
        using var copy = new MemoryStream();
        var buffer = new byte[81920];
        while (true)
        {
            var read = await stream.ReadAsync(buffer, 0, buffer.Length, cancellationToken).ConfigureAwait(false);
            if (read == 0) break;
            await copy.WriteAsync(buffer, 0, read, cancellationToken).ConfigureAwait(false);
            if (copy.Length > _options.MaximumBytes)
                throw Error(ProviderErrorCode.ContentTooLarge,
                    $"Document exceeds the configured maximum of {_options.MaximumBytes} bytes.", itemId);
        }
        return copy.ToArray();
    }

    private DocumentProviderException Error(
        ProviderErrorCode code,
        string message,
        string? itemId = null,
        Exception? innerException = null) =>
        new(code, message, ProviderName, ConnectionId, itemId, innerException);

    private sealed class GraphItem
    {
        public string Id { get; init; } = string.Empty;
        public string Name { get; init; } = string.Empty;
        public string ETag { get; init; } = string.Empty;
        public long Size { get; init; }
        public string ParentId { get; init; } = string.Empty;
        public string DriveId { get; init; } = string.Empty;
        public string? MimeType { get; init; }
        public string? DownloadUrl { get; init; }
    }
}
