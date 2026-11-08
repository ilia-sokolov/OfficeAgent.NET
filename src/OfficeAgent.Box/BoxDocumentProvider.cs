using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using OfficeAgent.Abstractions;
using OfficeAgent.Core.DocumentProviders;

namespace OfficeAgent.Box;

/// <summary>Configures one Box provider connection.</summary>
public sealed class BoxDocumentProviderOptions
{
    /// <summary>Gets or sets the provider connection identifier.</summary>
    public string ConnectionId { get; set; } = string.Empty;

    /// <summary>Gets or sets the Box API base URL.</summary>
    public string ApiBaseUrl { get; set; } = "https://api.box.com/2.0";

    /// <summary>Gets or sets the Box upload API base URL.</summary>
    public string UploadBaseUrl { get; set; } = "https://upload.box.com/api/2.0";

    /// <summary>
    /// Gets or sets the maximum accepted document size in bytes. The default is 50 MB, the
    /// size Box recommends as the ceiling for its direct upload endpoints; the provider does
    /// not implement chunked uploads.
    /// </summary>
    public long MaximumBytes { get; set; } = 50 * 1024 * 1024;

    /// <summary>Gets or sets the allowed document extensions.</summary>
    public IReadOnlyCollection<string> AllowedExtensions { get; set; } = new[] { ".docx" };

    /// <summary>
    /// Gets or sets the Box folder id that receives new documents. Empty means the
    /// connection cannot create documents; registering existing ones still works.
    /// </summary>
    public string CreationFolderId { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the change mode used when an operation does not state one. The
    /// default is <see cref="ChangeMode.Tracked"/>, which suits documents a person
    /// reviews. Set <see cref="ChangeMode.Direct"/> for generated or machine-owned files,
    /// or for decks - PresentationML refuses Tracked.
    /// </summary>
    public ChangeMode DefaultChangeMode { get; set; } = ChangeMode.Tracked;
    /// <summary>
    /// Gets or sets how much of the matched text a tracked <c>changeText</c> that does not
    /// state a granularity marks. The default, <see cref="ChangeGranularity.Span"/>, marks the
    /// whole matched text; <see cref="ChangeGranularity.Minimal"/> marks only the words that
    /// differ.
    /// </summary>
    public ChangeGranularity DefaultChangeGranularity { get; set; } = ChangeGranularity.Span;

    /// <summary>
    /// Gets or sets whether a workbook previewed or committed through this connection has its
    /// formulas calculated. Takes effect only when a formula evaluator is registered
    /// (<c>AddExcelFormulaCalculation</c>). Defaults to <see langword="false"/>.
    /// </summary>
    public bool CalculateFormulas { get; set; }

    /// <summary>
    /// Gets or sets whether a plan whose calculation produces a new formula error or a circular
    /// reference is refused with <c>formula-error</c>. Applies only with
    /// <see cref="CalculateFormulas"/>. Defaults to <see langword="false"/>.
    /// </summary>
    public bool RejectFormulaErrors { get; set; }
}

/// <summary>
/// Box document provider over the Box REST API: a registry of references to Box files,
/// addressed by an opaque id assigned at registration time. A document is registered by a
/// Box file id or a Box file URL (<c>https://app.box.com/file/{id}</c>).
/// </summary>
/// <remarks>
/// Versions are the file's <c>etag</c>. <c>Replace</c> uploads a new file version with
/// <c>If-Match</c> on the opened etag, so Box itself refuses the write when the file changed
/// (412), and Box keeps the previous version in the file's history. New documents and
/// versioned siblings are uploaded as new files, which Box refuses with 409 when the name is
/// taken, so creation never overwrites. Removing a registration never deletes Box content.
/// </remarks>
public sealed class BoxDocumentProvider : IDocumentProvider, IDocumentCreatingProvider, IConnectionEditingDefaults, IConnectionGranularityDefaults, IConnectionCalculationDefaults
{
    /// <summary>Gets the provider discriminator used in document references.</summary>
    public const string ProviderName = "box";

    private const string FileFields = "id,type,name,size,etag,parent,item_status";

    private readonly BoxDocumentProviderOptions _options;
    private readonly HttpClient _http;
    private readonly IBoxAccessTokenProvider _tokens;
    private readonly IBoxRegistrationStore _store;
    private readonly string _apiBaseUrl;
    private readonly string _uploadBaseUrl;
    private readonly HashSet<string> _allowedExtensions;
    private readonly string? _creationFolderId;

    /// <summary>Initializes a provider over the Box content the token's identity can reach.</summary>
    public BoxDocumentProvider(
        BoxDocumentProviderOptions options,
        HttpClient httpClient,
        IBoxAccessTokenProvider tokenProvider,
        IBoxRegistrationStore? registrationStore = null)
    {
        _options = options ?? throw new ArgumentNullException(nameof(options));
        _http = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
        _tokens = tokenProvider ?? throw new ArgumentNullException(nameof(tokenProvider));
        _store = registrationStore ?? new InMemoryBoxRegistrationStore();

        if (string.IsNullOrWhiteSpace(options.ConnectionId))
            throw new ArgumentException("A Box provider connection id is required.", nameof(options));
        if (string.IsNullOrWhiteSpace(options.ApiBaseUrl) || string.IsNullOrWhiteSpace(options.UploadBaseUrl))
            throw new ArgumentException("The Box API and upload base URLs are required.", nameof(options));
        if (options.MaximumBytes <= 0)
            throw new ArgumentOutOfRangeException(nameof(options), "MaximumBytes must be positive.");
        if (options.AllowedExtensions is null || options.AllowedExtensions.Count == 0)
            throw new ArgumentException("At least one allowed document extension is required.", nameof(options));
        if (!string.IsNullOrWhiteSpace(options.CreationFolderId) && !IsBoxId(options.CreationFolderId.Trim()))
            throw new ArgumentException("CreationFolderId must be a numeric Box folder id.", nameof(options));

        _allowedExtensions = new HashSet<string>(
            options.AllowedExtensions.Select(NormalizeExtension),
            StringComparer.OrdinalIgnoreCase);
        _apiBaseUrl = options.ApiBaseUrl.TrimEnd('/');
        _uploadBaseUrl = options.UploadBaseUrl.TrimEnd('/');
        _creationFolderId = string.IsNullOrWhiteSpace(options.CreationFolderId)
            ? null
            : options.CreationFolderId.Trim();
    }

    /// <inheritdoc />
    public string Provider => ProviderName;

    /// <inheritdoc />
    public string ConnectionId => _options.ConnectionId;

    /// <inheritdoc />
    public ChangeMode DefaultChangeMode => _options.DefaultChangeMode;

    /// <inheritdoc />
    public ChangeGranularity DefaultChangeGranularity => _options.DefaultChangeGranularity;

    /// <inheritdoc />
    public bool CalculateFormulas => _options.CalculateFormulas;

    /// <inheritdoc />
    public bool RejectFormulaErrors => _options.RejectFormulaErrors;

    /// <inheritdoc />
    public async Task<DocumentReference> RegisterAsync(
        string source,
        CancellationToken cancellationToken = default)
    {
        var fileId = ParseSource(source);
        var file = await GetFileAsync(fileId, itemId: null, cancellationToken).ConfigureAwait(false);
        EnsureDocument(file, itemId: null);

        var id = await _store.AddAsync(new BoxFileRef(file.Id), cancellationToken).ConfigureAwait(false);
        return CreateReference(id, file);
    }

    /// <inheritdoc />
    public async Task<DocumentReference> CreateAsync(
        string name,
        Stream content,
        CancellationToken cancellationToken = default)
    {
        if (content is null) throw new ArgumentNullException(nameof(content));
        if (_creationFolderId is null)
            throw Error(
                ProviderErrorCode.ConfigurationError,
                "This Box connection has no creation folder. Configure CreationFolderId, " +
                "or register an existing document instead.");

        var fileName = ValidateCreateName(name);
        var bytes = await ReadStreamAsync(content, itemId: null, cancellationToken).ConfigureAwait(false);
        BoxFile created;
        try
        {
            created = await UploadNewAsync(_creationFolderId, fileName, bytes, itemId: null, cancellationToken).ConfigureAwait(false);
        }
        catch (DocumentProviderException ex) when (ex.Code != ProviderErrorCode.AlreadyExists)
        {
            throw Error(ex.Code, SafeCreateFailureMessage(ex.Code, fileName), itemId: null, ex);
        }
        catch (Exception ex) when (ex is not OperationCanceledException and not DocumentProviderException)
        {
            // Transport and token-provider failures can include the request URL. Keep those
            // details only on the inner exception available to the host.
            throw Error(
                ProviderErrorCode.IO,
                $"Box could not create document '{fileName}' in the configured folder.",
                itemId: null,
                ex);
        }

        string registrationId;
        try
        {
            // Box confirmed the upload: finish registering even if the caller cancels now.
            registrationId = await _store.AddAsync(new BoxFileRef(created.Id), CancellationToken.None).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            throw RegistrationFailed(
                $"Document '{fileName}' was created in Box, but its registration could not be " +
                "persisted. The file is stored but unregistered. Do not retry creation; register it, " +
                "or report it to the host/operator for recovery.",
                itemId: null, fileName, bytes, ex);
        }
        return CreateReference(registrationId, created);
    }

    /// <inheritdoc />
    public async Task<DocumentContent> OpenReadAsync(
        DocumentReference reference,
        CancellationToken cancellationToken = default)
    {
        ValidateReference(reference);
        var registered = await ResolveRegistrationAsync(reference.ItemId, cancellationToken).ConfigureAwait(false);
        var file = await GetFileAsync(registered.FileId, reference.ItemId, cancellationToken).ConfigureAwait(false);
        EnsureDocument(file, reference.ItemId);

        if (!string.IsNullOrEmpty(reference.Version) &&
            !string.Equals(reference.Version, file.ETag, StringComparison.Ordinal))
            throw new DocumentVersionConflictException(reference.Version!, file.ETag, ProviderName, ConnectionId, reference.ItemId);

        var bytes = await DownloadAsync(file.Id, reference.ItemId, cancellationToken).ConfigureAwait(false);
        return new DocumentContent(CreateReference(reference.ItemId, file), new MemoryStream(bytes, writable: false));
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
                "Box document ids are opaque and provider-assigned; a destination id cannot be chosen by the caller. Use Mode=NewVersion/NewDocument (optionally with NewName) to mint a new document, or Mode=Replace to overwrite the source.",
                source.ItemId);
        if (options.Mode == SaveMode.Replace && !string.IsNullOrWhiteSpace(options.NewName))
            throw Error(ProviderErrorCode.InvalidArgument,
                "Replace mode overwrites the source in place and cannot rename. Use Mode=NewVersion/NewDocument with NewName to mint a renamed document.",
                source.ItemId);

        var registered = await ResolveRegistrationAsync(source.ItemId, cancellationToken).ConfigureAwait(false);
        var bytes = await ReadStreamAsync(content, source.ItemId, cancellationToken).ConfigureAwait(false);
        var file = await GetFileAsync(registered.FileId, source.ItemId, cancellationToken).ConfigureAwait(false);
        EnsureDocument(file, source.ItemId);

        var expectedVersion = options.ExpectedVersion ?? source.Version;
        if (!string.IsNullOrEmpty(expectedVersion) &&
            !string.Equals(expectedVersion, file.ETag, StringComparison.Ordinal))
            throw new DocumentVersionConflictException(expectedVersion!, file.ETag, ProviderName, ConnectionId, source.ItemId);

        if (options.Mode == SaveMode.Replace)
        {
            if (string.IsNullOrEmpty(file.ETag))
                throw Error(ProviderErrorCode.WriteRejected,
                    "Box did not report an etag for the file, so the replace cannot be made conditional; nothing was stored.",
                    source.ItemId);
            // If-Match makes Box refuse the version upload when the file changed since the
            // check above, so a write between the check and the upload is never overwritten.
            var replaced = await UploadVersionAsync(file, bytes, source.ItemId, cancellationToken).ConfigureAwait(false);
            return CreateReference(source.ItemId, replaced);
        }

        if (string.IsNullOrEmpty(file.ParentId))
            throw Error(ProviderErrorCode.AccessDenied,
                "The configured Box identity cannot see the folder this file is in, so it cannot save a " +
                "sibling document; nothing was stored. Collaborate it on the folder, or use Mode=Replace.",
                source.ItemId);

        var destinationName = string.IsNullOrWhiteSpace(options.NewName)
            ? await NextVersionedNameAsync(file, bytes.LongLength, source.ItemId, cancellationToken).ConfigureAwait(false)
            : ValidateName(options.NewName!);
        var saved = await UploadNewAsync(file.ParentId, destinationName, bytes, source.ItemId, cancellationToken).ConfigureAwait(false);

        string newId;
        try
        {
            newId = await _store.AddAsync(new BoxFileRef(saved.Id), CancellationToken.None).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            throw RegistrationFailed(
                $"Document '{destinationName}' was saved to Box, but its registration could not be " +
                "persisted. Do not save it again; register the saved file to recover it.",
                source.ItemId, destinationName, bytes, ex);
        }
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

    // ── Source parsing ────────────────────────────────────────────────────────

    /// <summary>
    /// Resolves a registration source to a Box file id: a bare numeric id, or a Box file URL
    /// such as <c>https://app.box.com/file/123</c> or <c>https://acme.app.box.com/file/123</c>.
    /// </summary>
    private string ParseSource(string source)
    {
        if (string.IsNullOrWhiteSpace(source))
            throw Error(ProviderErrorCode.InvalidArgument,
                "A Box registration requires a source: a Box file id or a Box file URL (https://app.box.com/file/...).");

        var trimmed = source.Trim();
        if (!trimmed.StartsWith("http://", StringComparison.OrdinalIgnoreCase) &&
            !trimmed.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
        {
            if (!IsBoxId(trimmed))
                throw Error(ProviderErrorCode.InvalidArgument,
                    "Register a Box document by its numeric file id or its Box file URL, not by a name or path.");
            return trimmed;
        }

        if (!Uri.TryCreate(trimmed, UriKind.Absolute, out var uri) ||
            !(uri.Host.Equals("box.com", StringComparison.OrdinalIgnoreCase) ||
              uri.Host.EndsWith(".box.com", StringComparison.OrdinalIgnoreCase)))
            throw Error(ProviderErrorCode.InvalidArgument, "The URL is not a Box URL.");

        var segments = uri.AbsolutePath.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
        for (var i = 0; i < segments.Length - 1; i++)
        {
            if (segments[i] == "file" && IsBoxId(segments[i + 1]))
                return segments[i + 1];
        }
        if (segments.Length > 0 && segments[0] is "s" or "v")
            throw Error(ProviderErrorCode.InvalidArgument,
                "The URL is a Box shared link. Register the file by its id or its https://app.box.com/file/{id} URL.");
        if (segments.Contains("folder"))
            throw Error(ProviderErrorCode.InvalidArgument,
                "The URL names a Box folder. Register a file instead.");
        throw Error(ProviderErrorCode.InvalidArgument,
            "The URL does not name a Box file. Use https://app.box.com/file/{id} or the file id.");
    }

    private static bool IsBoxId(string value) =>
        value.Length is > 0 and <= 30 && value.All(c => c is >= '0' and <= '9');

    // ── Box plumbing ──────────────────────────────────────────────────────────

    private async Task<BoxFile> GetFileAsync(string fileId, string? itemId, CancellationToken cancellationToken)
    {
        var url = $"{_apiBaseUrl}/files/{Uri.EscapeDataString(fileId)}?fields={Uri.EscapeDataString(FileFields)}";
        using var response = await SendAsync(HttpMethod.Get, url, content: null, ifMatch: null, cancellationToken).ConfigureAwait(false);
        await ThrowOnFailureAsync(response, itemId).ConfigureAwait(false);
        return ParseFile(await response.Content.ReadAsStringAsync().ConfigureAwait(false), itemId);
    }

    private async Task<byte[]> DownloadAsync(string fileId, string itemId, CancellationToken cancellationToken)
    {
        var url = $"{_apiBaseUrl}/files/{Uri.EscapeDataString(fileId)}/content";
        var response = await SendAsync(HttpMethod.Get, url, content: null, ifMatch: null, cancellationToken,
            HttpCompletionOption.ResponseHeadersRead).ConfigureAwait(false);
        try
        {
            // Box answers with a redirect to a short-lived, pre-authenticated download URL. When
            // the HttpClient does not follow it, follow it here without the bearer token, so the
            // credential never travels to the download host.
            var status = (int)response.StatusCode;
            if (status is >= 300 and < 400 && response.Headers.Location is { } location)
            {
                var target = location.IsAbsoluteUri ? location : new Uri(new Uri(url), location);
                response.Dispose();
                response = await _http.GetAsync(target, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(false);
                if (!response.IsSuccessStatusCode)
                    throw Error(ProviderErrorCode.IO,
                        $"Box content download failed with status {(int)response.StatusCode}.", itemId);
            }
            else if (response.StatusCode == HttpStatusCode.Accepted)
            {
                throw Error(ProviderErrorCode.IO,
                    "Box has not finished processing this file yet; try again shortly.", itemId);
            }
            else
            {
                await ThrowOnFailureAsync(response, itemId).ConfigureAwait(false);
            }

            if (response.Content.Headers.ContentLength is { } length && length > _options.MaximumBytes)
                throw TooLarge(itemId);
            var stream = await response.Content.ReadAsStreamAsync().ConfigureAwait(false);
            using (stream)
            {
                return await ReadBoundedAsync(stream, itemId, cancellationToken).ConfigureAwait(false);
            }
        }
        finally
        {
            response.Dispose();
        }
    }

    private async Task<BoxFile> UploadVersionAsync(BoxFile file, byte[] bytes, string itemId, CancellationToken cancellationToken)
    {
        var url = $"{_uploadBaseUrl}/files/{Uri.EscapeDataString(file.Id)}/content?fields={Uri.EscapeDataString(FileFields)}";
        using var body = UploadBody(new Dictionary<string, object> { ["name"] = file.Name }, file.Name, bytes);
        using var response = await SendAsync(HttpMethod.Post, url, body, file.ETag, cancellationToken).ConfigureAwait(false);
        await ThrowOnFailureAsync(response, itemId, write: true).ConfigureAwait(false);
        return ParseUploaded(await response.Content.ReadAsStringAsync().ConfigureAwait(false), itemId);
    }

    private async Task<BoxFile> UploadNewAsync(
        string folderId, string name, byte[] bytes, string? itemId, CancellationToken cancellationToken)
    {
        var url = $"{_uploadBaseUrl}/files/content?fields={Uri.EscapeDataString(FileFields)}";
        var attributes = new Dictionary<string, object>
        {
            ["name"] = name,
            ["parent"] = new Dictionary<string, string> { ["id"] = folderId }
        };
        using var body = UploadBody(attributes, name, bytes);
        using var response = await SendAsync(HttpMethod.Post, url, body, ifMatch: null, cancellationToken).ConfigureAwait(false);
        if (response.StatusCode == HttpStatusCode.Conflict)
        {
            var code = await ReadErrorCodeAsync(response).ConfigureAwait(false);
            if (code is "" or "item_name_in_use" or "conflict" or "name_temporarily_reserved")
                throw AlreadyExists(name, itemId);
        }
        await ThrowOnFailureAsync(response, itemId, write: true).ConfigureAwait(false);
        return ParseUploaded(await response.Content.ReadAsStringAsync().ConfigureAwait(false), itemId);
    }

    /// <summary>
    /// The multipart body Box's upload endpoints take: the <c>attributes</c> JSON part, which
    /// must come first, then the <c>file</c> part. Box names the file from the attributes.
    /// </summary>
    private static MultipartFormDataContent UploadBody(Dictionary<string, object> attributes, string name, byte[] bytes)
    {
        var body = new MultipartFormDataContent("officeagent-" + Guid.NewGuid().ToString("N"));
        var attributesPart = new StringContent(JsonSerializer.Serialize(attributes), Encoding.UTF8);
        attributesPart.Headers.ContentType = null;
        body.Add(attributesPart, "attributes");
        var filePart = new ByteArrayContent(bytes);
        filePart.Headers.ContentType = new MediaTypeHeaderValue(OfficeContentTypes.ForName(name));
        body.Add(filePart, "file", AsciiFileName(name));
        return body;
    }

    private static string AsciiFileName(string name)
    {
        var builder = new StringBuilder(name.Length);
        foreach (var c in name)
            builder.Append(c is >= ' ' and < (char)127 and not '"' and not '\\' ? c : '_');
        return builder.ToString();
    }

    /// <summary>
    /// Whether Box would accept a new file with this name in the folder, asked with Box's
    /// upload preflight check (<c>OPTIONS /files/content</c>), which uploads nothing.
    /// </summary>
    private async Task<bool> NameIsFreeAsync(
        string folderId, string name, long size, string? itemId, CancellationToken cancellationToken)
    {
        var payload = JsonSerializer.Serialize(new Dictionary<string, object>
        {
            ["name"] = name,
            ["size"] = size,
            ["parent"] = new Dictionary<string, string> { ["id"] = folderId }
        });
        using var body = new StringContent(payload, Encoding.UTF8, "application/json");
        using var response = await SendAsync(HttpMethod.Options, $"{_apiBaseUrl}/files/content", body, ifMatch: null, cancellationToken).ConfigureAwait(false);
        if (response.StatusCode == HttpStatusCode.Conflict) return false;
        await ThrowOnFailureAsync(response, itemId).ConfigureAwait(false);
        return true;
    }

    private async Task<HttpResponseMessage> SendAsync(
        HttpMethod method,
        string url,
        HttpContent? content,
        string? ifMatch,
        CancellationToken cancellationToken,
        HttpCompletionOption completion = HttpCompletionOption.ResponseContentRead)
    {
        var token = await _tokens.GetAccessTokenAsync(cancellationToken).ConfigureAwait(false);
        using var request = new HttpRequestMessage(method, url);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        if (ifMatch is not null)
            request.Headers.TryAddWithoutValidation("If-Match", ifMatch);
        if (content is not null)
            request.Content = content;
        return await _http.SendAsync(request, completion, cancellationToken).ConfigureAwait(false);
    }

    private async Task ThrowOnFailureAsync(HttpResponseMessage response, string? itemId, bool write = false)
    {
        if (response.IsSuccessStatusCode) return;

        var code = await ReadErrorCodeAsync(response).ConfigureAwait(false);
        var suffix = string.IsNullOrEmpty(code) ? string.Empty : $" Box code: {code}.";
        var status = (int)response.StatusCode;
        throw response.StatusCode switch
        {
            // Retrying cannot help a locked file, so the message says what can.
            _ when write && code == "access_denied_item_locked" => Error(ProviderErrorCode.WriteRejected,
                "The document is locked in Box; nothing was stored. Ask whoever locked it to unlock it, " +
                $"or save as a new document instead.{suffix}", itemId),
            _ when code == "forbidden_by_policy" => Error(ProviderErrorCode.AccessDenied,
                $"A Box Shield policy blocks this action on the file. Ask the Box administrator about the policy.{suffix}", itemId),
            _ when write && code is "storage_limit_exceeded" => Error(ProviderErrorCode.WriteRejected,
                $"The Box account's storage limit is reached; nothing was stored.{suffix}", itemId),
            _ when code is "file_size_limit_exceeded" => Error(ProviderErrorCode.ContentTooLarge,
                $"Box rejected the content as too large for the folder owner's limit.{suffix}", itemId),
            HttpStatusCode.NotFound => Error(ProviderErrorCode.NotFound,
                $"The referenced Box item does not exist or is not visible to the configured identity.{suffix}", itemId),
            HttpStatusCode.Unauthorized or HttpStatusCode.Forbidden => Error(ProviderErrorCode.AccessDenied,
                $"Box refused access to the item.{suffix}", itemId),
            HttpStatusCode.PreconditionFailed => new DocumentVersionConflictException(
                "(If-Match)", "(changed)", ProviderName, ConnectionId, itemId),
            HttpStatusCode.RequestEntityTooLarge => Error(ProviderErrorCode.ContentTooLarge,
                $"Box rejected the content as too large.{suffix}", itemId),
            // For an upload, any other refusal (400, 409, 429 throttled) is Box declining the
            // request: nothing was stored. A server error is not a refusal - Box may have
            // processed the upload - so its outcome is unknown.
            _ when write && status is >= 400 and < 500 => Error(ProviderErrorCode.WriteRejected,
                $"Box refused the upload with status {status}; nothing was stored.{suffix}", itemId),
            _ when write => Error(ProviderErrorCode.OutcomeUnknown,
                $"Box failed the upload with status {status}; it may have been stored.{suffix}", itemId),
            _ => Error(ProviderErrorCode.IO,
                $"Box request failed with status {status}.{suffix}", itemId)
        };
    }

    /// <summary>
    /// Reads the machine <c>code</c> from a Box error body. The human message is not copied:
    /// it can quote names and ids.
    /// </summary>
    private static async Task<string> ReadErrorCodeAsync(HttpResponseMessage response)
    {
        try
        {
            var payload = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            if (string.IsNullOrWhiteSpace(payload)) return string.Empty;
            using var json = JsonDocument.Parse(payload);
            return json.RootElement.ValueKind == JsonValueKind.Object &&
                   json.RootElement.TryGetProperty("code", out var code) &&
                   code.ValueKind == JsonValueKind.String
                ? code.GetString() ?? string.Empty
                : string.Empty;
        }
        catch (JsonException)
        {
            return string.Empty;
        }
    }

    private BoxFile ParseUploaded(string payload, string? itemId)
    {
        // The upload endpoints answer with a list: { "total_count": 1, "entries": [ file ] }.
        try
        {
            using var json = JsonDocument.Parse(payload);
            if (json.RootElement.TryGetProperty("entries", out var entries) &&
                entries.ValueKind == JsonValueKind.Array &&
                entries.GetArrayLength() > 0)
                return ParseFile(entries[0], itemId);
        }
        catch (JsonException ex)
        {
            throw Error(ProviderErrorCode.IO, "Box returned an unreadable upload result.", itemId, ex);
        }
        throw Error(ProviderErrorCode.IO, "Box returned an upload result without the uploaded file.", itemId);
    }

    private BoxFile ParseFile(string payload, string? itemId)
    {
        try
        {
            using var json = JsonDocument.Parse(payload);
            return ParseFile(json.RootElement, itemId);
        }
        catch (JsonException ex)
        {
            throw Error(ProviderErrorCode.IO, "Box returned an unreadable file object.", itemId, ex);
        }
    }

    private BoxFile ParseFile(JsonElement root, string? itemId)
    {
        try
        {
            return new BoxFile
            {
                Id = root.GetProperty("id").GetString() ?? string.Empty,
                Type = StringOf(root, "type"),
                Name = StringOf(root, "name"),
                Size = root.TryGetProperty("size", out var size) && size.ValueKind == JsonValueKind.Number
                    ? size.GetInt64()
                    : 0,
                ETag = StringOf(root, "etag"),
                ParentId = root.TryGetProperty("parent", out var parent) && parent.ValueKind == JsonValueKind.Object
                    ? StringOf(parent, "id")
                    : string.Empty,
                ItemStatus = StringOf(root, "item_status")
            };
        }
        catch (Exception ex) when (ex is KeyNotFoundException or InvalidOperationException or FormatException)
        {
            throw Error(ProviderErrorCode.IO, "Box returned an unreadable file object.", itemId, ex);
        }
    }

    private static string StringOf(JsonElement element, string name) =>
        element.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String
            ? value.GetString() ?? string.Empty
            : string.Empty;

    // ── Registration & validation ─────────────────────────────────────────────

    private async Task<BoxFileRef> ResolveRegistrationAsync(string registrationId, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(registrationId))
            throw Error(ProviderErrorCode.InvalidArgument, "An item id is required.");
        return await _store.ResolveAsync(registrationId, cancellationToken).ConfigureAwait(false)
            ?? throw Error(ProviderErrorCode.NotFound,
                $"The provider document '{registrationId}' is not registered.", registrationId);
    }

    /// <summary>
    /// Refuses what this connection does not accept, on metadata alone and before any content
    /// is downloaded: items that are not files, trashed files, disallowed extensions, and files
    /// over the size cap.
    /// </summary>
    private void EnsureDocument(BoxFile file, string? itemId)
    {
        if (!string.IsNullOrEmpty(file.Type) && !string.Equals(file.Type, "file", StringComparison.Ordinal))
            throw Error(ProviderErrorCode.InvalidArgument,
                $"The Box item is a {file.Type}, not a file.", itemId);
        if (!string.IsNullOrEmpty(file.ItemStatus) && !string.Equals(file.ItemStatus, "active", StringComparison.Ordinal))
            throw Error(ProviderErrorCode.NotFound,
                $"Box file '{file.Name}' is {file.ItemStatus}.", itemId);
        ValidateExtension(file.Name, itemId);
        if (file.Size > _options.MaximumBytes)
            throw Error(ProviderErrorCode.ContentTooLarge,
                $"Document '{file.Name}' exceeds the configured maximum of {_options.MaximumBytes} bytes.", itemId);
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

    private string ValidateCreateName(string name)
    {
        if (!string.IsNullOrWhiteSpace(name) && PortableDocumentName.ContainsInvalidCharacter(name))
            throw Error(ProviderErrorCode.InvalidArgument,
                "A document name must be a bare file name without path separators, " +
                "control characters, or other invalid filename characters.");
        if (!string.IsNullOrWhiteSpace(name) && PortableDocumentName.IsWindowsDeviceName(name))
            throw Error(ProviderErrorCode.InvalidArgument,
                "A document name must not use a reserved Windows device name.");
        if (!string.IsNullOrWhiteSpace(name) &&
            (!string.Equals(name, name.Trim(), StringComparison.Ordinal) ||
             name.EndsWith(".", StringComparison.Ordinal)))
            throw Error(ProviderErrorCode.InvalidArgument,
                "A document name must not begin or end with whitespace or end with a dot.");
        return ValidateName(name);
    }

    private static string SafeCreateFailureMessage(ProviderErrorCode code, string fileName) => code switch
    {
        ProviderErrorCode.NotFound =>
            $"The configured Box creation folder for '{fileName}' was not found.",
        ProviderErrorCode.AccessDenied =>
            $"The configured Box identity cannot create '{fileName}' in the creation folder.",
        ProviderErrorCode.ContentTooLarge =>
            $"Box rejected '{fileName}' because its content is too large.",
        ProviderErrorCode.WriteRejected =>
            $"Box refused to create '{fileName}'; nothing was stored.",
        ProviderErrorCode.OutcomeUnknown =>
            $"Box failed while creating '{fileName}'; it may have been stored.",
        _ => $"Box could not create document '{fileName}' in the configured folder."
    };

    private void ValidateReference(DocumentReference reference)
    {
        if (reference is null) throw new ArgumentNullException(nameof(reference));
        if (!string.Equals(reference.Provider, ProviderName, StringComparison.OrdinalIgnoreCase) ||
            !string.Equals(reference.ConnectionId, ConnectionId, StringComparison.Ordinal))
            throw new DocumentProviderException(
                ProviderErrorCode.InvalidArgument,
                $"The document reference (provider='{reference.Provider}', connectionId='{reference.ConnectionId}') does not belong to this Box connection (provider='{ProviderName}', connectionId='{ConnectionId}').",
                ProviderName, ConnectionId, reference.ItemId);
    }

    /// <summary>
    /// Picks the next versioned sibling name Box would accept in the file's folder, matching
    /// the filesystem provider's convention: <c>contract.docx</c> → <c>contract.v2.docx</c>, ….
    /// </summary>
    private async Task<string> NextVersionedNameAsync(BoxFile source, long size, string itemId, CancellationToken cancellationToken)
    {
        var extension = Path.GetExtension(source.Name);
        var stem = Path.GetFileNameWithoutExtension(source.Name);
        var (baseStem, startVersion) = ParseVersionedStem(stem);

        for (var n = Math.Max(2, startVersion + 1); n < startVersion + 1000; n++)
        {
            var candidate = $"{baseStem}.v{n}{extension}";
            if (await NameIsFreeAsync(source.ParentId, candidate, size, itemId, cancellationToken).ConfigureAwait(false))
                return candidate;
        }
        throw Error(ProviderErrorCode.IO, "Exhausted version suffixes for the source file.", itemId);
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

    private DocumentReference CreateReference(string registrationId, BoxFile file) => new()
    {
        Provider = ProviderName,
        ConnectionId = ConnectionId,
        ItemId = registrationId,
        Version = file.ETag,
        Name = file.Name,
        ContentType = OfficeContentTypes.ForName(file.Name)
    };

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
                throw TooLarge(itemId);
        }
        return copy.ToArray();
    }

    private DocumentProviderException TooLarge(string? itemId) => Error(ProviderErrorCode.ContentTooLarge,
        $"Document exceeds the configured maximum of {_options.MaximumBytes} bytes.", itemId);

    private DocumentProviderException AlreadyExists(string name, string? itemId) => Error(
        ProviderErrorCode.AlreadyExists,
        $"A document named '{name}' already exists in that Box folder; choose another name.",
        itemId);

    /// <summary>
    /// Box confirmed the upload but the registration store failed: carry the file's name and
    /// content hash so it can be found and registered, never its Box id.
    /// </summary>
    private DocumentRegistrationFailedException RegistrationFailed(
        string message, string? itemId, string outputName, byte[] bytes, Exception innerException)
    {
        using var sha = System.Security.Cryptography.SHA256.Create();
        return new(message, ProviderName, ConnectionId, itemId, outputName,
            string.Concat(sha.ComputeHash(bytes).Select(b => b.ToString("x2"))), innerException);
    }

    private DocumentProviderException Error(
        ProviderErrorCode code,
        string message,
        string? itemId = null,
        Exception? innerException = null) =>
        new(code, message, ProviderName, ConnectionId, itemId, innerException);

    private sealed class BoxFile
    {
        public string Id { get; init; } = string.Empty;
        public string Type { get; init; } = string.Empty;
        public string Name { get; init; } = string.Empty;
        public long Size { get; init; }
        public string ETag { get; init; } = string.Empty;
        public string ParentId { get; init; } = string.Empty;
        public string ItemStatus { get; init; } = string.Empty;
    }
}
