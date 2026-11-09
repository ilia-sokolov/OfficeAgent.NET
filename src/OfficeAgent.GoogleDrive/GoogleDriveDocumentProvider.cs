using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using OfficeAgent.Abstractions;
using OfficeAgent.Core.DocumentProviders;

namespace OfficeAgent.GoogleDrive;

/// <summary>Configures one Google Drive provider connection.</summary>
public sealed class GoogleDriveDocumentProviderOptions
{
    /// <summary>Gets or sets the provider connection identifier.</summary>
    public string ConnectionId { get; set; } = string.Empty;

    /// <summary>Gets or sets the Drive API v3 base URL.</summary>
    public string ApiBaseUrl { get; set; } = "https://www.googleapis.com/drive/v3";

    /// <summary>Gets or sets the Drive API v3 media upload base URL.</summary>
    public string UploadBaseUrl { get; set; } = "https://www.googleapis.com/upload/drive/v3";

    /// <summary>Gets or sets the maximum accepted document size in bytes.</summary>
    public long MaximumBytes { get; set; } = 100 * 1024 * 1024;

    /// <summary>Gets or sets the allowed document extensions.</summary>
    public IReadOnlyCollection<string> AllowedExtensions { get; set; } = new[] { ".docx" };

    /// <summary>
    /// Gets or sets the Drive folder id that receives new documents. Empty means the
    /// connection cannot create documents; registering existing ones still works.
    /// </summary>
    public string CreationFolderId { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the Drive folder id of a templates folder. A templates listing reads it, and
    /// every file listed from it counts as a template. Empty means the connection has no
    /// templates folder.
    /// </summary>
    public string TemplatesFolder { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets where <see cref="TemplatesFolder"/> sits, such as
    /// <see cref="TemplateLocationKind.Organization"/> for the organization's approved templates.
    /// Unset means <see cref="TemplateLocationKind.Connection"/> when a templates folder is set.
    /// </summary>
    public TemplateLocationKind? TemplateLocation { get; set; }

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
/// Google Drive document provider over the Drive API v3 REST endpoints: a registry of
/// references to Office files (<c>.docx</c>, <c>.pptx</c>, <c>.xlsx</c>) in My Drive or a
/// shared drive, addressed by an opaque id assigned at registration time. A document is
/// registered by a Drive file id or a Drive or Docs URL. Native Google Docs, Sheets, and
/// Slides are refused by MIME type and never converted.
/// </summary>
/// <remarks>
/// <para>
/// Versions are the file's <c>headRevisionId</c>, which changes with every content upload.
/// Drive API v3 offers no server-side precondition on a content update (v3 files carry no
/// ETag), so <see cref="SaveDocumentOptions.Mode"/> <c>Replace</c> re-reads the head revision
/// immediately before uploading and refuses a mismatch. A write that lands between that check
/// and the upload is overwritten; Drive keeps it as an earlier revision in the file's version
/// history, where it can be restored.
/// </para>
/// <para>
/// Drive allows two files with the same name in a folder, so "never overwrite" on creation is
/// guaranteed by Drive (an upload always makes a new file) and the
/// <see cref="ProviderErrorCode.AlreadyExists"/> refusal is a name check made just before the
/// upload. Removing a registration never deletes Drive content.
/// </para>
/// <para>
/// Registration has no folder scope: any file the identity can reach can be registered. Listing
/// does: it reads the creation folder, or the templates folder for a templates listing, and the
/// folders below it, with <c>files.list</c>.
/// </para>
/// </remarks>
public sealed class GoogleDriveDocumentProvider : IDocumentProvider, IDocumentCreatingProvider, IDocumentListingProvider, IConnectionTemplateSource, IConnectionEditingDefaults, IConnectionGranularityDefaults, IConnectionCalculationDefaults
{
    /// <summary>Gets the provider discriminator used in document references.</summary>
    public const string ProviderName = "googledrive";

    private const string GoogleAppsMimePrefix = "application/vnd.google-apps.";
    private const string FolderMimeType = "application/vnd.google-apps.folder";
    private const string ShortcutMimeType = "application/vnd.google-apps.shortcut";
    private const string FileFields =
        "id,name,mimeType,size,version,headRevisionId,parents,driveId,trashed," +
        "capabilities(canModifyContent),contentRestrictions(readOnly)";
    private const string ListFields = "nextPageToken,files(id,name,mimeType,size,modifiedTime,trashed)";

    private static readonly HttpMethod Patch = new("PATCH");

    private readonly GoogleDriveDocumentProviderOptions _options;
    private readonly HttpClient _http;
    private readonly IGoogleDriveAccessTokenProvider _tokens;
    private readonly IGoogleDriveRegistrationStore _store;
    private readonly string _apiBaseUrl;
    private readonly string _uploadBaseUrl;
    private readonly HashSet<string> _allowedExtensions;
    private readonly string? _creationFolderId;
    private readonly string? _templatesFolderId;
    private readonly TemplateLocationKind _templateLocation;

    /// <summary>Initializes a provider over the Drive files the token's identity can reach.</summary>
    public GoogleDriveDocumentProvider(
        GoogleDriveDocumentProviderOptions options,
        HttpClient httpClient,
        IGoogleDriveAccessTokenProvider tokenProvider,
        IGoogleDriveRegistrationStore? registrationStore = null)
    {
        _options = options ?? throw new ArgumentNullException(nameof(options));
        _http = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
        _tokens = tokenProvider ?? throw new ArgumentNullException(nameof(tokenProvider));
        _store = registrationStore ?? new InMemoryGoogleDriveRegistrationStore();

        if (string.IsNullOrWhiteSpace(options.ConnectionId))
            throw new ArgumentException("A Google Drive provider connection id is required.", nameof(options));
        if (string.IsNullOrWhiteSpace(options.ApiBaseUrl) || string.IsNullOrWhiteSpace(options.UploadBaseUrl))
            throw new ArgumentException("The Drive API and upload base URLs are required.", nameof(options));
        if (options.MaximumBytes <= 0)
            throw new ArgumentOutOfRangeException(nameof(options), "MaximumBytes must be positive.");
        if (options.AllowedExtensions is null || options.AllowedExtensions.Count == 0)
            throw new ArgumentException("At least one allowed document extension is required.", nameof(options));

        _allowedExtensions = new HashSet<string>(
            options.AllowedExtensions.Select(NormalizeExtension),
            StringComparer.OrdinalIgnoreCase);
        _apiBaseUrl = options.ApiBaseUrl.TrimEnd('/');
        _uploadBaseUrl = options.UploadBaseUrl.TrimEnd('/');
        _creationFolderId = string.IsNullOrWhiteSpace(options.CreationFolderId)
            ? null
            : options.CreationFolderId.Trim();
        _templatesFolderId = string.IsNullOrWhiteSpace(options.TemplatesFolder)
            ? null
            : options.TemplatesFolder.Trim();
        if (_templatesFolderId is not null && !IsFileId(_templatesFolderId))
            throw new ArgumentException("TemplatesFolder must be a Google Drive folder id.", nameof(options));
        _templateLocation = CloudListing.TemplateLocationFor(_templatesFolderId is not null, options.TemplateLocation, nameof(options));
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
    /// <remarks>The templates folder's Drive folder id.</remarks>
    public string? TemplatesFolder => _templatesFolderId;

    /// <inheritdoc />
    public TemplateLocationKind TemplateLocation => _templateLocation;

    /// <inheritdoc />
    public async Task<DocumentReference> RegisterAsync(
        string source,
        CancellationToken cancellationToken = default)
    {
        var fileId = ParseSource(source);
        var file = await GetFileAsync(fileId, itemId: null, cancellationToken).ConfigureAwait(false);
        EnsureOfficeFile(file, itemId: null);

        var id = await _store.AddAsync(new GoogleDriveFileRef(file.Id), cancellationToken).ConfigureAwait(false);
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
                "This Google Drive connection has no creation folder. Configure CreationFolderId, " +
                "or register an existing document instead.");

        var fileName = ValidateCreateName(name);
        var bytes = await ReadStreamAsync(content, itemId: null, cancellationToken).ConfigureAwait(false);
        DriveFile created;
        try
        {
            var folder = await GetFileAsync(_creationFolderId, itemId: null, cancellationToken).ConfigureAwait(false);
            if (!string.Equals(folder.MimeType, FolderMimeType, StringComparison.Ordinal) || folder.Trashed)
                throw Error(ProviderErrorCode.ConfigurationError,
                    $"The configured Google Drive creation destination for '{fileName}' is not a folder.");
            if (await NameExistsAsync(folder.Id, folder.DriveId, fileName, itemId: null, cancellationToken).ConfigureAwait(false))
                throw AlreadyExists(fileName, itemId: null);

            created = await UploadNewAsync(folder.Id, fileName, bytes, itemId: null, cancellationToken).ConfigureAwait(false);
        }
        catch (DocumentProviderException ex) when (
            ex.Code is not (ProviderErrorCode.AlreadyExists or ProviderErrorCode.ConfigurationError))
        {
            throw Error(ex.Code, SafeCreateFailureMessage(ex.Code, fileName), itemId: null, ex);
        }
        catch (Exception ex) when (ex is not OperationCanceledException and not DocumentProviderException)
        {
            // Transport and token-provider failures can include the request URL, which
            // contains the host-configured folder id. Keep it on the inner exception only.
            throw Error(
                ProviderErrorCode.IO,
                $"Google Drive could not create document '{fileName}' in the configured folder.",
                itemId: null,
                ex);
        }

        string registrationId;
        try
        {
            // Drive confirmed the upload: finish registering even if the caller cancels now.
            registrationId = await _store.AddAsync(new GoogleDriveFileRef(created.Id), CancellationToken.None).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            throw RegistrationFailed(
                $"Document '{fileName}' was created in Google Drive, but its registration could not be " +
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
        EnsureOfficeFile(file, reference.ItemId);

        var current = VersionOf(file);
        if (!string.IsNullOrEmpty(reference.Version) &&
            !string.Equals(reference.Version, current, StringComparison.Ordinal))
            throw new DocumentVersionConflictException(reference.Version!, current, ProviderName, ConnectionId, reference.ItemId);

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
                "Google Drive document ids are opaque and provider-assigned; a destination id cannot be chosen by the caller. Use Mode=NewVersion/NewDocument (optionally with NewName) to mint a new document, or Mode=Replace to overwrite the source.",
                source.ItemId);
        if (options.Mode == SaveMode.Replace && !string.IsNullOrWhiteSpace(options.NewName))
            throw Error(ProviderErrorCode.InvalidArgument,
                "Replace mode overwrites the source in place and cannot rename. Use Mode=NewVersion/NewDocument with NewName to mint a renamed document.",
                source.ItemId);

        var registered = await ResolveRegistrationAsync(source.ItemId, cancellationToken).ConfigureAwait(false);
        // Read the output first, so the version check below is the last request before the upload.
        var bytes = await ReadStreamAsync(content, source.ItemId, cancellationToken).ConfigureAwait(false);
        var file = await GetFileAsync(registered.FileId, source.ItemId, cancellationToken).ConfigureAwait(false);
        EnsureOfficeFile(file, source.ItemId);

        var current = VersionOf(file);
        var expectedVersion = options.ExpectedVersion ?? source.Version;
        if (!string.IsNullOrEmpty(expectedVersion) &&
            !string.Equals(expectedVersion, current, StringComparison.Ordinal))
            throw new DocumentVersionConflictException(expectedVersion!, current, ProviderName, ConnectionId, source.ItemId);

        if (options.Mode == SaveMode.Replace)
        {
            if (file.ReadOnlyRestricted)
                throw Error(ProviderErrorCode.WriteRejected,
                    "The document is locked in Google Drive: a content restriction makes it read-only, so nothing " +
                    "was stored. Ask its owner to unlock it, or save as a new document instead.",
                    source.ItemId);
            if (file.CanModifyContent == false)
                throw Error(ProviderErrorCode.AccessDenied,
                    "The configured Google identity cannot change this file's content; nothing was stored. " +
                    "Share the file with it as an editor, or save as a new document instead.",
                    source.ItemId);

            var replaced = await UploadReplaceAsync(file, bytes, source.ItemId, cancellationToken).ConfigureAwait(false);
            return CreateReference(source.ItemId, replaced);
        }

        var parentId = file.Parents.FirstOrDefault();
        if (string.IsNullOrEmpty(parentId))
            throw Error(ProviderErrorCode.AccessDenied,
                "The configured Google identity cannot see the folder this file is in, so it cannot save a " +
                "sibling document; nothing was stored. Share the folder with it, or use Mode=Replace.",
                source.ItemId);

        string destinationName;
        if (string.IsNullOrWhiteSpace(options.NewName))
        {
            destinationName = await NextVersionedNameAsync(parentId!, file, source.ItemId, cancellationToken).ConfigureAwait(false);
        }
        else
        {
            destinationName = ValidateName(options.NewName!);
            if (await NameExistsAsync(parentId!, file.DriveId, destinationName, source.ItemId, cancellationToken).ConfigureAwait(false))
                throw AlreadyExists(destinationName, source.ItemId);
        }

        var saved = await UploadNewAsync(parentId!, destinationName, bytes, source.ItemId, cancellationToken).ConfigureAwait(false);
        string newId;
        try
        {
            newId = await _store.AddAsync(new GoogleDriveFileRef(saved.Id), CancellationToken.None).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            throw RegistrationFailed(
                $"Document '{destinationName}' was saved to Google Drive, but its registration could not be " +
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

    /// <inheritdoc />
    /// <remarks>
    /// Lists the creation folder, or the templates folder for a templates listing, with
    /// <c>files.list</c> and a <c>'folder' in parents and trashed = false</c> query. A connection
    /// without that folder reports <see cref="ProviderErrorCode.ConfigurationError"/>, or an empty
    /// page for templates. <see cref="DocumentListQuery.Folder"/> is a path of folder names below
    /// it, found with one name query per name. Each entry's source is the Drive file id. Native
    /// Google files, shortcuts and files over the size cap are not listed. A page makes at most
    /// <see cref="CloudListing.MaximumServiceCalls"/> requests, folder lookups included.
    /// </remarks>
    public async Task<DocumentListPage> ListAsync(DocumentListQuery query, CancellationToken cancellationToken = default)
    {
        var request = CloudListingRequest.Create(query, _allowedExtensions, out var problem);
        if (request is null)
        {
            var (code, message) = CloudListing.FolderError(problem);
            throw Error(code, message);
        }
        var rootId = request.Templates ? _templatesFolderId : _creationFolderId;
        if (rootId is null)
        {
            if (request.Templates) return new DocumentListPage();
            throw Error(ProviderErrorCode.ConfigurationError,
                "This Google Drive connection has no folder to list. Configure CreationFolderId, or TemplatesFolder for templates.");
        }

        try
        {
            var calls = CloudListing.MaximumServiceCalls;
            CloudListingCursor cursor;
            if (string.IsNullOrEmpty(query.ContinuationToken))
            {
                var root = await GetFileAsync(rootId, itemId: null, cancellationToken).ConfigureAwait(false);
                calls--;
                if (!string.Equals(root.MimeType, FolderMimeType, StringComparison.Ordinal) || root.Trashed)
                    throw Error(ProviderErrorCode.ConfigurationError, "The configured Google Drive listing folder is not a folder.");

                var folderId = root.Id;
                foreach (var segment in request.FolderSegments)
                {
                    if (calls <= 1) throw Error(ProviderErrorCode.InvalidArgument, CloudListing.TooDeepMessage);
                    folderId = await FindFolderAsync(folderId, root.DriveId, segment, cancellationToken).ConfigureAwait(false)
                        ?? throw Error(ProviderErrorCode.NotFound, $"No folder named '{segment}' is in the listed Google Drive folder.");
                    calls--;
                }
                cursor = CloudListing.Start(folderId, request, root.DriveId);
            }
            else if (!CloudListingCursor.TryUnprotect(query.ContinuationToken!, ConnectionId, request.Fingerprint, out cursor))
            {
                throw Error(ProviderErrorCode.InvalidArgument, CloudListing.InvalidTokenMessage);
            }

            var driveId = cursor.Context;
            return await CloudListing.WalkAsync(
                request, cursor,
                (folderId, pageToken, ct) => ListChildrenAsync(folderId, driveId, pageToken, ct),
                calls, ConnectionId, cancellationToken).ConfigureAwait(false);
        }
        catch (Exception ex) when (ex is not OperationCanceledException and not DocumentProviderException)
        {
            // Transport and token failures can carry request URLs, which name folder ids.
            throw Error(ProviderErrorCode.IO, "Google Drive could not be reached to list the folder.", itemId: null, ex);
        }
    }

    // ── Source parsing ────────────────────────────────────────────────────────

    /// <summary>
    /// Resolves a registration source to a Drive file id. Accepts a bare file id, a Drive URL
    /// (<c>drive.google.com/file/d/{id}/…</c>, <c>drive.google.com/open?id={id}</c>), or a Docs
    /// editor URL (<c>docs.google.com/document/d/{id}/…</c>), which is how Drive opens an
    /// Office file in the browser.
    /// </summary>
    private string ParseSource(string source)
    {
        if (string.IsNullOrWhiteSpace(source))
            throw Error(ProviderErrorCode.InvalidArgument,
                "A Google Drive registration requires a source: a Drive file id or a Drive file URL " +
                "(https://drive.google.com/file/d/.../view).");

        var trimmed = source.Trim();
        if (!trimmed.StartsWith("http://", StringComparison.OrdinalIgnoreCase) &&
            !trimmed.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
        {
            if (!IsFileId(trimmed))
                throw Error(ProviderErrorCode.InvalidArgument,
                    "Register a Google Drive document by its file id or its Drive URL, not by a name or path.");
            return trimmed;
        }

        if (!Uri.TryCreate(trimmed, UriKind.Absolute, out var uri) ||
            !(uri.Host.Equals("drive.google.com", StringComparison.OrdinalIgnoreCase) ||
              uri.Host.Equals("docs.google.com", StringComparison.OrdinalIgnoreCase)))
            throw Error(ProviderErrorCode.InvalidArgument,
                "The URL is not a Google Drive or Google Docs file URL.");

        var segments = uri.AbsolutePath.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
        for (var i = 0; i < segments.Length - 1; i++)
        {
            if (segments[i] == "d" && IsFileId(segments[i + 1]))
                return segments[i + 1];
        }

        foreach (var pair in uri.Query.TrimStart('?').Split('&'))
        {
            var equals = pair.IndexOf('=');
            if (equals > 0 && pair.Substring(0, equals) == "id")
            {
                var value = Uri.UnescapeDataString(pair.Substring(equals + 1));
                if (IsFileId(value)) return value;
            }
        }

        if (segments.Contains("folders"))
            throw Error(ProviderErrorCode.InvalidArgument,
                "The URL names a Drive folder. Register a file instead.");
        throw Error(ProviderErrorCode.InvalidArgument,
            "The URL does not name a Drive file. Use the file's share link or its id.");
    }

    private static bool IsFileId(string value) =>
        value.Length is > 0 and <= 200 &&
        value.All(c => c is (>= 'a' and <= 'z') or (>= 'A' and <= 'Z') or (>= '0' and <= '9') or '-' or '_');

    // ── Drive plumbing ────────────────────────────────────────────────────────

    private string FileUrl(string fileId) =>
        $"{_apiBaseUrl}/files/{Uri.EscapeDataString(fileId)}?supportsAllDrives=true&fields={Uri.EscapeDataString(FileFields)}";

    private async Task<DriveFile> GetFileAsync(string fileId, string? itemId, CancellationToken cancellationToken)
    {
        using var response = await SendAsync(HttpMethod.Get, FileUrl(fileId), content: null, cancellationToken).ConfigureAwait(false);
        await ThrowOnFailureAsync(response, itemId).ConfigureAwait(false);
        return ParseFile(await response.Content.ReadAsStringAsync().ConfigureAwait(false), itemId);
    }

    private async Task<byte[]> DownloadAsync(string fileId, string itemId, CancellationToken cancellationToken)
    {
        var url = $"{_apiBaseUrl}/files/{Uri.EscapeDataString(fileId)}?alt=media&supportsAllDrives=true";
        using var response = await SendAsync(HttpMethod.Get, url, content: null, cancellationToken,
            HttpCompletionOption.ResponseHeadersRead).ConfigureAwait(false);
        await ThrowOnFailureAsync(response, itemId).ConfigureAwait(false);
        if (response.Content.Headers.ContentLength is { } length && length > _options.MaximumBytes)
            throw TooLarge(itemId);

        var stream = await response.Content.ReadAsStreamAsync().ConfigureAwait(false);
        using (stream)
        {
            return await ReadBoundedAsync(stream, itemId, cancellationToken).ConfigureAwait(false);
        }
    }

    private async Task<DriveFile> UploadReplaceAsync(DriveFile file, byte[] bytes, string itemId, CancellationToken cancellationToken)
    {
        var url = $"{_uploadBaseUrl}/files/{Uri.EscapeDataString(file.Id)}?uploadType=media&supportsAllDrives=true&fields={Uri.EscapeDataString(FileFields)}";
        using var body = new ByteArrayContent(bytes);
        body.Headers.ContentType = new MediaTypeHeaderValue(
            string.IsNullOrEmpty(file.MimeType) ? OfficeContentTypes.ForName(file.Name) : file.MimeType);
        using var response = await SendAsync(Patch, url, body, cancellationToken).ConfigureAwait(false);
        await ThrowOnFailureAsync(response, itemId, write: true).ConfigureAwait(false);
        return ParseFile(await response.Content.ReadAsStringAsync().ConfigureAwait(false), itemId);
    }

    private async Task<DriveFile> UploadNewAsync(
        string parentId, string name, byte[] bytes, string? itemId, CancellationToken cancellationToken)
    {
        var url = $"{_uploadBaseUrl}/files?uploadType=multipart&supportsAllDrives=true&fields={Uri.EscapeDataString(FileFields)}";
        var contentType = OfficeContentTypes.ForName(name);
        // Naming the Office MIME type explicitly keeps Drive from converting the upload.
        var metadata = JsonSerializer.Serialize(new Dictionary<string, object>
        {
            ["name"] = name,
            ["parents"] = new[] { parentId },
            ["mimeType"] = contentType
        });

        using var body = new MultipartContent("related", "officeagent-" + Guid.NewGuid().ToString("N"));
        var metadataPart = new StringContent(metadata, Encoding.UTF8);
        metadataPart.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json; charset=UTF-8");
        body.Add(metadataPart);
        var mediaPart = new ByteArrayContent(bytes);
        mediaPart.Headers.ContentType = new MediaTypeHeaderValue(contentType);
        body.Add(mediaPart);

        using var response = await SendAsync(HttpMethod.Post, url, body, cancellationToken).ConfigureAwait(false);
        await ThrowOnFailureAsync(response, itemId, write: true).ConfigureAwait(false);
        return ParseFile(await response.Content.ReadAsStringAsync().ConfigureAwait(false), itemId);
    }

    /// <summary>Whether a file (not in the trash) with exactly this name is in the folder.</summary>
    private async Task<bool> NameExistsAsync(
        string folderId, string? driveId, string name, string? itemId, CancellationToken cancellationToken)
    {
        var query = $"name = '{EscapeQuery(name)}' and '{EscapeQuery(folderId)}' in parents and trashed = false";
        var url = new StringBuilder($"{_apiBaseUrl}/files?q={Uri.EscapeDataString(query)}")
            .Append("&fields=").Append(Uri.EscapeDataString("files(id,name)"))
            .Append("&pageSize=10&supportsAllDrives=true&includeItemsFromAllDrives=true");
        if (!string.IsNullOrEmpty(driveId))
            url.Append("&corpora=drive&driveId=").Append(Uri.EscapeDataString(driveId));

        using var response = await SendAsync(HttpMethod.Get, url.ToString(), content: null, cancellationToken).ConfigureAwait(false);
        await ThrowOnFailureAsync(response, itemId).ConfigureAwait(false);
        try
        {
            using var json = JsonDocument.Parse(await response.Content.ReadAsStringAsync().ConfigureAwait(false));
            return json.RootElement.TryGetProperty("files", out var files) &&
                   files.ValueKind == JsonValueKind.Array &&
                   files.EnumerateArray().Any(f =>
                       f.TryGetProperty("name", out var n) &&
                       string.Equals(n.GetString(), name, StringComparison.OrdinalIgnoreCase));
        }
        catch (JsonException ex)
        {
            throw Error(ProviderErrorCode.IO, "Google Drive returned an unreadable file list.", itemId, ex);
        }
    }

    /// <summary>One page of a folder's children, folders included, so a recursive listing can descend.</summary>
    private async Task<CloudListingPage> ListChildrenAsync(
        string folderId, string? driveId, string? pageToken, CancellationToken cancellationToken)
    {
        var query = $"'{EscapeQuery(folderId)}' in parents and trashed = false";
        var url = new StringBuilder($"{_apiBaseUrl}/files?q={Uri.EscapeDataString(query)}")
            .Append("&fields=").Append(Uri.EscapeDataString(ListFields))
            .Append("&pageSize=").Append(CloudListing.ServicePageSize.ToString(System.Globalization.CultureInfo.InvariantCulture))
            .Append("&supportsAllDrives=true&includeItemsFromAllDrives=true");
        if (!string.IsNullOrEmpty(driveId))
            url.Append("&corpora=drive&driveId=").Append(Uri.EscapeDataString(driveId));
        if (!string.IsNullOrEmpty(pageToken))
            url.Append("&pageToken=").Append(Uri.EscapeDataString(pageToken));

        using var response = await SendAsync(HttpMethod.Get, url.ToString(), content: null, cancellationToken).ConfigureAwait(false);
        await ThrowOnFailureAsync(response, itemId: null).ConfigureAwait(false);
        try
        {
            using var json = JsonDocument.Parse(await response.Content.ReadAsStringAsync().ConfigureAwait(false));
            var root = json.RootElement;
            var children = new List<CloudListingChild>();
            if (root.TryGetProperty("files", out var files) && files.ValueKind == JsonValueKind.Array)
            {
                foreach (var file in files.EnumerateArray())
                {
                    var id = StringOf(file, "id");
                    var mimeType = StringOf(file, "mimeType");
                    long? size = file.TryGetProperty("size", out var sizeValue) ? ParseLong(sizeValue) : null;
                    children.Add(new CloudListingChild
                    {
                        Id = id,
                        Name = StringOf(file, "name"),
                        IsFolder = string.Equals(mimeType, FolderMimeType, StringComparison.Ordinal) && IsFileId(id),
                        // Native files and shortcuts cannot be registered, so they are not offered.
                        Unusable = !IsFileId(id) ||
                                   mimeType.StartsWith(GoogleAppsMimePrefix, StringComparison.Ordinal) ||
                                   (file.TryGetProperty("trashed", out var trashed) && trashed.ValueKind == JsonValueKind.True) ||
                                   size > _options.MaximumBytes,
                        Length = size,
                        LastModified = ParseTime(StringOf(file, "modifiedTime"))
                    });
                }
            }
            var next = StringOf(root, "nextPageToken");
            return new CloudListingPage { Items = children, NextToken = next.Length == 0 ? null : next };
        }
        catch (Exception ex) when (ex is JsonException or InvalidOperationException or FormatException)
        {
            throw Error(ProviderErrorCode.IO, "Google Drive returned an unreadable file list.", itemId: null, ex);
        }
    }

    /// <summary>The id of the subfolder with this name, or <see langword="null"/> when there is none.</summary>
    private async Task<string?> FindFolderAsync(string parentId, string? driveId, string name, CancellationToken cancellationToken)
    {
        var query = $"name = '{EscapeQuery(name)}' and '{EscapeQuery(parentId)}' in parents " +
                    $"and mimeType = '{FolderMimeType}' and trashed = false";
        var url = new StringBuilder($"{_apiBaseUrl}/files?q={Uri.EscapeDataString(query)}")
            .Append("&fields=").Append(Uri.EscapeDataString("files(id,name)"))
            .Append("&pageSize=10&supportsAllDrives=true&includeItemsFromAllDrives=true");
        if (!string.IsNullOrEmpty(driveId))
            url.Append("&corpora=drive&driveId=").Append(Uri.EscapeDataString(driveId));

        using var response = await SendAsync(HttpMethod.Get, url.ToString(), content: null, cancellationToken).ConfigureAwait(false);
        await ThrowOnFailureAsync(response, itemId: null).ConfigureAwait(false);
        try
        {
            using var json = JsonDocument.Parse(await response.Content.ReadAsStringAsync().ConfigureAwait(false));
            if (!json.RootElement.TryGetProperty("files", out var files) || files.ValueKind != JsonValueKind.Array)
                return null;
            foreach (var file in files.EnumerateArray())
            {
                var id = StringOf(file, "id");
                if (IsFileId(id) && string.Equals(StringOf(file, "name"), name, StringComparison.OrdinalIgnoreCase))
                    return id;
            }
            return null;
        }
        catch (JsonException ex)
        {
            throw Error(ProviderErrorCode.IO, "Google Drive returned an unreadable file list.", itemId: null, ex);
        }
    }

    private static DateTimeOffset? ParseTime(string value) =>
        DateTimeOffset.TryParse(value, System.Globalization.CultureInfo.InvariantCulture,
            System.Globalization.DateTimeStyles.AssumeUniversal, out var parsed)
            ? parsed
            : null;

    private static string EscapeQuery(string value) => value.Replace("\\", "\\\\").Replace("'", "\\'");

    private async Task<HttpResponseMessage> SendAsync(
        HttpMethod method,
        string url,
        HttpContent? content,
        CancellationToken cancellationToken,
        HttpCompletionOption completion = HttpCompletionOption.ResponseContentRead)
    {
        var token = await _tokens.GetAccessTokenAsync(cancellationToken).ConfigureAwait(false);
        using var request = new HttpRequestMessage(method, url);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        if (content is not null)
            request.Content = content;
        return await _http.SendAsync(request, completion, cancellationToken).ConfigureAwait(false);
    }

    private async Task ThrowOnFailureAsync(HttpResponseMessage response, string? itemId, bool write = false)
    {
        if (response.IsSuccessStatusCode) return;

        var reason = await ReadErrorReasonAsync(response).ConfigureAwait(false);
        var suffix = string.IsNullOrEmpty(reason) ? string.Empty : $" Drive reason: {reason}.";
        var status = (int)response.StatusCode;
        throw response.StatusCode switch
        {
            HttpStatusCode.NotFound => Error(ProviderErrorCode.NotFound,
                $"The referenced Google Drive file does not exist or is not visible to the configured identity.{suffix}", itemId),
            // Drive reports rate limits and quota as 403 as well as 429: a refusal, not a permission problem.
            _ when IsThrottleOrQuota(reason) || response.StatusCode == (HttpStatusCode)429 => write
                ? Error(ProviderErrorCode.WriteRejected,
                    $"Google Drive refused the upload (status {status}, rate limit or quota); nothing was stored. Try again later.{suffix}", itemId)
                : Error(ProviderErrorCode.IO,
                    $"Google Drive throttled the request (status {status}). Try again later.{suffix}", itemId),
            HttpStatusCode.Unauthorized or HttpStatusCode.Forbidden => Error(ProviderErrorCode.AccessDenied,
                $"Google Drive refused access to the file.{suffix}", itemId),
            HttpStatusCode.PreconditionFailed => new DocumentVersionConflictException(
                "(precondition)", "(changed)", ProviderName, ConnectionId, itemId),
            HttpStatusCode.RequestEntityTooLarge => Error(ProviderErrorCode.ContentTooLarge,
                $"Google Drive rejected the content as too large.{suffix}", itemId),
            // Any other refusal of an upload is Drive declining it: nothing was stored. A server
            // error is not a refusal - Drive may have processed the upload - so it is unknown.
            _ when write && status is >= 400 and < 500 => Error(ProviderErrorCode.WriteRejected,
                $"Google Drive refused the upload with status {status}; nothing was stored.{suffix}", itemId),
            _ when write => Error(ProviderErrorCode.OutcomeUnknown,
                $"Google Drive failed the upload with status {status}; it may have been stored.{suffix}", itemId),
            _ => Error(ProviderErrorCode.IO,
                $"Google Drive request failed with status {status}.{suffix}", itemId)
        };
    }

    private static bool IsThrottleOrQuota(string reason) => reason is
        "rateLimitExceeded" or "userRateLimitExceeded" or "sharingRateLimitExceeded" or
        "dailyLimitExceeded" or "storageQuotaExceeded" or "teamDriveFileLimitExceeded" or
        "numChildrenInNonRootLimitExceeded";

    /// <summary>
    /// Reads the machine reason (<c>error.errors[0].reason</c>, else <c>error.status</c>) from a
    /// Drive error body. The human message is not copied: it can quote ids and names.
    /// </summary>
    private static async Task<string> ReadErrorReasonAsync(HttpResponseMessage response)
    {
        try
        {
            var payload = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            if (string.IsNullOrWhiteSpace(payload)) return string.Empty;
            using var json = JsonDocument.Parse(payload);
            if (!json.RootElement.TryGetProperty("error", out var error) || error.ValueKind != JsonValueKind.Object)
                return string.Empty;
            if (error.TryGetProperty("errors", out var errors) && errors.ValueKind == JsonValueKind.Array)
                foreach (var entry in errors.EnumerateArray())
                    if (entry.TryGetProperty("reason", out var reason) && reason.ValueKind == JsonValueKind.String)
                        return reason.GetString() ?? string.Empty;
            return error.TryGetProperty("status", out var status) && status.ValueKind == JsonValueKind.String
                ? status.GetString() ?? string.Empty
                : string.Empty;
        }
        catch (JsonException)
        {
            return string.Empty;
        }
    }

    private DriveFile ParseFile(string payload, string? itemId)
    {
        try
        {
            using var json = JsonDocument.Parse(payload);
            var root = json.RootElement;
            return new DriveFile
            {
                Id = root.GetProperty("id").GetString() ?? string.Empty,
                Name = StringOf(root, "name"),
                MimeType = StringOf(root, "mimeType"),
                Size = root.TryGetProperty("size", out var size) ? ParseLong(size) : 0,
                Version = root.TryGetProperty("version", out var version) ? ParseLong(version) : 0,
                HeadRevisionId = StringOf(root, "headRevisionId"),
                DriveId = root.TryGetProperty("driveId", out var driveId) ? driveId.GetString() : null,
                Parents = root.TryGetProperty("parents", out var parents) && parents.ValueKind == JsonValueKind.Array
                    ? parents.EnumerateArray().Select(p => p.GetString() ?? string.Empty).Where(p => p.Length > 0).ToArray()
                    : Array.Empty<string>(),
                Trashed = root.TryGetProperty("trashed", out var trashed) && trashed.ValueKind == JsonValueKind.True,
                CanModifyContent = root.TryGetProperty("capabilities", out var capabilities) &&
                                   capabilities.TryGetProperty("canModifyContent", out var canModify) &&
                                   canModify.ValueKind is JsonValueKind.True or JsonValueKind.False
                    ? canModify.GetBoolean()
                    : null,
                ReadOnlyRestricted = root.TryGetProperty("contentRestrictions", out var restrictions) &&
                                     restrictions.ValueKind == JsonValueKind.Array &&
                                     restrictions.EnumerateArray().Any(r =>
                                         r.TryGetProperty("readOnly", out var readOnly) &&
                                         readOnly.ValueKind == JsonValueKind.True)
            };
        }
        catch (Exception ex) when (ex is JsonException or KeyNotFoundException or InvalidOperationException or FormatException)
        {
            throw Error(ProviderErrorCode.IO, "Google Drive returned an unreadable file resource.", itemId, ex);
        }
    }

    private static string StringOf(JsonElement element, string name) =>
        element.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String
            ? value.GetString() ?? string.Empty
            : string.Empty;

    // Drive returns int64 fields as JSON strings.
    private static long ParseLong(JsonElement value) => value.ValueKind == JsonValueKind.String
        ? long.Parse(value.GetString()!, System.Globalization.CultureInfo.InvariantCulture)
        : value.GetInt64();

    /// <summary>
    /// The version a reference carries: the head revision id, which changes with every content
    /// upload and not with renames or sharing changes. A file without one (Drive reports it only
    /// for files with binary content) falls back to the file's change counter.
    /// </summary>
    private static string VersionOf(DriveFile file) =>
        string.IsNullOrEmpty(file.HeadRevisionId)
            ? "v" + file.Version.ToString(System.Globalization.CultureInfo.InvariantCulture)
            : file.HeadRevisionId;

    // ── Registration & validation ─────────────────────────────────────────────

    private async Task<GoogleDriveFileRef> ResolveRegistrationAsync(string registrationId, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(registrationId))
            throw Error(ProviderErrorCode.InvalidArgument, "An item id is required.");
        return await _store.ResolveAsync(registrationId, cancellationToken).ConfigureAwait(false)
            ?? throw Error(ProviderErrorCode.NotFound,
                $"The provider document '{registrationId}' is not registered.", registrationId);
    }

    /// <summary>
    /// Refuses what is not an Office file this connection accepts: trashed files, folders,
    /// shortcuts, native Google files, disallowed extensions, and files over the size cap.
    /// Runs on metadata alone, before any content is downloaded.
    /// </summary>
    private void EnsureOfficeFile(DriveFile file, string? itemId)
    {
        if (file.Trashed)
            throw Error(ProviderErrorCode.NotFound,
                $"Google Drive file '{file.Name}' is in the trash.", itemId);
        if (string.Equals(file.MimeType, FolderMimeType, StringComparison.Ordinal))
            throw Error(ProviderErrorCode.InvalidArgument,
                $"'{file.Name}' is a Google Drive folder, not a document.", itemId);
        if (string.Equals(file.MimeType, ShortcutMimeType, StringComparison.Ordinal))
            throw Error(ProviderErrorCode.InvalidArgument,
                $"'{file.Name}' is a Google Drive shortcut. Register the file it points to instead.", itemId);
        if (file.MimeType.StartsWith(GoogleAppsMimePrefix, StringComparison.Ordinal))
            throw Error(ProviderErrorCode.ExtensionNotAllowed,
                $"'{file.Name}' is a native {NativeKind(file.MimeType)} file, not an Office file. OfficeAgent edits " +
                "Word, PowerPoint and Excel files stored in Drive and does not convert Google files, because " +
                "converting loses comments and suggested edits. Download it as an Office file and upload that " +
                "copy to edit it here, or edit it in Google's own editor.",
                itemId);

        ValidateExtension(file.Name, itemId);
        if (file.Size > _options.MaximumBytes)
            throw Error(ProviderErrorCode.ContentTooLarge,
                $"Document '{file.Name}' exceeds the configured maximum of {_options.MaximumBytes} bytes.", itemId);
    }

    private static string NativeKind(string mimeType) => mimeType.Substring(GoogleAppsMimePrefix.Length) switch
    {
        "document" => "Google Docs",
        "spreadsheet" => "Google Sheets",
        "presentation" => "Google Slides",
        "form" => "Google Forms",
        "drawing" => "Google Drawings",
        _ => "Google Workspace"
    };

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
            $"The configured Google Drive creation folder for '{fileName}' was not found.",
        ProviderErrorCode.AccessDenied =>
            $"The configured Google identity cannot create '{fileName}' in the creation folder.",
        ProviderErrorCode.ContentTooLarge =>
            $"Google Drive rejected '{fileName}' because its content is too large.",
        ProviderErrorCode.WriteRejected =>
            $"Google Drive refused to create '{fileName}'; nothing was stored.",
        ProviderErrorCode.OutcomeUnknown =>
            $"Google Drive failed while creating '{fileName}'; it may have been stored.",
        _ => $"Google Drive could not create document '{fileName}' in the configured folder."
    };

    private void ValidateReference(DocumentReference reference)
    {
        if (reference is null) throw new ArgumentNullException(nameof(reference));
        if (!string.Equals(reference.Provider, ProviderName, StringComparison.OrdinalIgnoreCase) ||
            !string.Equals(reference.ConnectionId, ConnectionId, StringComparison.Ordinal))
            throw new DocumentProviderException(
                ProviderErrorCode.InvalidArgument,
                $"The document reference (provider='{reference.Provider}', connectionId='{reference.ConnectionId}') does not belong to this Google Drive connection (provider='{ProviderName}', connectionId='{ConnectionId}').",
                ProviderName, ConnectionId, reference.ItemId);
    }

    /// <summary>
    /// Picks the next unused versioned sibling name in the file's folder, matching the
    /// filesystem provider's convention: <c>contract.docx</c> → <c>contract.v2.docx</c>, ….
    /// </summary>
    private async Task<string> NextVersionedNameAsync(
        string parentId, DriveFile source, string itemId, CancellationToken cancellationToken)
    {
        var extension = Path.GetExtension(source.Name);
        var stem = Path.GetFileNameWithoutExtension(source.Name);
        var (baseStem, startVersion) = ParseVersionedStem(stem);

        for (var n = Math.Max(2, startVersion + 1); n < startVersion + 1000; n++)
        {
            var candidate = $"{baseStem}.v{n}{extension}";
            if (!await NameExistsAsync(parentId, source.DriveId, candidate, itemId, cancellationToken).ConfigureAwait(false))
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

    private DocumentReference CreateReference(string registrationId, DriveFile file) => new()
    {
        Provider = ProviderName,
        ConnectionId = ConnectionId,
        ItemId = registrationId,
        Version = VersionOf(file),
        Name = file.Name,
        ContentType = string.IsNullOrEmpty(file.MimeType) || file.MimeType == "application/octet-stream"
            ? OfficeContentTypes.ForName(file.Name)
            : file.MimeType
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
        $"A document named '{name}' already exists in that Google Drive folder; choose another name.",
        itemId);

    /// <summary>
    /// Drive confirmed the upload but the registration store failed: carry the file's name and
    /// content hash so it can be found and registered, never its Drive id.
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

    private sealed class DriveFile
    {
        public string Id { get; init; } = string.Empty;
        public string Name { get; init; } = string.Empty;
        public string MimeType { get; init; } = string.Empty;
        public long Size { get; init; }
        public long Version { get; init; }
        public string HeadRevisionId { get; init; } = string.Empty;
        public string? DriveId { get; init; }
        public IReadOnlyList<string> Parents { get; init; } = Array.Empty<string>();
        public bool Trashed { get; init; }
        public bool? CanModifyContent { get; init; }
        public bool ReadOnlyRestricted { get; init; }
    }
}
