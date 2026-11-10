using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using OfficeAgent.Abstractions;
using OfficeAgent.Core.DocumentProviders;

namespace OfficeAgent.IManage;

/// <summary>Configures one iManage Work provider connection.</summary>
public sealed class IManageDocumentProviderOptions
{
    /// <summary>Gets or sets the provider connection identifier.</summary>
    public string ConnectionId { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the iManage server URL, for example <c>https://cloudimanage.com</c> or your
    /// on-premises Work server. Registration URLs must be on this host.
    /// </summary>
    public string ServerUrl { get; set; } = string.Empty;

    /// <summary>Gets or sets the REST API base path on <see cref="ServerUrl"/>.</summary>
    public string ApiBasePath { get; set; } = "/work/api/v2";

    /// <summary>
    /// Gets or sets the iManage customer id that every Work API v2 path names. An administrator
    /// can read it from the server's <c>/api</c> discovery response.
    /// </summary>
    public string CustomerId { get; set; } = string.Empty;

    /// <summary>Gets or sets the maximum accepted document size in bytes.</summary>
    public long MaximumBytes { get; set; } = 100 * 1024 * 1024;

    /// <summary>Gets or sets the allowed document extensions.</summary>
    public IReadOnlyCollection<string> AllowedExtensions { get; set; } = new[] { ".docx" };

    /// <summary>
    /// Gets or sets the folder or workspace id, <c>LIBRARY!number</c>, that receives new
    /// documents: <c>CreateAsync</c> and <see cref="SaveMode.NewDocument"/> saves. Leave empty to
    /// disable creation.
    /// </summary>
    public string CreationFolderId { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the document class given to new documents, for libraries that require one.
    /// Leave empty to let the library apply its default.
    /// </summary>
    public string CreationDocumentClass { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the folder or workspace id, <c>LIBRARY!number</c>, of a templates folder. A
    /// templates listing reads it, and every document listed from it counts as a template. Empty
    /// means the connection has no templates folder.
    /// </summary>
    public string TemplatesFolder { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets where <see cref="TemplatesFolder"/> sits, such as
    /// <see cref="TemplateLocationKind.Organization"/> for the firm's approved precedents.
    /// Unset means <see cref="TemplateLocationKind.Connection"/> when a templates folder is set.
    /// </summary>
    public TemplateLocationKind? TemplateLocation { get; set; }

    /// <summary>
    /// Gets or sets whether the connection refuses every save and creation with
    /// <see cref="ProviderErrorCode.AccessDenied"/> before any request reaches iManage, whatever
    /// the configured credentials may write. Registering, opening, listing and using a document as
    /// a template still work, and registrations are kept as usual, because they live in
    /// the host's registration store rather than in iManage. Use it for a templates folder people share, beside a writable
    /// connection for the documents made from it. A read-only connection cannot name <see cref="CreationFolderId"/>. Defaults to <see langword="false"/>.
    /// </summary>
    public bool ReadOnly { get; set; }

    /// <summary>Gets or sets the comment recorded on each version and document OfficeAgent saves.</summary>
    public string SaveComment { get; set; } = "Saved by OfficeAgent.";

    /// <summary>
    /// Gets or sets whether a save locks the document while it checks the latest version and
    /// uploads the new one, so no other writer can slip a version in between. Defaults to
    /// <see langword="true"/>. Turn it off only for a server that does not accept document locks.
    /// </summary>
    public bool LockDuringSave { get; set; } = true;

    /// <summary>
    /// Gets or sets the change mode used when an operation does not state one. The default is
    /// <see cref="ChangeMode.Tracked"/>, which suits contracts a lawyer reviews.
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
    /// formulas calculated. Takes effect only when a formula evaluator is registered.
    /// </summary>
    public bool CalculateFormulas { get; set; }

    /// <summary>
    /// Gets or sets whether a plan whose calculation produces a new formula error is refused.
    /// Applies only with <see cref="CalculateFormulas"/>.
    /// </summary>
    public bool RejectFormulaErrors { get; set; }
}

/// <summary>
/// iManage Work document provider over the Work REST API v2: a registry of references to
/// iManage documents, addressed by an opaque id assigned at registration time. A document is
/// registered by its id (<c>LIBRARY!number.version</c>) or by an iManage Work URL that carries
/// it, and the registration follows the document's latest version.
/// </summary>
/// <remarks>
/// iManage keeps every version, so no save overwrites content. <see cref="SaveMode.Replace"/>
/// and <see cref="SaveMode.NewVersion"/> add a new version to the same document number; Replace
/// keeps the registration id and NewVersion mints another. <see cref="SaveMode.NewDocument"/>
/// and creation file a new document in the configured creation folder. Every save is conditional
/// on the version that was opened, and by default the document is locked while the save checks
/// and uploads. A document checked out by someone else is refused, never written. Edits address
/// opaque ids; server URLs and tokens never reach the agent. A listing is the exception for ids:
/// each entry's source is the document id registration takes, <c>LIBRARY!number.version</c>.
/// Removing a registration never deletes iManage content. Registration has no folder scope;
/// listing reads the creation folder, or the templates folder for a templates listing, and the
/// folders below it.
/// </remarks>
public sealed class IManageDocumentProvider : IDocumentProvider, IDocumentCreatingProvider, IDocumentListingProvider, IConnectionTemplateSource, IConnectionEditingDefaults, IConnectionGranularityDefaults, IConnectionCalculationDefaults
{
    /// <summary>Gets the provider discriminator used in document references.</summary>
    public const string ProviderName = "imanage";

    // iManage authenticates REST calls with this header. Unlike Authorization, HttpClient does
    // not strip it on a redirect, so the provider refuses to follow one with credentials.
    internal const string AuthHeader = "X-Auth-Token";

    private readonly IManageDocumentProviderOptions _options;
    private readonly HttpClient _http;
    private readonly IIManageAccessTokenProvider _tokens;
    private readonly IIManageRegistrationStore _store;
    private readonly HashSet<string> _allowedExtensions;
    private readonly Uri _server;
    private readonly string _apiRoot;
    private readonly string? _creationFolderId;
    private readonly string? _creationLibrary;
    private readonly string? _templatesFolderId;
    private readonly TemplateLocationKind _templateLocation;

    /// <summary>Initializes a provider over one iManage Work customer.</summary>
    public IManageDocumentProvider(
        IManageDocumentProviderOptions options,
        HttpClient httpClient,
        IIManageAccessTokenProvider tokenProvider,
        IIManageRegistrationStore? registrationStore = null)
    {
        _options = options ?? throw new ArgumentNullException(nameof(options));
        if (options.ReadOnly && !string.IsNullOrWhiteSpace(options.CreationFolderId))
            throw new ArgumentException(ReadOnlyConnections.CreationDestinationNamed("CreationFolderId"), nameof(options));
        _http = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
        _tokens = tokenProvider ?? throw new ArgumentNullException(nameof(tokenProvider));
        _store = registrationStore ?? new InMemoryIManageRegistrationStore();

        if (string.IsNullOrWhiteSpace(options.ConnectionId))
            throw new ArgumentException("An iManage provider connection id is required.", nameof(options));
        if (!Uri.TryCreate(options.ServerUrl, UriKind.Absolute, out var server) || server.Scheme != Uri.UriSchemeHttps)
            throw new ArgumentException("An absolute https iManage ServerUrl is required.", nameof(options));
        if (!IManageIds.IsLibrary(options.CustomerId))
            throw new ArgumentException("An iManage CustomerId is required.", nameof(options));
        if (string.IsNullOrWhiteSpace(options.ApiBasePath))
            throw new ArgumentException("An iManage API base path is required.", nameof(options));
        if (options.MaximumBytes <= 0)
            throw new ArgumentOutOfRangeException(nameof(options), "MaximumBytes must be positive.");
        if (options.AllowedExtensions is null || options.AllowedExtensions.Count == 0)
            throw new ArgumentException("At least one allowed document extension is required.", nameof(options));
        if (!string.IsNullOrWhiteSpace(options.CreationFolderId))
        {
            if (!IManageIds.TryParseContainerId(options.CreationFolderId, out var library))
                throw new ArgumentException(
                    "CreationFolderId must be an iManage folder or workspace id such as 'ACTIVE!773'.", nameof(options));
            _creationFolderId = options.CreationFolderId.Trim();
            _creationLibrary = library;
        }
        if (!string.IsNullOrWhiteSpace(options.TemplatesFolder))
        {
            if (!IManageIds.TryParseContainerId(options.TemplatesFolder, out _))
                throw new ArgumentException(
                    "TemplatesFolder must be an iManage folder or workspace id such as 'ACTIVE!773'.", nameof(options));
            _templatesFolderId = options.TemplatesFolder.Trim();
        }
        _templateLocation = CloudListing.TemplateLocationFor(_templatesFolderId is not null, options.TemplateLocation, nameof(options));

        _server = server;
        _apiRoot = server.GetLeftPart(UriPartial.Authority) + "/" + options.ApiBasePath.Trim().Trim('/') +
                   "/customers/" + Uri.EscapeDataString(options.CustomerId.Trim());
        _allowedExtensions = new HashSet<string>(
            options.AllowedExtensions.Select(NormalizeExtension),
            StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Creates an <see cref="HttpClient"/> suited to this provider: it does not follow redirects,
    /// so the <c>X-Auth-Token</c> header cannot travel to another host. Use it when you construct
    /// the provider or <see cref="IManageOAuthTokenProvider"/> yourself.
    /// </summary>
    public static HttpClient CreateHttpClient() =>
        new(new HttpClientHandler { AllowAutoRedirect = false });

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
    /// <remarks>The templates folder's id, <c>LIBRARY!number</c>.</remarks>
    public string? TemplatesFolder => _templatesFolderId;

    /// <inheritdoc />
    public TemplateLocationKind TemplateLocation => _templateLocation;

    /// <summary>
    /// Gets whether the connection refuses every save and creation; see
    /// <see cref="IManageDocumentProviderOptions.ReadOnly"/>.
    /// </summary>
    public bool ReadOnly => _options.ReadOnly;

    /// <inheritdoc />
    public async Task<DocumentReference> RegisterAsync(
        string source,
        CancellationToken cancellationToken = default)
    {
        var document = ResolveSource(source);
        var latest = await GetLatestAsync(document, itemId: null, cancellationToken).ConfigureAwait(false);

        // A registration follows the latest version, so an older one would be edited as
        // something it is not. Say so rather than silently open a different version.
        if (latest.Version != document.Version)
            throw Error(ProviderErrorCode.InvalidArgument,
                $"The source names version {document.Version}, but the document's latest version is " +
                $"{latest.Version}. OfficeAgent edits a document's latest version; register the latest " +
                "version's id or link instead.");

        ValidateExtension(latest.FileName, itemId: null);
        if (latest.Size > _options.MaximumBytes)
            throw Error(ProviderErrorCode.ContentTooLarge,
                $"Document '{latest.FileName}' exceeds the configured maximum of {_options.MaximumBytes} bytes.");

        var id = await _store.AddAsync(
            new IManageDocumentRef(latest.Library, latest.DocumentNumber, latest.Version),
            cancellationToken).ConfigureAwait(false);
        return CreateReference(id, latest);
    }

    /// <inheritdoc />
    public async Task<DocumentReference> CreateAsync(
        string name,
        Stream content,
        CancellationToken cancellationToken = default)
    {
        if (content is null) throw new ArgumentNullException(nameof(content));
        RefuseWhenReadOnly(itemId: null);
        RequireCreationFolder();

        var fileName = ValidateCreateName(name);
        var bytes = await ReadStreamAsync(content, itemId: null, cancellationToken).ConfigureAwait(false);
        Profile created;
        try
        {
            created = await CreateInFolderAsync(fileName, bytes, itemId: null, cancellationToken).ConfigureAwait(false);
        }
        catch (DocumentProviderException ex) when (ex.Code != ProviderErrorCode.AlreadyExists)
        {
            throw Error(ex.Code, SafeCreateFailureMessage(ex.Code, fileName), itemId: null, ex);
        }
        catch (Exception ex) when (ex is not OperationCanceledException and not DocumentProviderException)
        {
            // Transport and token failures can carry the request URL, which names the
            // host-configured folder. Keep those details only on the inner exception.
            throw Error(ProviderErrorCode.IO,
                $"iManage could not create document '{fileName}' in the configured folder.", itemId: null, ex);
        }

        return await RegisterStoredAsync(created, fileName, bytes, sourceItemId: null,
            $"Document '{fileName}' was created in iManage, but its registration could not be persisted. " +
            "The document is stored but unregistered. Do not retry creation; register it by its document id, " +
            "or report it to the host/operator for recovery.").ConfigureAwait(false);
    }

    /// <inheritdoc />
    public async Task<DocumentContent> OpenReadAsync(
        DocumentReference reference,
        CancellationToken cancellationToken = default)
    {
        ValidateReference(reference);
        var registered = await ResolveRegistrationAsync(reference.ItemId, cancellationToken).ConfigureAwait(false);
        var latest = await GetLatestAsync(registered, reference.ItemId, cancellationToken).ConfigureAwait(false);

        // Limits come from the profile, before a byte of content is downloaded.
        ValidateExtension(latest.FileName, reference.ItemId);
        if (latest.Size > _options.MaximumBytes)
            throw Error(ProviderErrorCode.ContentTooLarge,
                $"Document exceeds the configured maximum of {_options.MaximumBytes} bytes.", reference.ItemId);
        var current = latest.Token;
        if (!string.IsNullOrEmpty(reference.Version) &&
            !string.Equals(reference.Version, current, StringComparison.Ordinal))
            throw new DocumentVersionConflictException(reference.Version!, current, ProviderName, ConnectionId, reference.ItemId);

        var bytes = await DownloadAsync(latest, reference.ItemId, cancellationToken).ConfigureAwait(false);
        return new DocumentContent(CreateReference(reference.ItemId, latest), new MemoryStream(bytes, writable: false));
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
        RefuseWhenReadOnly(source.ItemId);
        if (!string.IsNullOrEmpty(options.DestinationItemId))
            throw Error(ProviderErrorCode.InvalidArgument,
                "iManage document ids are opaque and provider-assigned; a destination id cannot be chosen by the caller. Use Mode=NewVersion to add an iManage version, Mode=NewDocument (optionally with NewName) to file a new document, or Mode=Replace to save a new version under the same id.",
                source.ItemId);
        if (options.Mode != SaveMode.NewDocument && !string.IsNullOrWhiteSpace(options.NewName))
            throw Error(ProviderErrorCode.InvalidArgument,
                "An iManage version keeps its document's name. Use Mode=NewDocument with NewName to file a renamed document.",
                source.ItemId);
        if (options.Mode == SaveMode.NewDocument)
            RequireCreationFolder();

        var registered = await ResolveRegistrationAsync(source.ItemId, cancellationToken).ConfigureAwait(false);
        var latest = await GetLatestAsync(registered, source.ItemId, cancellationToken).ConfigureAwait(false);

        var expectedVersion = options.ExpectedVersion ?? source.Version;
        if (!string.IsNullOrEmpty(expectedVersion) &&
            !string.Equals(expectedVersion, latest.Token, StringComparison.Ordinal))
            throw new DocumentVersionConflictException(expectedVersion!, latest.Token, ProviderName, ConnectionId, source.ItemId);

        var bytes = await ReadStreamAsync(content, source.ItemId, cancellationToken).ConfigureAwait(false);

        if (options.Mode == SaveMode.NewDocument)
        {
            var destinationName = string.IsNullOrWhiteSpace(options.NewName)
                ? NextVersionedName(
                    await FolderNamesAsync(source.ItemId, cancellationToken).ConfigureAwait(false), latest.FileName, source.ItemId)
                : ValidateName(options.NewName!);
            var filed = await CreateInFolderAsync(destinationName, bytes, source.ItemId, cancellationToken).ConfigureAwait(false);
            return await RegisterStoredAsync(filed, destinationName, bytes, source.ItemId,
                $"Document '{destinationName}' was filed in iManage, but its registration could not be " +
                "persisted. Do not save it again; register the filed document to recover it.").ConfigureAwait(false);
        }

        var saved = await AddVersionAsync(registered, latest, bytes, source.ItemId, cancellationToken).ConfigureAwait(false);
        if (options.Mode == SaveMode.Replace)
            return CreateReference(source.ItemId, saved);

        return await RegisterStoredAsync(saved, saved.FileName, bytes, source.ItemId,
            $"Version {saved.Version} of '{saved.FileName}' was saved to iManage, but its registration could " +
            "not be persisted. Do not save it again; register the document to recover it.").ConfigureAwait(false);
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
    /// <c>GET /libraries/{library}/folders/{id}/children</c> and <c>limit</c>/<c>offset</c> paging.
    /// A connection without that folder reports <see cref="ProviderErrorCode.ConfigurationError"/>,
    /// or an empty page for templates. <see cref="DocumentListQuery.Folder"/> is a path of folder
    /// names below it, found by reading each level's children. Each entry's source is the listed
    /// version's id, <c>LIBRARY!number.version</c>. Search folders are not descended into, so a
    /// listing cannot reach documents filed elsewhere. A page makes at most
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
                "This iManage connection has no folder to list. Configure CreationFolderId, or TemplatesFolder for templates.");
        }

        try
        {
            var calls = CloudListing.MaximumServiceCalls;
            CloudListingCursor cursor;
            if (string.IsNullOrEmpty(query.ContinuationToken))
            {
                var folderId = rootId;
                foreach (var segment in request.FolderSegments)
                {
                    string? found = null;
                    string? offset = null;
                    do
                    {
                        // At least one request must be left for the listing itself.
                        if (calls <= 1) throw Error(ProviderErrorCode.InvalidArgument, CloudListing.TooDeepMessage);
                        var page = await ListChildrenAsync(folderId, offset, cancellationToken).ConfigureAwait(false);
                        calls--;
                        found = page.Items.FirstOrDefault(item =>
                            item.IsFolder && string.Equals(item.Name, segment, StringComparison.OrdinalIgnoreCase))?.Id;
                        offset = page.NextToken;
                    }
                    while (found is null && offset is not null);
                    folderId = found ?? throw Error(ProviderErrorCode.NotFound,
                        $"No folder named '{segment}' is in the listed iManage folder.");
                }
                cursor = CloudListing.Start(folderId, request);
            }
            else if (!CloudListingCursor.TryUnprotect(query.ContinuationToken!, ConnectionId, request.Fingerprint, out cursor))
            {
                throw Error(ProviderErrorCode.InvalidArgument, CloudListing.InvalidTokenMessage);
            }

            return await CloudListing.WalkAsync(
                request, cursor, ListChildrenAsync, calls, ConnectionId, cancellationToken).ConfigureAwait(false);
        }
        catch (Exception ex) when (ex is not OperationCanceledException and not DocumentProviderException)
        {
            // Transport and token failures can carry request URLs, which name the folder.
            throw Error(ProviderErrorCode.IO, "iManage could not be reached to list the folder.", itemId: null, ex);
        }
    }

    // ── Sources ───────────────────────────────────────────────────────────────

    private IManageDocumentRef ResolveSource(string source)
    {
        if (string.IsNullOrWhiteSpace(source))
            throw Error(ProviderErrorCode.InvalidArgument,
                "An iManage registration requires a source: a document id such as 'ACTIVE!12345.1' " +
                "or an iManage Work link to the document.");

        var trimmed = source.Trim();
        if (trimmed.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
            trimmed.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
        {
            if (!Uri.TryCreate(trimmed, UriKind.Absolute, out var url))
                throw Error(ProviderErrorCode.InvalidArgument, "The iManage link is not a valid URL.");
            if (url.Scheme != Uri.UriSchemeHttps ||
                !string.Equals(url.Host, _server.Host, StringComparison.OrdinalIgnoreCase))
                throw Error(ProviderErrorCode.InvalidArgument,
                    "The link is not on this connection's iManage server; register a document from the server the connection is configured for.");
            if (!IManageIds.TryParseDocumentUrl(url, out var linked))
                throw Error(ProviderErrorCode.InvalidArgument,
                    "The iManage link does not name a document version. Copy the document's link, or register it by its id such as 'ACTIVE!12345.1'.");
            return linked;
        }

        if (IManageIds.TryParseDocumentId(trimmed, out var document))
            return document;
        if (IManageIds.LooksLikeContainerId(trimmed))
            throw Error(ProviderErrorCode.InvalidArgument,
                "That id names a folder or workspace, or a document without its version. Register a document by " +
                "its full id, 'LIBRARY!number.version', such as 'ACTIVE!12345.1'.");
        throw Error(ProviderErrorCode.InvalidArgument,
            "Register an iManage document by its id, 'LIBRARY!number.version' such as 'ACTIVE!12345.1', " +
            "or by an iManage Work link to it, not by a name or path.");
    }

    // ── iManage plumbing ──────────────────────────────────────────────────────

    private string LibraryUrl(string library) => $"{_apiRoot}/libraries/{library}";

    private string DocumentUrl(string library, string documentId) => $"{LibraryUrl(library)}/documents/{documentId}";

    /// <summary>
    /// Reads every version's profile and returns the latest. Listing versions, rather than
    /// asking for "the latest", keeps the lookup working from any version the registration knows.
    /// </summary>
    private async Task<Profile> GetLatestAsync(IManageDocumentRef document, string? itemId, CancellationToken cancellationToken)
    {
        using var response = await SendAsync(
            HttpMethod.Get, $"{DocumentUrl(document.Library, document.DocumentId)}/versions",
            content: null, HttpCompletionOption.ResponseContentRead, cancellationToken).ConfigureAwait(false);
        await ThrowOnFailureAsync(response, itemId, write: false).ConfigureAwait(false);
        var payload = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

        Profile? latest = null;
        foreach (var profile in ParseProfiles(payload, document.Library, document.DocumentNumber, itemId))
            if (latest is null || profile.Version > latest.Version)
                latest = profile;
        return latest ?? throw Error(ProviderErrorCode.NotFound,
            "The referenced iManage document has no version this identity can read.", itemId);
    }

    /// <summary>
    /// Adds a new version to the document. With <see cref="IManageDocumentProviderOptions.LockDuringSave"/>
    /// the document is locked first and the latest version read again under the lock, so a version
    /// another writer added after the caller's read is a conflict, not silently superseded.
    /// </summary>
    private async Task<Profile> AddVersionAsync(
        IManageDocumentRef registered, Profile latest, byte[] bytes, string itemId, CancellationToken cancellationToken)
    {
        if (latest.IsCheckedOut)
            throw Locked(itemId, code: string.Empty);

        var locked = false;
        try
        {
            if (_options.LockDuringSave)
            {
                await LockAsync(latest, itemId, cancellationToken).ConfigureAwait(false);
                locked = true;

                // Under the lock, compare the version number and size. Whether taking the lock
                // touches edit_date is not documented, so it is left out of this second check;
                // the caller's opened version was already compared in full.
                var current = await GetLatestAsync(registered, itemId, cancellationToken).ConfigureAwait(false);
                if (current.Version != latest.Version || current.Size != latest.Size)
                    throw new DocumentVersionConflictException(latest.Token, current.Token, ProviderName, ConnectionId, itemId);
            }

            var profile = ProfileJson(new Dictionary<string, object?>
            {
                ["name"] = latest.Name,
                ["extension"] = latest.Extension,
                ["size"] = bytes.LongLength,
                ["comment"] = NullIfEmpty(_options.SaveComment)
            });
            return await UploadAsync(
                $"{DocumentUrl(latest.Library, latest.Id)}/versions",
                profile, bytes, latest.FileName, latest.Library, latest.DocumentNumber, itemId,
                cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            // The lock is released whatever happened; a lock left behind would block the
            // document's owner. A failure to release it never hides the save's own outcome.
            if (locked)
                await UnlockQuietlyAsync(latest).ConfigureAwait(false);
        }
    }

    private async Task LockAsync(Profile latest, string itemId, CancellationToken cancellationToken)
    {
        using var body = new StringContent(
            JsonSerializer.Serialize(new Dictionary<string, object?> { ["comments"] = "OfficeAgent is saving a new version." }),
            Encoding.UTF8, "application/json");
        HttpResponseMessage response;
        try
        {
            response = await SendAsync(
                HttpMethod.Post, $"{DocumentUrl(latest.Library, latest.Id)}/lock",
                body, HttpCompletionOption.ResponseContentRead, cancellationToken).ConfigureAwait(false);
        }
        catch (Exception ex) when (ex is not OperationCanceledException and not DocumentProviderException)
        {
            await UnlockQuietlyAsync(latest).ConfigureAwait(false);
            throw Error(ProviderErrorCode.WriteRejected,
                "iManage could not be reached to lock the document for saving; nothing was stored.", itemId, ex);
        }

        using (response)
        {
            if (response.IsSuccessStatusCode) return;
            var code = await ReadErrorCodeAsync(response).ConfigureAwait(false);
            var status = (int)response.StatusCode;
            if (IsLocked(response.StatusCode, code))
                throw Locked(itemId, code);
            if (status is 401 or 403)
                throw Error(ProviderErrorCode.AccessDenied,
                    $"iManage refused to lock the document for this identity; nothing was stored.{Suffix(code)}", itemId);
            if (status is 404 or 405 or 501)
                throw Error(ProviderErrorCode.ConfigurationError,
                    "This iManage server did not accept a document lock request; nothing was stored. If the server " +
                    $"does not support document locks, set LockDuringSave to false for this connection.{Suffix(code)}", itemId);
            if (status >= 500)
                await UnlockQuietlyAsync(latest).ConfigureAwait(false);
            throw Error(ProviderErrorCode.WriteRejected,
                $"iManage could not lock the document for saving (status {status}); nothing was stored.{Suffix(code)}", itemId);
        }
    }

    private async Task UnlockQuietlyAsync(Profile latest)
    {
        try
        {
            using var response = await SendAsync(
                HttpMethod.Delete, $"{DocumentUrl(latest.Library, latest.Id)}/lock",
                content: null, HttpCompletionOption.ResponseContentRead, CancellationToken.None).ConfigureAwait(false);
        }
        catch (Exception)
        {
            // Best effort: an upload may already have released the lock.
        }
    }

    private async Task<Profile> CreateInFolderAsync(string fileName, byte[] bytes, string? itemId, CancellationToken cancellationToken)
    {
        // iManage gives every new document a new number, so creation can never overwrite. The
        // name check keeps OfficeAgent's rule that a taken name is refused, not duplicated.
        var names = await FolderNamesAsync(itemId, cancellationToken).ConfigureAwait(false);
        if (names.Contains(fileName))
            throw Error(ProviderErrorCode.AlreadyExists,
                "A document with that name already exists in the configured iManage folder; choose another name.", itemId);

        var profile = ProfileJson(new Dictionary<string, object?>
        {
            ["name"] = Path.GetFileNameWithoutExtension(fileName),
            ["extension"] = Path.GetExtension(fileName).TrimStart('.'),
            ["size"] = bytes.LongLength,
            ["comment"] = NullIfEmpty(_options.SaveComment),
            ["class"] = NullIfEmpty(_options.CreationDocumentClass)
        });
        return await UploadAsync(
            $"{LibraryUrl(_creationLibrary!)}/folders/{_creationFolderId}/documents",
            profile, bytes, fileName, _creationLibrary!, documentNumber: 0, itemId, cancellationToken).ConfigureAwait(false);
    }

    private async Task<HashSet<string>> FolderNamesAsync(string? itemId, CancellationToken cancellationToken)
    {
        using var response = await SendAsync(
            HttpMethod.Get, $"{LibraryUrl(_creationLibrary!)}/folders/{_creationFolderId}/children",
            content: null, HttpCompletionOption.ResponseContentRead, cancellationToken).ConfigureAwait(false);
        await ThrowOnFailureAsync(response, itemId, write: false).ConfigureAwait(false);
        var payload = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var child in ParseProfiles(payload, _creationLibrary!, 0, itemId, documentsOnly: true))
            names.Add(child.FileName);
        return names;
    }

    /// <summary>
    /// One page of a folder's children with <c>limit</c>/<c>offset</c> paging, subfolders included,
    /// so a recursive listing can descend. A full page means another may follow; the offset is the
    /// page token.
    /// </summary>
    private async Task<CloudListingPage> ListChildrenAsync(string folderId, string? offset, CancellationToken cancellationToken)
    {
        if (!IManageIds.TryParseContainerId(folderId, out var library))
            throw Error(ProviderErrorCode.IO, "iManage listed a folder whose id is not a folder id.");
        var start = int.TryParse(offset, System.Globalization.NumberStyles.None,
            System.Globalization.CultureInfo.InvariantCulture, out var parsed) ? parsed : 0;
        var limit = CloudListing.ServicePageSize.ToString(System.Globalization.CultureInfo.InvariantCulture);
        using var response = await SendAsync(
            HttpMethod.Get,
            $"{LibraryUrl(library)}/folders/{folderId.Trim()}/children?limit={limit}&offset={start.ToString(System.Globalization.CultureInfo.InvariantCulture)}",
            content: null, HttpCompletionOption.ResponseContentRead, cancellationToken).ConfigureAwait(false);
        await ThrowOnFailureAsync(response, itemId: null, write: false).ConfigureAwait(false);
        var payload = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

        var children = new List<CloudListingChild>();
        int count;
        try
        {
            using var json = JsonDocument.Parse(payload);
            var root = json.RootElement;
            var data = root.ValueKind == JsonValueKind.Object && root.TryGetProperty("data", out var d) ? d : root;
            if (data.ValueKind == JsonValueKind.Object && data.TryGetProperty("results", out var results))
                data = results;
            if (data.ValueKind != JsonValueKind.Array)
                return new CloudListingPage();

            count = data.GetArrayLength();
            foreach (var element in data.EnumerateArray())
            {
                if (element.ValueKind != JsonValueKind.Object) continue;
                var wstype = ReadString(element, "wstype") ?? string.Empty;
                if (wstype.Equals("folder", StringComparison.OrdinalIgnoreCase) ||
                    wstype.Equals("tab", StringComparison.OrdinalIgnoreCase))
                {
                    // A search folder's children are query results from anywhere in the library,
                    // so it is not descended into.
                    var id = ReadString(element, "id") ?? string.Empty;
                    var search = (ReadString(element, "folder_type") ?? string.Empty)
                        .IndexOf("search", StringComparison.OrdinalIgnoreCase) >= 0;
                    if (IManageIds.TryParseContainerId(id, out _) && !search)
                        children.Add(new CloudListingChild { Id = id.Trim(), Name = ReadString(element, "name") ?? string.Empty, IsFolder = true });
                    continue;
                }
                if (wstype.Length > 0 && !wstype.Equals("document", StringComparison.OrdinalIgnoreCase))
                    continue;

                var profiles = new List<Profile>();
                AddProfile(profiles, element, library, 0, documentsOnly: true);
                if (profiles.Count == 0) continue;
                var profile = profiles[0];
                var valid = IManageIds.IsLibrary(profile.Library) && profile.DocumentNumber > 0 && profile.Version > 0;
                children.Add(new CloudListingChild
                {
                    Id = valid ? $"{profile.Library}!{profile.DocumentNumber}.{profile.Version}" : string.Empty,
                    Name = profile.FileName,
                    Unusable = !valid || profile.Size > _options.MaximumBytes,
                    Length = profile.Size,
                    LastModified = DateTimeOffset.TryParse(profile.EditDate, System.Globalization.CultureInfo.InvariantCulture,
                        System.Globalization.DateTimeStyles.AssumeUniversal, out var edited)
                        ? edited
                        : null
                });
            }
        }
        catch (Exception ex) when (ex is JsonException or InvalidOperationException or FormatException)
        {
            throw Error(ProviderErrorCode.IO, "iManage returned an unreadable folder listing.", itemId: null, ex);
        }

        // The page size is what was asked for, so a short page is the last one.
        return new CloudListingPage
        {
            Items = children,
            NextToken = count >= CloudListing.ServicePageSize
                ? (start + count).ToString(System.Globalization.CultureInfo.InvariantCulture)
                : null
        };
    }

    private async Task<Profile> UploadAsync(
        string url,
        string profileJson,
        byte[] bytes,
        string fileName,
        string library,
        long documentNumber,
        string? itemId,
        CancellationToken cancellationToken)
    {
        using var body = new MultipartFormDataContent();
        var profile = new StringContent(profileJson, Encoding.UTF8);
        profile.Headers.ContentType = new MediaTypeHeaderValue("application/json");
        body.Add(profile, "profile");
        var file = new ByteArrayContent(bytes);
        file.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
        body.Add(file, "file", fileName);

        using var response = await SendAsync(
            HttpMethod.Post, url, body, HttpCompletionOption.ResponseContentRead, cancellationToken).ConfigureAwait(false);
        await ThrowOnFailureAsync(response, itemId, write: true).ConfigureAwait(false);
        var payload = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
        var saved = ParseProfiles(payload, library, documentNumber, itemId).FirstOrDefault();
        if (saved is null || saved.Version <= 0 || saved.DocumentNumber <= 0)
            throw Error(ProviderErrorCode.IO,
                "iManage accepted the upload but did not describe the stored document; it may have been stored.", itemId);
        return saved;
    }

    private async Task<byte[]> DownloadAsync(Profile latest, string itemId, CancellationToken cancellationToken)
    {
        var response = await SendAsync(
            HttpMethod.Get, $"{DocumentUrl(latest.Library, latest.Id)}/download",
            content: null, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(false);
        try
        {
            if (IsRedirect(response.StatusCode) && response.Headers.Location is { } location)
            {
                // A storage redirect is fetched without the token, like a pre-authenticated
                // download URL; only a hop that stays on the iManage server keeps it.
                var target = location.IsAbsoluteUri ? location : new Uri(response.RequestMessage!.RequestUri!, location);
                response.Dispose();
                if (target.Scheme != Uri.UriSchemeHttps)
                    throw Error(ProviderErrorCode.IO, "iManage redirected the download to an insecure address.", itemId);
                response = SameServer(target)
                    ? await SendAsync(HttpMethod.Get, target.AbsoluteUri, null, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(false)
                    : await _http.GetAsync(target, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(false);
            }

            if (!response.IsSuccessStatusCode)
            {
                await ThrowOnFailureAsync(response, itemId, write: false).ConfigureAwait(false);
            }
            if (response.Content.Headers.ContentLength is { } length && length > _options.MaximumBytes)
                throw Error(ProviderErrorCode.ContentTooLarge,
                    $"Document exceeds the configured maximum of {_options.MaximumBytes} bytes.", itemId);

            using var stream = await response.Content.ReadAsStreamAsync().ConfigureAwait(false);
            return await ReadBoundedAsync(stream, itemId, cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            response.Dispose();
        }
    }

    private async Task<HttpResponseMessage> SendAsync(
        HttpMethod method, string url, HttpContent? content, HttpCompletionOption completion, CancellationToken cancellationToken)
    {
        var token = await _tokens.GetAccessTokenAsync(cancellationToken).ConfigureAwait(false);
        using var request = new HttpRequestMessage(method, url);
        request.Headers.TryAddWithoutValidation(AuthHeader, token);
        if (content is not null)
            request.Content = content;
        var response = await _http.SendAsync(request, completion, cancellationToken).ConfigureAwait(false);

        // An HttpClient that follows redirects would already have sent the token onward. That
        // cannot be undone here, but the call fails loudly instead of using the answer.
        var answered = response.RequestMessage?.RequestUri;
        if (answered is not null && !SameServer(answered))
        {
            response.Dispose();
            throw Error(ProviderErrorCode.IO,
                "An iManage request was redirected off the configured server. Give the provider an HttpClient " +
                "that does not follow redirects, such as IManageDocumentProvider.CreateHttpClient().");
        }
        return response;
    }

    private bool SameServer(Uri url) =>
        string.Equals(url.Host, _server.Host, StringComparison.OrdinalIgnoreCase) && url.Port == _server.Port;

    private static bool IsRedirect(HttpStatusCode status) => (int)status is >= 300 and < 400;

    private async Task ThrowOnFailureAsync(HttpResponseMessage response, string? itemId, bool write)
    {
        if (response.IsSuccessStatusCode) return;

        var code = await ReadErrorCodeAsync(response).ConfigureAwait(false);
        var suffix = Suffix(code);
        var status = (int)response.StatusCode;
        throw response.StatusCode switch
        {
            // Retrying cannot help a checked-out document, so the message says what can.
            _ when write && IsLocked(response.StatusCode, code) => Locked(itemId, code),
            HttpStatusCode.NotFound => Error(ProviderErrorCode.NotFound,
                $"The referenced iManage document or folder no longer exists, or this identity cannot see it.{suffix}", itemId),
            HttpStatusCode.Unauthorized or HttpStatusCode.Forbidden => Error(ProviderErrorCode.AccessDenied,
                $"iManage refused access.{suffix}", itemId),
            HttpStatusCode.PreconditionFailed => new DocumentVersionConflictException(
                "(opened version)", "(changed)", ProviderName, ConnectionId, itemId),
            HttpStatusCode.RequestEntityTooLarge => Error(ProviderErrorCode.ContentTooLarge,
                $"iManage rejected the content as too large.{suffix}", itemId),
            // For an upload, any other refusal status (400, 409, 429, ...) is iManage declining the
            // request: nothing was stored. A server error is not a refusal; the request may have
            // been processed, so the outcome is unknown.
            _ when write && status is >= 400 and < 500 => Error(ProviderErrorCode.WriteRejected,
                $"iManage refused the upload with status {status}; nothing was stored.{suffix}", itemId),
            _ when write => Error(ProviderErrorCode.OutcomeUnknown,
                $"iManage failed the upload with status {status}; it may have been stored.{suffix}", itemId),
            _ => Error(ProviderErrorCode.IO, $"iManage request failed with status {status}.{suffix}", itemId)
        };
    }

    private DocumentProviderException Locked(string? itemId, string code) => Error(
        ProviderErrorCode.WriteRejected,
        "The document is checked out or locked in iManage, for example open for editing by another user; " +
        "nothing was stored. Ask whoever holds it to check it in, or save as a new document instead." + Suffix(code),
        itemId);

    // iManage's error codes are not published; these match the words lock and check-out states use.
    private static bool IsLocked(HttpStatusCode status, string code) =>
        (int)status == 423 ||
        code.IndexOf("LOCK", StringComparison.OrdinalIgnoreCase) >= 0 ||
        code.IndexOf("CHECKED_OUT", StringComparison.OrdinalIgnoreCase) >= 0 ||
        code.IndexOf("CHECKEDOUT", StringComparison.OrdinalIgnoreCase) >= 0 ||
        code.IndexOf("IN_USE", StringComparison.OrdinalIgnoreCase) >= 0;

    private static string Suffix(string code) =>
        string.IsNullOrEmpty(code) ? string.Empty : $" iManage code: {code}.";

    /// <summary>
    /// Reads an error code from an iManage error body. Only a short token is kept: messages can
    /// name libraries, users, or paths, so they never reach a result.
    /// </summary>
    private static async Task<string> ReadErrorCodeAsync(HttpResponseMessage response)
    {
        try
        {
            var payload = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            if (string.IsNullOrWhiteSpace(payload)) return string.Empty;
            using var json = JsonDocument.Parse(payload);
            var root = json.RootElement;
            if (root.ValueKind != JsonValueKind.Object) return string.Empty;
            string? code = null;
            if (root.TryGetProperty("error", out var error))
            {
                if (error.ValueKind == JsonValueKind.Object && error.TryGetProperty("code", out var nested))
                    code = Scalar(nested);
                else if (error.ValueKind == JsonValueKind.String)
                    code = error.GetString();
            }
            if (code is null && root.TryGetProperty("code", out var top))
                code = Scalar(top);
            return SafeCode(code);
        }
        catch (JsonException)
        {
            return string.Empty;
        }
    }

    private static string? Scalar(JsonElement element) => element.ValueKind switch
    {
        JsonValueKind.String => element.GetString(),
        JsonValueKind.Number => element.GetRawText(),
        _ => null
    };

    private static string SafeCode(string? code)
    {
        if (string.IsNullOrEmpty(code) || code!.Length > 64) return string.Empty;
        foreach (var c in code)
            if (!(char.IsLetterOrDigit(c) || c is '_' or '-' or '.'))
                return string.Empty;
        return code;
    }

    private static string ProfileJson(Dictionary<string, object?> fields)
    {
        var profile = fields.Where(pair => pair.Value is not null).ToDictionary(pair => pair.Key, pair => pair.Value);
        return JsonSerializer.Serialize(new Dictionary<string, object?>
        {
            ["doc_profile"] = profile,
            ["warnings_for_required_and_disabled_fields"] = true
        });
    }

    private static string? NullIfEmpty(string? value) => string.IsNullOrWhiteSpace(value) ? null : value!.Trim();

    /// <summary>
    /// Reads document profiles from a response: one profile under <c>data</c>, a list under
    /// <c>data</c>, or a list under <c>data.results</c>.
    /// </summary>
    private List<Profile> ParseProfiles(string payload, string library, long documentNumber, string? itemId, bool documentsOnly = false)
    {
        try
        {
            using var json = JsonDocument.Parse(payload);
            var root = json.RootElement;
            var data = root.ValueKind == JsonValueKind.Object && root.TryGetProperty("data", out var d) ? d : root;
            if (data.ValueKind == JsonValueKind.Object && data.TryGetProperty("results", out var results))
                data = results;

            var profiles = new List<Profile>();
            if (data.ValueKind == JsonValueKind.Array)
            {
                foreach (var element in data.EnumerateArray())
                    if (element.ValueKind == JsonValueKind.Object)
                        AddProfile(profiles, element, library, documentNumber, documentsOnly);
            }
            else if (data.ValueKind == JsonValueKind.Object)
            {
                AddProfile(profiles, data, library, documentNumber, documentsOnly);
            }
            return profiles;
        }
        catch (Exception ex) when (ex is JsonException or InvalidOperationException or FormatException)
        {
            throw Error(ProviderErrorCode.IO, "iManage returned an unreadable document profile.", itemId, ex);
        }
    }

    private static void AddProfile(List<Profile> profiles, JsonElement element, string library, long documentNumber, bool documentsOnly)
    {
        var wstype = ReadString(element, "wstype");
        if (documentsOnly && wstype is not null && !string.Equals(wstype, "document", StringComparison.OrdinalIgnoreCase))
            return;

        var id = ReadString(element, "id") ?? string.Empty;
        IManageIds.TryParseDocumentId(id, out var parsed);
        var profileLibrary = ReadString(element, "database") ?? (string.IsNullOrEmpty(parsed.Library) ? library : parsed.Library);
        var number = ReadInt64(element, "document_number") ?? (parsed.DocumentNumber > 0 ? parsed.DocumentNumber : documentNumber);
        var version = (int)(ReadInt64(element, "version") ?? parsed.Version);
        if (!IManageIds.IsLibrary(profileLibrary))
            profileLibrary = library;
        if (string.IsNullOrEmpty(id) || parsed.DocumentNumber <= 0)
            id = $"{profileLibrary}!{number}.{version}";

        profiles.Add(new Profile
        {
            Id = id,
            Library = profileLibrary,
            DocumentNumber = number,
            Version = version,
            Name = ReadString(element, "name") ?? string.Empty,
            Extension = (ReadString(element, "extension") ?? string.Empty).Trim().TrimStart('.'),
            FullFileName = ReadString(element, "full_file_name"),
            Size = ReadInt64(element, "size") ?? 0,
            EditDate = ReadString(element, "edit_date") ?? string.Empty,
            IsCheckedOut = element.TryGetProperty("is_checked_out", out var checkedOut) &&
                           checkedOut.ValueKind == JsonValueKind.True
        });
    }

    private static string? ReadString(JsonElement element, string name) =>
        element.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String ? value.GetString() : null;

    private static long? ReadInt64(JsonElement element, string name)
    {
        if (!element.TryGetProperty(name, out var value)) return null;
        if (value.ValueKind == JsonValueKind.Number && value.TryGetInt64(out var number)) return number;
        if (value.ValueKind == JsonValueKind.String &&
            long.TryParse(value.GetString(), System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out var parsed))
            return parsed;
        return null;
    }

    // ── Registration & validation ─────────────────────────────────────────────

    private async Task<DocumentReference> RegisterStoredAsync(
        Profile stored, string outputName, byte[] bytes, string? sourceItemId, string failureMessage)
    {
        string id;
        try
        {
            // iManage confirmed the upload. Finish registering even if the caller cancels now,
            // so a stored document is not left unregistered behind a plain cancellation.
            id = await _store.AddAsync(
                new IManageDocumentRef(stored.Library, stored.DocumentNumber, stored.Version),
                CancellationToken.None).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            throw RegistrationFailed(failureMessage, sourceItemId, outputName, bytes, ex);
        }
        return CreateReference(id, stored);
    }

    private async Task<IManageDocumentRef> ResolveRegistrationAsync(string registrationId, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(registrationId))
            throw Error(ProviderErrorCode.InvalidArgument, "An item id is required.");
        return await _store.ResolveAsync(registrationId, cancellationToken).ConfigureAwait(false)
            ?? throw Error(ProviderErrorCode.NotFound,
                $"The provider document '{registrationId}' is not registered.", registrationId);
    }

    private void RequireCreationFolder()
    {
        if (_creationFolderId is null)
            throw Error(ProviderErrorCode.ConfigurationError,
                "This iManage connection has no creation folder. Configure CreationFolderId, or save as a new " +
                "version of a registered document instead.");
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
        if (Path.GetFileNameWithoutExtension(name).Trim().Length == 0)
            throw Error(ProviderErrorCode.InvalidArgument, "A document name needs text before its extension.");
        ValidateExtension(name, itemId: null);
        return name;
    }

    private string ValidateCreateName(string name)
    {
        if (!string.IsNullOrWhiteSpace(name) && PortableDocumentName.ContainsInvalidCharacter(name))
            throw Error(ProviderErrorCode.InvalidArgument,
                "A document name must be a bare file name without path separators, " +
                "control characters, or other invalid filename characters.");
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
            $"The configured iManage creation folder for '{fileName}' was not found.",
        ProviderErrorCode.AccessDenied =>
            $"The configured iManage identity cannot file '{fileName}' in the creation folder.",
        ProviderErrorCode.ContentTooLarge =>
            $"iManage rejected '{fileName}' because its content is too large.",
        ProviderErrorCode.WriteRejected =>
            $"iManage refused to file '{fileName}' in the configured folder; nothing was stored.",
        ProviderErrorCode.OutcomeUnknown =>
            $"iManage failed while filing '{fileName}'; it may have been stored.",
        _ => $"iManage could not create document '{fileName}' in the configured folder."
    };

    private void ValidateReference(DocumentReference reference)
    {
        if (reference is null) throw new ArgumentNullException(nameof(reference));
        if (!string.Equals(reference.Provider, ProviderName, StringComparison.OrdinalIgnoreCase) ||
            !string.Equals(reference.ConnectionId, ConnectionId, StringComparison.Ordinal))
            throw new DocumentProviderException(
                ProviderErrorCode.InvalidArgument,
                $"The document reference (provider='{reference.Provider}', connectionId='{reference.ConnectionId}') does not belong to this iManage connection (provider='{ProviderName}', connectionId='{ConnectionId}').",
                ProviderName, ConnectionId, reference.ItemId);
    }

    /// <summary>
    /// Picks the next unused versioned name in the creation folder, matching the other
    /// providers' convention: <c>contract.docx</c> → <c>contract.v2.docx</c>, ….
    /// </summary>
    private string NextVersionedName(HashSet<string> taken, string sourceName, string itemId)
    {
        var extension = Path.GetExtension(sourceName);
        var stem = Path.GetFileNameWithoutExtension(sourceName);
        var (baseStem, startVersion) = ParseVersionedStem(stem);
        for (var n = Math.Max(2, startVersion + 1); n < startVersion + 1000; n++)
        {
            var candidate = $"{baseStem}.v{n}{extension}";
            if (!taken.Contains(candidate))
                return candidate;
        }
        throw Error(ProviderErrorCode.IO, "Exhausted version suffixes for the source document.", itemId);
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

    private DocumentReference CreateReference(string registrationId, Profile profile) => new()
    {
        Provider = ProviderName,
        ConnectionId = ConnectionId,
        ItemId = registrationId,
        Version = profile.Token,
        Name = profile.FileName,
        ContentType = OfficeContentTypes.ForName(profile.FileName)
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
                throw Error(ProviderErrorCode.ContentTooLarge,
                    $"Document exceeds the configured maximum of {_options.MaximumBytes} bytes.", itemId);
        }
        return copy.ToArray();
    }

    /// <summary>
    /// iManage confirmed the upload but the registration store failed: carry the document's name
    /// and content hash so it can be found and registered, never its library or number.
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

    /// <summary>The refusal every save and creation on a read-only connection meets before the service is called.</summary>
    private void RefuseWhenReadOnly(string? itemId)
    {
        if (_options.ReadOnly)
            throw Error(ProviderErrorCode.AccessDenied, ReadOnlyConnections.WriteRefused, itemId);
    }

    private sealed class Profile
    {
        public string Id { get; init; } = string.Empty;
        public string Library { get; init; } = string.Empty;
        public long DocumentNumber { get; init; }
        public int Version { get; init; }
        public string Name { get; init; } = string.Empty;
        public string Extension { get; init; } = string.Empty;
        public string? FullFileName { get; init; }
        public long Size { get; init; }
        public string EditDate { get; init; } = string.Empty;
        public bool IsCheckedOut { get; init; }

        /// <summary>The display file name: the profile name with its extension.</summary>
        public string FileName
        {
            get
            {
                var extension = Extension.ToLowerInvariant();
                if (extension.Length == 0)
                    return string.IsNullOrEmpty(FullFileName) ? Name : FullFileName!;
                return Name.EndsWith("." + extension, StringComparison.OrdinalIgnoreCase)
                    ? Name
                    : $"{Name}.{extension}";
            }
        }

        /// <summary>
        /// The opaque version token: the version number, its last edit time, and its size. A new
        /// version and an edit of the latest one in place both change it.
        /// </summary>
        public string Token => $"v{Version}|{EditDate}|{Size}";
    }
}
