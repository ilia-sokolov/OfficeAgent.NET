using System.Net;
using System.Net.Http.Headers;
using System.Security.Cryptography;
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

    /// <summary>
    /// Gets or sets the size in bytes from which a save is uploaded in chunks through a Graph
    /// upload session instead of in one request. Defaults to 4 MiB, Microsoft's recommendation.
    /// It must be positive and at most 250 MB, the limit of a single-request upload.
    /// </summary>
    public long UploadSessionThresholdBytes { get; set; } = 4 * 1024 * 1024;

    /// <summary>
    /// Gets or sets the size in bytes of each chunk of an upload session. It must be a multiple of
    /// 320 KiB (327,680 bytes) and less than 60 MiB, as Graph requires. Defaults to 10 MiB,
    /// Microsoft's recommendation for a stable connection; a slow one may do better with less.
    /// </summary>
    public int UploadSessionChunkBytes { get; set; } = 10 * 1024 * 1024;

    /// <summary>Gets or sets the allowed document extensions.</summary>
    public IReadOnlyCollection<string> AllowedExtensions { get; set; } = new[] { ".docx" };

    /// <summary>Gets or sets the Graph drive id used as the destination for new documents.</summary>
    public string CreationDriveId { get; set; } = string.Empty;

    /// <summary>Gets or sets the destination folder's Graph drive-item id for new documents.</summary>
    public string CreationFolderItemId { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the change mode used when an operation does not state one. The
    /// default is <see cref="ChangeMode.Tracked"/>, which suits documents a person
    /// reviews. Set <see cref="ChangeMode.Direct"/> for a library of generated or
    /// machine-owned files, or one serving decks - PresentationML refuses Tracked.
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
    /// formulas calculated, so the saved file carries current values instead of none or stale
    /// ones. Takes effect only when a formula evaluator is registered
    /// (<c>AddExcelFormulaCalculation</c>); without one, nothing is calculated. Defaults to
    /// <see langword="false"/>.
    /// </summary>
    public bool CalculateFormulas { get; set; }

    /// <summary>
    /// Gets or sets whether a plan whose calculation produces a new formula error (such as
    /// <c>#DIV/0!</c> or <c>#REF!</c>) or a circular reference is refused with
    /// <c>formula-error</c> instead of committed with the errors reported. Applies only with
    /// <see cref="CalculateFormulas"/>. Defaults to <see langword="false"/>.
    /// </summary>
    public bool RejectFormulaErrors { get; set; }

    /// <summary>
    /// Gets or sets the folder that holds this connection's templates: a SharePoint/OneDrive
    /// folder URL or a <c>driveId/itemId</c> pair naming the folder. Listing with
    /// <see cref="DocumentListQuery.Templates"/> lists it, and every document in it is reported as
    /// a template. Empty for none. An organization template library (the SharePoint organization
    /// assets library) is not discovered automatically; name its folder here.
    /// </summary>
    public string? TemplatesFolder { get; set; }

    /// <summary>
    /// Gets or sets what kind of template location <see cref="TemplatesFolder"/> is. Left at
    /// <see cref="TemplateLocationKind.None"/>, it becomes <see cref="TemplateLocationKind.Connection"/>
    /// when a templates folder is set. Set <see cref="TemplateLocationKind.Organization"/> for an
    /// organization template library.
    /// </summary>
    public TemplateLocationKind TemplateLocation { get; set; }
}

/// <summary>
/// SharePoint document provider over the Microsoft Graph drive-item REST endpoints:
/// a registry of references to drive items, addressed by an opaque item id assigned
/// at registration time. A document is registered by a SharePoint/OneDrive URL or by
/// a <c>driveId/itemId</c> pair, so the connection is not pinned to a single drive -
/// any drive the configured identity can reach is registrable. The provider stores
/// only the registration id → (drive id, item id) mapping (via
/// <see cref="ISharePointRegistrationStore"/>) and routes open/save back to Graph.
/// When a creation drive and folder are configured, new documents are uploaded there
/// without changing the cross-drive registration boundary.
/// The agent only sees opaque ids - never site URLs, drive ids, or tokens. Saves use
/// the drive item's ETag for optimistic concurrency, and removing a registration
/// never deletes SharePoint content. A save at or above
/// <see cref="SharePointDocumentProviderOptions.UploadSessionThresholdBytes"/> is uploaded in
/// chunks through a Graph upload session, under the same preconditions.
/// </summary>
/// <remarks>
/// <see cref="ListAsync"/> walks a folder's children through Graph and filters names on this
/// side; it does not use Graph search, whose index lags behind new files. A recursive listing
/// descends at most <see cref="MaximumListingDepth"/> folders, queues at most
/// <see cref="MaximumListingFolders"/> folders in one walk, and sends at most
/// <see cref="MaximumListingRequests"/> Graph requests per call; a page cut short by the request
/// bound carries a continuation token.
/// </remarks>
public sealed class SharePointDocumentProvider : IDocumentProvider, IDocumentCreatingProvider, IDocumentListingProvider, IConnectionTemplateSource, IConnectionEditingDefaults, IConnectionGranularityDefaults, IConnectionCalculationDefaults
{
    /// <summary>Gets the provider discriminator used in document references.</summary>
    public const string ProviderName = "sharepoint";

    /// <summary>How many folders below the listed one a recursive listing descends.</summary>
    public const int MaximumListingDepth = 5;

    /// <summary>How many subfolders one recursive walk visits in total; deeper folders past it are not listed.</summary>
    public const int MaximumListingFolders = 100;

    /// <summary>How many Graph requests one listing call sends before it returns a page.</summary>
    public const int MaximumListingRequests = 10;

    private const int ListingPageSize = 200;

    // Graph's rules for upload sessions: every chunk but the last is a multiple of 320 KiB, and
    // no request carries 60 MiB or more. A single-request upload takes at most 250 MB.
    private const int ChunkUnitBytes = 320 * 1024;
    private const int MaximumChunkBytes = 60 * 1024 * 1024 - ChunkUnitBytes;
    private const long MaximumSimpleUploadBytes = 250L * 1000 * 1000;

    // A chunk that fails for a transient reason is retried this many times before the save is rejected.
    private const int MaximumChunkRetries = 3;
    private static readonly TimeSpan MaximumRetryDelay = TimeSpan.FromSeconds(30);
    private static readonly TimeSpan SessionCancelTimeout = TimeSpan.FromSeconds(30);

    private readonly SharePointDocumentProviderOptions _options;
    private readonly HttpClient _http;
    private readonly IAccessTokenProvider _tokens;
    private readonly ISharePointRegistrationStore _store;
    private readonly string _graphBaseUrl;
    private readonly HashSet<string> _allowedExtensions;
    private readonly string? _creationDriveId;
    private readonly string? _creationFolderItemId;

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
        if (options.UploadSessionThresholdBytes <= 0 ||
            options.UploadSessionThresholdBytes > MaximumSimpleUploadBytes)
            throw new ArgumentOutOfRangeException(nameof(options),
                "UploadSessionThresholdBytes must be positive and at most 250 MB, the limit of a single-request upload.");
        if (options.UploadSessionChunkBytes <= 0 ||
            options.UploadSessionChunkBytes % ChunkUnitBytes != 0 ||
            options.UploadSessionChunkBytes > MaximumChunkBytes)
            throw new ArgumentOutOfRangeException(nameof(options),
                "UploadSessionChunkBytes must be a multiple of 320 KiB (327,680 bytes) and less than 60 MiB.");
        if (options.AllowedExtensions is null || options.AllowedExtensions.Count == 0)
            throw new ArgumentException("At least one allowed document extension is required.", nameof(options));
        if (string.IsNullOrWhiteSpace(options.CreationDriveId) !=
            string.IsNullOrWhiteSpace(options.CreationFolderItemId))
            throw new ArgumentException(
                "SharePoint creation requires both CreationDriveId and CreationFolderItemId.",
                nameof(options));

        _allowedExtensions = new HashSet<string>(
            options.AllowedExtensions.Select(NormalizeExtension),
            StringComparer.OrdinalIgnoreCase);
        _graphBaseUrl = options.GraphBaseUrl.TrimEnd('/');
        _creationDriveId = string.IsNullOrWhiteSpace(options.CreationDriveId)
            ? null
            : options.CreationDriveId.Trim();
        _creationFolderItemId = string.IsNullOrWhiteSpace(options.CreationFolderItemId)
            ? null
            : options.CreationFolderItemId.Trim();

        TemplatesFolder = string.IsNullOrWhiteSpace(options.TemplatesFolder) ? null : options.TemplatesFolder!.Trim();
        if (TemplatesFolder is not null && !IsUrl(TemplatesFolder) && !IsAnyDriveItemPair(TemplatesFolder))
            throw new ArgumentException(
                "TemplatesFolder must be a SharePoint/OneDrive folder URL or a 'driveId/itemId' pair naming the folder.",
                nameof(options));
        TemplateLocation = options.TemplateLocation == TemplateLocationKind.None && TemplatesFolder is not null
            ? TemplateLocationKind.Connection
            : options.TemplateLocation;
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
    /// <remarks>The folder URL or <c>driveId/itemId</c> pair the host configured, or <see langword="null"/>.</remarks>
    public string? TemplatesFolder { get; }

    /// <inheritdoc />
    public TemplateLocationKind TemplateLocation { get; }

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

    /// <inheritdoc />
    public async Task<DocumentReference> CreateAsync(
        string name,
        Stream content,
        CancellationToken cancellationToken = default)
    {
        if (content is null) throw new ArgumentNullException(nameof(content));
        if (_creationDriveId is null || _creationFolderItemId is null)
            throw Error(
                ProviderErrorCode.ConfigurationError,
                "This SharePoint connection has no creation destination. Configure both " +
                "CreationDriveId and CreationFolderItemId, or register an existing document instead.");

        var fileName = ValidateCreateName(name);
        var payload = await PrepareContentAsync(content, itemId: null, cancellationToken).ConfigureAwait(false);
        var destination = $"{ItemUrl(_creationDriveId, _creationFolderItemId)}:/{Uri.EscapeDataString(fileName)}:";
        UploadResult upload;
        try
        {
            upload = await UploadContentAsync(
                simpleUrl: $"{destination}/content?@microsoft.graph.conflictBehavior=fail",
                sessionUrl: $"{destination}/createUploadSession",
                newName: fileName,
                payload,
                ifMatch: null,
                itemId: null,
                cancellationToken,
                conflictCode: ProviderErrorCode.AlreadyExists).ConfigureAwait(false);
        }
        catch (DocumentProviderException ex) when (ex.Code != ProviderErrorCode.AlreadyExists)
        {
            throw Error(
                ex.Code,
                SafeCreateFailureMessage(ex.Code, fileName),
                itemId: null,
                ex);
        }
        catch (Exception ex) when (
            ex is not OperationCanceledException and not DocumentProviderException)
        {
            // Transport and token-provider failures can include the request URL, which
            // contains the host-configured drive and folder ids. Keep those details only
            // on the inner exception available to the host.
            throw Error(
                ProviderErrorCode.IO,
                $"SharePoint could not create document '{fileName}' in the configured destination.",
                itemId: null,
                ex);
        }

        var created = upload.Item;
        var driveId = string.IsNullOrEmpty(created.DriveId) ? _creationDriveId : created.DriveId;
        string registrationId;
        try
        {
            // Once Graph confirms the upload, finish registration even if the caller
            // cancels so the new remote file is not needlessly orphaned.
            registrationId = await _store.AddAsync(
                new SharePointItemRef(driveId, created.Id), CancellationToken.None).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            throw RegistrationFailed(
                $"Document '{fileName}' was created in SharePoint, but its registration " +
                "could not be persisted. The file is stored but unregistered. Do not retry creation; " +
                "register it by name, or report it to the host/operator for recovery.",
                itemId: null, fileName, upload.Sha256, ex);
        }

        return CreateReference(registrationId, created);
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

        var payload = await PrepareContentAsync(content, source.ItemId, cancellationToken).ConfigureAwait(false);

        if (options.Mode == SaveMode.Replace)
        {
            var itemUrl = ItemUrl(driveId, registered.ItemId);
            var replaced = await UploadContentAsync(
                simpleUrl: $"{itemUrl}/content",
                sessionUrl: $"{itemUrl}/createUploadSession",
                newName: null,
                payload,
                ifMatch: item.ETag,
                itemId: source.ItemId,
                cancellationToken).ConfigureAwait(false);
            return CreateReference(source.ItemId, replaced.Item);
        }

        var destinationName = string.IsNullOrWhiteSpace(options.NewName)
            ? await NextVersionedNameAsync(driveId, item, cancellationToken).ConfigureAwait(false)
            : ValidateName(options.NewName!);
        var destination = $"{ItemUrl(driveId, item.ParentId)}:/{Uri.EscapeDataString(destinationName)}:";
        var upload = await UploadContentAsync(
            simpleUrl: $"{destination}/content?@microsoft.graph.conflictBehavior=fail",
            sessionUrl: $"{destination}/createUploadSession",
            newName: destinationName,
            payload,
            ifMatch: null,
            itemId: source.ItemId,
            cancellationToken,
            conflictCode: ProviderErrorCode.AlreadyExists).ConfigureAwait(false);
        var saved = upload.Item;

        var savedDriveId = string.IsNullOrEmpty(saved.DriveId) ? driveId : saved.DriveId;
        string newId;
        try
        {
            // Graph confirmed the upload. Finish registering even if the caller cancels now,
            // so a stored file is not left unregistered behind a plain cancellation.
            newId = await _store.AddAsync(new SharePointItemRef(savedDriveId, saved.Id), CancellationToken.None).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            throw RegistrationFailed(
                $"Document '{destinationName}' was saved to SharePoint, but its registration could not be " +
                "persisted. Do not save it again; register the saved file to recover it.",
                source.ItemId, destinationName, upload.Sha256, ex);
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

    // ── Listing ───────────────────────────────────────────────────────────────

    /// <inheritdoc />
    /// <remarks>
    /// <see cref="DocumentListQuery.Folder"/> is a folder URL, a <c>driveId/itemId</c> pair
    /// (<c>b!…/01…</c> or <c>drives/{driveId}/items/{itemId}</c>), or a path such as
    /// <c>Contracts/2026</c> relative to the connection's creation folder, or to its templates
    /// folder when listing templates. With no folder, the creation folder is listed; a connection
    /// with none needs a folder. Each entry's source is the document's <c>driveId/itemId</c>
    /// pair. The continuation token carries the walk's position, including Graph's own next-page
    /// link, which is only followed when it points at the configured Graph endpoint.
    /// </remarks>
    public async Task<DocumentListPage> ListAsync(DocumentListQuery query, CancellationToken cancellationToken = default)
    {
        if (query is null) throw new ArgumentNullException(nameof(query));
        if (query.Templates && TemplatesFolder is null) return new DocumentListPage();

        var fingerprint = DocumentListingRules.Fingerprint(query, ProviderName);
        var state = string.IsNullOrWhiteSpace(query.ContinuationToken)
            ? await StartListingAsync(query, fingerprint, cancellationToken).ConfigureAwait(false)
            : ListingState.Parse(query.ContinuationToken!, fingerprint, _graphBaseUrl)
              ?? throw Error(ProviderErrorCode.InvalidArgument,
                  "The continuation token is not one this listing issued. Repeat the same query, or start again without the token.");

        var extensions = DocumentListingRules.Extensions(query, _allowedExtensions);
        var pageSize = DocumentListingRules.PageSize(query);
        var items = new List<DocumentListEntry>();
        var requests = 0;

        while (true)
        {
            if (state.Current is null)
            {
                if (state.Queue.Count == 0) return new DocumentListPage { Items = items };
                state.Current = state.Queue[0];
                state.Queue.RemoveAt(0);
                state.PageUrl = ChildrenUrl(state.Current);
                state.Skip = 0;
            }

            if (requests == MaximumListingRequests)
                return new DocumentListPage { Items = items, ContinuationToken = state.Encode() };
            requests++;

            var (children, nextLink) = await ChildrenPageAsync(state.PageUrl!, cancellationToken).ConfigureAwait(false);
            for (var index = state.Skip; index < children.Count; index++)
            {
                var child = children[index];
                if (child.IsFolder)
                {
                    if (query.Recursive && state.Current.Depth < MaximumListingDepth && state.FoldersQueued < MaximumListingFolders)
                    {
                        state.FoldersQueued++;
                        state.Queue.Add(new ListingFolder(
                            string.IsNullOrEmpty(child.DriveId) ? state.Current.DriveId : child.DriveId,
                            child.Id,
                            Combine(state.Current.Display, child.Name),
                            state.Current.Depth + 1));
                    }
                    continue;
                }

                if (!child.IsFile || child.Size > _options.MaximumBytes ||
                    !DocumentListingRules.Matches(child.Name, extensions, query))
                    continue;

                if (items.Count == pageSize)
                {
                    // Another match exists: resume at it, on the same Graph page.
                    state.Skip = index;
                    return new DocumentListPage { Items = items, ContinuationToken = state.Encode() };
                }

                items.Add(new DocumentListEntry
                {
                    Source = $"{(string.IsNullOrEmpty(child.DriveId) ? state.Current.DriveId : child.DriveId)}/{child.Id}",
                    Name = child.Name,
                    Folder = state.Current.Display,
                    ContentType = child.MimeType ?? ContentTypeFor(child.Name),
                    Length = child.Size,
                    LastModified = child.LastModified,
                    IsTemplate = query.Templates || DocumentListingRules.IsTemplateName(child.Name)
                });
            }

            if (nextLink is not null)
            {
                state.PageUrl = nextLink;
                state.Skip = 0;
            }
            else
            {
                state.Current = null;
            }
        }
    }

    /// <summary>Resolves the folder a first page lists and starts the walk there.</summary>
    private async Task<ListingState> StartListingAsync(DocumentListQuery query, string fingerprint, CancellationToken cancellationToken)
    {
        var folder = query.Folder?.Trim();
        var addressed = !string.IsNullOrEmpty(folder) && (IsUrl(folder!) || IsDriveItemPair(folder!));
        if (query.Templates && addressed)
            throw Error(ProviderErrorCode.InvalidArgument,
                "When listing templates, folder is a path inside the templates folder, not a URL or 'driveId/itemId' pair.");

        ListingFolder? root = query.Templates
            ? await ResolveFolderAsync(TemplatesFolder!, cancellationToken).ConfigureAwait(false)
            : _creationDriveId is not null && _creationFolderItemId is not null
                ? new ListingFolder(_creationDriveId, _creationFolderItemId, null, 0)
                : null;

        ListingFolder start;
        if (addressed)
            start = await ResolveFolderAsync(folder!, cancellationToken).ConfigureAwait(false);
        else if (root is null)
            throw Error(ProviderErrorCode.InvalidArgument,
                "This SharePoint connection has no default folder to list. Pass the folder's URL or its 'driveId/itemId' pair.");
        else if (string.IsNullOrEmpty(folder))
            start = root;
        else
            start = await ResolveRelativeFolderAsync(root, folder!, cancellationToken).ConfigureAwait(false);

        return new ListingState(fingerprint) { Queue = { start } };
    }

    /// <summary>Resolves a folder URL or <c>driveId/itemId</c> pair to the folder it names.</summary>
    private async Task<ListingFolder> ResolveFolderAsync(string folder, CancellationToken cancellationToken)
    {
        const string select = "?$select=id,name,folder,file,parentReference";
        string url;
        string? pairDriveId = null;
        if (IsUrl(folder))
        {
            url = $"{_graphBaseUrl}/shares/{EncodeShareUrl(folder)}/driveItem{select}";
        }
        else
        {
            var (driveId, itemId) = ParseDriveAndItemId(folder);
            pairDriveId = driveId;
            url = ItemUrl(driveId, itemId) + select;
        }

        var item = await GetListingItemAsync(url, cancellationToken).ConfigureAwait(false);
        if (!item.IsFolder)
            throw Error(ProviderErrorCode.InvalidArgument, "The folder to list names a file, not a folder.");
        var drive = string.IsNullOrEmpty(item.DriveId) ? pairDriveId : item.DriveId;
        if (string.IsNullOrEmpty(drive))
            throw Error(ProviderErrorCode.IO, "Graph did not report which drive the folder belongs to; it cannot be listed.");
        return new ListingFolder(drive!, item.Id, null, 0);
    }

    /// <summary>Resolves a path such as <c>Contracts/2026</c> beneath a folder, segment by segment escaped.</summary>
    private async Task<ListingFolder> ResolveRelativeFolderAsync(ListingFolder parent, string path, CancellationToken cancellationToken)
    {
        var segments = path.Replace('\\', '/').Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
        if (segments.Length == 0) return parent;
        if (segments.Any(segment => segment.Trim() is "." or ".."))
            throw Error(ProviderErrorCode.AccessDenied, "A folder path cannot step outside the folder it is relative to.");

        var escaped = string.Join("/", segments.Select(Uri.EscapeDataString));
        var item = await GetListingItemAsync(
            $"{ItemUrl(parent.DriveId, parent.ItemId)}:/{escaped}:?$select=id,name,folder,file,parentReference",
            cancellationToken).ConfigureAwait(false);
        if (!item.IsFolder)
            throw Error(ProviderErrorCode.InvalidArgument, "The folder to list names a file, not a folder.");
        return new ListingFolder(
            string.IsNullOrEmpty(item.DriveId) ? parent.DriveId : item.DriveId,
            item.Id,
            Combine(parent.Display, string.Join("/", segments)),
            0);
    }

    private string ChildrenUrl(ListingFolder folder) =>
        $"{ItemUrl(folder.DriveId, folder.ItemId)}/children" +
        $"?$select=id,name,size,file,folder,lastModifiedDateTime,parentReference&$top={ListingPageSize}";

    private async Task<ListingItem> GetListingItemAsync(string url, CancellationToken cancellationToken)
    {
        using var response = await SendAsync(HttpMethod.Get, url, content: null, ifMatch: null, cancellationToken).ConfigureAwait(false);
        await ThrowOnFailureAsync(response, itemId: null).ConfigureAwait(false);
        try
        {
            using var json = JsonDocument.Parse(await response.Content.ReadAsStringAsync().ConfigureAwait(false));
            return ListingItem.Parse(json.RootElement);
        }
        catch (Exception ex) when (ex is JsonException or KeyNotFoundException or InvalidOperationException or FormatException)
        {
            throw Error(ProviderErrorCode.IO, "Graph returned an unreadable drive item.", itemId: null, ex);
        }
    }

    private async Task<(List<ListingItem> Children, string? NextLink)> ChildrenPageAsync(string url, CancellationToken cancellationToken)
    {
        using var response = await SendAsync(HttpMethod.Get, url, content: null, ifMatch: null, cancellationToken).ConfigureAwait(false);
        await ThrowOnFailureAsync(response, itemId: null).ConfigureAwait(false);
        try
        {
            using var json = JsonDocument.Parse(await response.Content.ReadAsStringAsync().ConfigureAwait(false));
            var children = new List<ListingItem>();
            if (json.RootElement.TryGetProperty("value", out var value) && value.ValueKind == JsonValueKind.Array)
                children.AddRange(value.EnumerateArray().Select(ListingItem.Parse));
            var nextLink = json.RootElement.TryGetProperty("@odata.nextLink", out var next) && next.ValueKind == JsonValueKind.String
                ? next.GetString()
                : null;
            // The next-page link is followed with the bearer token, so it must point where every
            // other request goes. A link anywhere else ends the folder rather than leaking the token.
            if (nextLink is not null && !nextLink.StartsWith(_graphBaseUrl + "/", StringComparison.OrdinalIgnoreCase))
                nextLink = null;
            return (children, nextLink);
        }
        catch (Exception ex) when (ex is JsonException or KeyNotFoundException or InvalidOperationException or FormatException)
        {
            throw Error(ProviderErrorCode.IO, "Graph returned an unreadable folder listing.", itemId: null, ex);
        }
    }

    private static string? Combine(string? parent, string child) =>
        string.IsNullOrEmpty(parent) ? child : parent + "/" + child;

    private static bool IsUrl(string value) =>
        value.StartsWith("https://", StringComparison.OrdinalIgnoreCase) ||
        value.StartsWith("http://", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Whether a folder is written as a <c>driveId/itemId</c> pair rather than as a path. A
    /// SharePoint or OneDrive for Business drive id starts with <c>b!</c>; the Graph resource form
    /// <c>drives/{driveId}/items/{itemId}</c> is unambiguous. Anything else is read as a path.
    /// </summary>
    private static bool IsDriveItemPair(string value)
    {
        var trimmed = value.Trim().Trim('/');
        if (trimmed.StartsWith("drives/", StringComparison.OrdinalIgnoreCase) &&
            trimmed.IndexOf("/items/", StringComparison.OrdinalIgnoreCase) > 0)
            return true;
        var slash = trimmed.IndexOf('/');
        return trimmed.StartsWith("b!", StringComparison.Ordinal) &&
               slash > 2 && slash == trimmed.LastIndexOf('/') && slash < trimmed.Length - 1;
    }

    /// <summary>
    /// Whether a configured folder is a <c>driveId/itemId</c> pair of any drive. A configured
    /// templates folder is never a path, so any single <c>/</c> between two ids is a pair here.
    /// </summary>
    private static bool IsAnyDriveItemPair(string value)
    {
        if (IsDriveItemPair(value)) return true;
        var trimmed = value.Trim().Trim('/');
        var slash = trimmed.IndexOf('/');
        return slash > 0 && slash == trimmed.LastIndexOf('/') && slash < trimmed.Length - 1;
    }

    /// <summary>One folder in a listing walk: where it is, how it is shown, and how deep it lies.</summary>
    private sealed record ListingFolder(string DriveId, string ItemId, string? Display, int Depth);

    /// <summary>The fields of a drive item a listing reads.</summary>
    private sealed class ListingItem
    {
        public string Id { get; private init; } = string.Empty;
        public string Name { get; private init; } = string.Empty;
        public string DriveId { get; private init; } = string.Empty;
        public bool IsFolder { get; private init; }
        public bool IsFile { get; private init; }
        public long Size { get; private init; }
        public string? MimeType { get; private init; }
        public DateTimeOffset? LastModified { get; private init; }

        public static ListingItem Parse(JsonElement item) => new()
        {
            Id = item.GetProperty("id").GetString() ?? string.Empty,
            Name = item.TryGetProperty("name", out var name) ? name.GetString() ?? string.Empty : string.Empty,
            DriveId = item.TryGetProperty("parentReference", out var parent) &&
                      parent.TryGetProperty("driveId", out var drive) &&
                      drive.ValueKind == JsonValueKind.String
                ? drive.GetString() ?? string.Empty
                : string.Empty,
            IsFolder = item.TryGetProperty("folder", out _),
            IsFile = item.TryGetProperty("file", out _),
            Size = item.TryGetProperty("size", out var size) && size.ValueKind == JsonValueKind.Number ? size.GetInt64() : 0,
            MimeType = item.TryGetProperty("file", out var file) &&
                       file.TryGetProperty("mimeType", out var mime) &&
                       mime.ValueKind == JsonValueKind.String
                ? mime.GetString()
                : null,
            LastModified = item.TryGetProperty("lastModifiedDateTime", out var modified) &&
                           modified.ValueKind == JsonValueKind.String &&
                           DateTimeOffset.TryParse(modified.GetString(), System.Globalization.CultureInfo.InvariantCulture,
                               System.Globalization.DateTimeStyles.AssumeUniversal, out var parsed)
                ? parsed
                : null
        };
    }

    /// <summary>
    /// A listing walk's position: the folder being read, the Graph page it is on and how many of
    /// that page's children were already handled, and the folders still to read. It travels to
    /// the caller as the continuation token.
    /// </summary>
    private sealed class ListingState
    {
        private const string Version = "sp1";
        private readonly string _fingerprint;

        public ListingState(string fingerprint) => _fingerprint = fingerprint;

        public List<ListingFolder> Queue { get; } = new();
        public ListingFolder? Current { get; set; }
        public string? PageUrl { get; set; }
        public int Skip { get; set; }
        public int FoldersQueued { get; set; }

        public string Encode()
        {
            using var buffer = new MemoryStream();
            using (var writer = new Utf8JsonWriter(buffer))
            {
                writer.WriteStartObject();
                writer.WriteString("v", Version);
                writer.WriteString("f", _fingerprint);
                writer.WriteNumber("c", FoldersQueued);
                if (Current is not null)
                {
                    writer.WritePropertyName("cur");
                    Write(writer, Current);
                    writer.WriteString("u", PageUrl);
                    writer.WriteNumber("s", Skip);
                }
                writer.WriteStartArray("q");
                foreach (var folder in Queue) Write(writer, folder);
                writer.WriteEndArray();
                writer.WriteEndObject();
            }
            return DocumentListingRules.EncodeToken(Encoding.UTF8.GetString(buffer.ToArray()));
        }

        /// <summary>
        /// Reads a token back, or <see langword="null"/> when it is not one this listing issued.
        /// A page link that does not point at the Graph endpoint is refused: it would carry the
        /// bearer token somewhere else.
        /// </summary>
        public static ListingState? Parse(string token, string fingerprint, string graphBaseUrl)
        {
            var text = DocumentListingRules.DecodeToken(token);
            if (text is null) return null;
            try
            {
                using var json = JsonDocument.Parse(text);
                var root = json.RootElement;
                if (root.GetProperty("v").GetString() != Version || root.GetProperty("f").GetString() != fingerprint)
                    return null;

                var state = new ListingState(fingerprint) { FoldersQueued = root.GetProperty("c").GetInt32() };
                if (root.TryGetProperty("cur", out var current))
                {
                    state.Current = Read(current);
                    state.PageUrl = root.GetProperty("u").GetString();
                    state.Skip = root.GetProperty("s").GetInt32();
                    if (state.PageUrl is null || state.Skip < 0 ||
                        !state.PageUrl.StartsWith(graphBaseUrl + "/", StringComparison.OrdinalIgnoreCase))
                        return null;
                }
                foreach (var folder in root.GetProperty("q").EnumerateArray())
                    state.Queue.Add(Read(folder));
                if (state.Queue.Count > MaximumListingFolders + 1 || state.FoldersQueued > MaximumListingFolders)
                    return null;
                return state;
            }
            catch (Exception ex) when (ex is JsonException or KeyNotFoundException or InvalidOperationException or FormatException)
            {
                return null;
            }
        }

        private static void Write(Utf8JsonWriter writer, ListingFolder folder)
        {
            writer.WriteStartObject();
            writer.WriteString("d", folder.DriveId);
            writer.WriteString("i", folder.ItemId);
            if (folder.Display is not null) writer.WriteString("p", folder.Display);
            writer.WriteNumber("l", folder.Depth);
            writer.WriteEndObject();
        }

        private static ListingFolder Read(JsonElement folder)
        {
            var driveId = folder.GetProperty("d").GetString();
            var itemId = folder.GetProperty("i").GetString();
            var depth = folder.GetProperty("l").GetInt32();
            if (string.IsNullOrEmpty(driveId) || string.IsNullOrEmpty(itemId) || depth < 0 || depth > MaximumListingDepth)
                throw new FormatException("A listing folder is incomplete.");
            return new ListingFolder(driveId!, itemId!,
                folder.TryGetProperty("p", out var display) ? display.GetString() : null, depth);
        }
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

    private async Task<GraphItem> UploadAsync(
        string url,
        byte[] bytes,
        string? ifMatch,
        string? itemId,
        CancellationToken cancellationToken,
        ProviderErrorCode conflictCode = ProviderErrorCode.IO)
    {
        using var body = new ByteArrayContent(bytes);
        body.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
        using var response = await SendAsync(HttpMethod.Put, url, body, ifMatch, cancellationToken).ConfigureAwait(false);
        if (response.StatusCode == HttpStatusCode.Conflict)
            throw ConflictError(conflictCode, itemId);
        await ThrowOnFailureAsync(response, itemId, write: true).ConfigureAwait(false);
        return ParseItem(await response.Content.ReadAsStringAsync().ConfigureAwait(false), itemId);
    }

    private DocumentProviderException ConflictError(ProviderErrorCode conflictCode, string? itemId) => Error(
        conflictCode,
        conflictCode == ProviderErrorCode.AlreadyExists
            ? "A document with that name already exists in the configured SharePoint folder; choose another name."
            : "A file already exists at the destination path; refusing to overwrite without Replace mode.",
        itemId);

    // ── Upload sessions ───────────────────────────────────────────────────────

    /// <summary>
    /// Reads the content to save. Seekable content at or above the upload-session threshold is
    /// left in place and read chunk by chunk as it is uploaded; anything else is read into memory
    /// once. Either way content over <see cref="SharePointDocumentProviderOptions.MaximumBytes"/>
    /// is refused before anything is sent.
    /// </summary>
    private async Task<UploadPayload> PrepareContentAsync(Stream content, string? itemId, CancellationToken cancellationToken)
    {
        if (content.CanSeek)
        {
            content.Position = 0;
            var length = content.Length;
            if (length > _options.MaximumBytes)
                throw Error(ProviderErrorCode.ContentTooLarge,
                    $"Document exceeds the configured maximum of {_options.MaximumBytes} bytes.", itemId);
            if (length >= _options.UploadSessionThresholdBytes)
                return new UploadPayload(content, length, bytes: null);
        }

        var bytes = await ReadStreamAsync(content, itemId, cancellationToken).ConfigureAwait(false);
        return new UploadPayload(new MemoryStream(bytes, writable: false), bytes.Length, bytes);
    }

    /// <summary>
    /// Saves content in one request below the upload-session threshold, and through a Graph
    /// upload session from it. Both keep the same preconditions: <c>If-Match</c> for a replace,
    /// <c>conflictBehavior=fail</c> for a new name.
    /// </summary>
    private async Task<UploadResult> UploadContentAsync(
        string simpleUrl,
        string sessionUrl,
        string? newName,
        UploadPayload payload,
        string? ifMatch,
        string? itemId,
        CancellationToken cancellationToken,
        ProviderErrorCode conflictCode = ProviderErrorCode.IO)
    {
        if (payload.Bytes is { } bytes && bytes.LongLength < _options.UploadSessionThresholdBytes)
        {
            var item = await UploadAsync(simpleUrl, bytes, ifMatch, itemId, cancellationToken, conflictCode).ConfigureAwait(false);
            return new UploadResult(item, Sha256Hex(bytes));
        }

        return await UploadInSessionAsync(
            sessionUrl, newName, payload, ifMatch, itemId, cancellationToken, conflictCode).ConfigureAwait(false);
    }

    /// <summary>
    /// Uploads content through a Graph upload session, in chunks that are multiples of 320 KiB.
    /// Graph stores the file only when it accepts the last chunk, which decides how a failure
    /// is reported: before that point nothing was stored, so the session is cancelled and the
    /// save is <see cref="ProviderErrorCode.WriteRejected"/>; once the last chunk is sent, a
    /// failure other than a refusal Graph states is <see cref="ProviderErrorCode.OutcomeUnknown"/>.
    /// </summary>
    private async Task<UploadResult> UploadInSessionAsync(
        string sessionUrl,
        string? newName,
        UploadPayload payload,
        string? ifMatch,
        string? itemId,
        CancellationToken cancellationToken,
        ProviderErrorCode conflictCode)
    {
        var uploadUrl = await CreateUploadSessionAsync(
            sessionUrl, newName, ifMatch, itemId, cancellationToken, conflictCode).ConfigureAwait(false);

        var length = payload.Length;
        var chunkBytes = _options.UploadSessionChunkBytes;
        var buffer = new byte[(int)Math.Min(chunkBytes, length)];
        using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        long hashed = 0;  // the content is hashed once, in order, as it is first read
        long offset = 0;  // the first byte the session expects next
        long sent = 0;    // the end of the furthest range sent
        var failures = 0;
        var lastChunkSent = false;
        try
        {
            while (true)
            {
                var count = (int)Math.Min(chunkBytes, length - offset);
                var last = offset + count == length;
                await ReadRangeAsync(payload.Stream, offset, buffer, count, itemId, cancellationToken).ConfigureAwait(false);
                if (offset + count > hashed)
                {
                    var skip = (int)(hashed - offset);
                    hash.AppendData(buffer, skip, count - skip);
                    hashed = offset + count;
                }
                if (last)
                    await EnsureEndOfContentAsync(payload.Stream, itemId, cancellationToken).ConfigureAwait(false);

                sent = Math.Max(sent, offset + count);
                // From here a lost response can hide a stored file.
                lastChunkSent = last;
                HttpResponseMessage response;
                try
                {
                    response = await SendChunkAsync(uploadUrl, buffer, count, offset, length, cancellationToken).ConfigureAwait(false);
                }
                catch (Exception ex) when (!last && IsTransient(ex, cancellationToken))
                {
                    // The connection dropped mid-chunk: Graph ignores a partial chunk, so ask the
                    // session where to resume.
                    if (++failures > MaximumChunkRetries)
                        throw Error(ProviderErrorCode.WriteRejected,
                            $"The upload failed {failures} times before its last chunk; nothing was stored.", itemId, ex);
                    offset = await ResumeOffsetAsync(uploadUrl, offset, sent, length, itemId, cancellationToken).ConfigureAwait(false);
                    continue;
                }
                catch (Exception ex) when (last && !(ex is OperationCanceledException && cancellationToken.IsCancellationRequested))
                {
                    throw Error(ProviderErrorCode.OutcomeUnknown,
                        "The upload failed while its last chunk was being sent; SharePoint may have stored the document.",
                        itemId, ex);
                }

                using (response)
                {
                    var status = (int)response.StatusCode;
                    if (status == 202 || (!last && response.IsSuccessStatusCode))
                    {
                        // Accepted, with more expected. After the last chunk that means Graph is
                        // still missing an earlier range, so nothing was stored yet.
                        lastChunkSent = false;
                        if (last && ++failures > MaximumChunkRetries)
                            throw Error(ProviderErrorCode.WriteRejected,
                                "Graph kept asking for content it had already been sent; nothing was stored.", itemId);
                        if (!last) failures = 0;
                        var next = await ReadNextExpectedAsync(response).ConfigureAwait(false);
                        if (next is not null)
                            offset = CheckResumeOffset(next.Value, sent, length, itemId);
                        else if (!last)
                            offset += count;
                        else
                            offset = await ResumeOffsetAsync(uploadUrl, offset, sent, length, itemId, cancellationToken).ConfigureAwait(false);
                        continue;
                    }

                    if (last && response.IsSuccessStatusCode)
                    {
                        GraphItem item;
                        try
                        {
                            item = ParseItem(await response.Content.ReadAsStringAsync().ConfigureAwait(false), itemId);
                        }
                        catch (DocumentProviderException ex)
                        {
                            throw Error(ProviderErrorCode.OutcomeUnknown,
                                "Graph accepted the last chunk but its reply could not be read; SharePoint may have stored the document.",
                                itemId, ex);
                        }
                        return new UploadResult(item, ToHex(hash.GetHashAndReset()));
                    }

                    if (last)
                    {
                        // A name conflict or a failed precondition is a refusal to store the file.
                        // Graph keeps such a session open for a corrected commit, so cancel it.
                        if (response.StatusCode == HttpStatusCode.Conflict)
                        {
                            await CancelSessionAsync(uploadUrl).ConfigureAwait(false);
                            throw ConflictError(conflictCode, itemId);
                        }
                        if (response.StatusCode == HttpStatusCode.PreconditionFailed)
                        {
                            await CancelSessionAsync(uploadUrl).ConfigureAwait(false);
                            throw await FailureAsync(response, itemId, write: true).ConfigureAwait(false);
                        }
                        var graphCode = await ReadGraphErrorCodeAsync(response).ConfigureAwait(false);
                        throw Error(ProviderErrorCode.OutcomeUnknown,
                            $"Graph failed the last chunk of the upload with status {status}; SharePoint may have stored the document.{GraphCodeSuffix(graphCode)}",
                            itemId);
                    }

                    if (response.StatusCode == HttpStatusCode.NotFound)
                    {
                        var graphCode = await ReadGraphErrorCodeAsync(response).ConfigureAwait(false);
                        throw Error(ProviderErrorCode.WriteRejected,
                            "The upload session expired before the last chunk was uploaded; nothing was stored. Save again." +
                            GraphCodeSuffix(graphCode), itemId);
                    }

                    if (status is 416 or 429 or >= 500)
                    {
                        // 416: Graph already holds part of this range. 429 and 5xx: try again later.
                        if (++failures > MaximumChunkRetries)
                            throw Error(ProviderErrorCode.WriteRejected,
                                $"Graph failed a chunk of the upload with status {status} {failures} times; nothing was stored.",
                                itemId);
                        if (status != 416)
                            await Task.Delay(RetryDelay(response, failures), cancellationToken).ConfigureAwait(false);
                        offset = await ResumeOffsetAsync(uploadUrl, offset, sent, length, itemId, cancellationToken).ConfigureAwait(false);
                        continue;
                    }

                    throw response.StatusCode == HttpStatusCode.Conflict
                        ? ConflictError(conflictCode, itemId)
                        : await FailureAsync(response, itemId, write: true).ConfigureAwait(false);
                }
            }
        }
        catch (Exception ex) when (!lastChunkSent)
        {
            // Nothing is stored until Graph accepts the last chunk, so this save is refused.
            // Cancelling the session discards what it holds.
            await CancelSessionAsync(uploadUrl).ConfigureAwait(false);
            if (ex is DocumentProviderException { Code: not ProviderErrorCode.OutcomeUnknown and not ProviderErrorCode.IO })
                throw;
            throw Error(ProviderErrorCode.WriteRejected,
                ex is OperationCanceledException && cancellationToken.IsCancellationRequested
                    ? "The save was cancelled before its last chunk was uploaded; nothing was stored."
                    : "The upload failed before its last chunk; nothing was stored.",
                itemId, ex);
        }
    }

    /// <summary>Opens an upload session and returns its pre-authenticated upload URL.</summary>
    private async Task<Uri> CreateUploadSessionAsync(
        string url,
        string? newName,
        string? ifMatch,
        string? itemId,
        CancellationToken cancellationToken,
        ProviderErrorCode conflictCode)
    {
        using var body = new StringContent(UploadSessionBody(newName), Encoding.UTF8, "application/json");
        HttpResponseMessage response;
        try
        {
            response = await SendAsync(HttpMethod.Post, url, body, ifMatch, cancellationToken).ConfigureAwait(false);
        }
        catch (Exception ex) when (ex is not DocumentProviderException)
        {
            // No content has been sent, so nothing can have been stored.
            throw Error(ProviderErrorCode.WriteRejected,
                "Graph could not open an upload session; nothing was stored.", itemId, ex);
        }

        using (response)
        {
            if (response.StatusCode == HttpStatusCode.Conflict)
                throw ConflictError(conflictCode, itemId);
            if (!response.IsSuccessStatusCode)
            {
                var failure = await FailureAsync(response, itemId, write: true).ConfigureAwait(false);
                throw failure.Code == ProviderErrorCode.OutcomeUnknown
                    ? Error(ProviderErrorCode.WriteRejected,
                        $"Graph failed to open an upload session with status {(int)response.StatusCode}; nothing was stored.",
                        itemId)
                    : failure;
            }

            string? uploadUrl = null;
            try
            {
                using var json = JsonDocument.Parse(await response.Content.ReadAsStringAsync().ConfigureAwait(false));
                if (json.RootElement.ValueKind == JsonValueKind.Object &&
                    json.RootElement.TryGetProperty("uploadUrl", out var value) &&
                    value.ValueKind == JsonValueKind.String)
                    uploadUrl = value.GetString();
            }
            catch (JsonException)
            {
            }

            if (!Uri.TryCreate(uploadUrl, UriKind.Absolute, out var uri) || uri.Scheme != Uri.UriSchemeHttps)
                throw Error(ProviderErrorCode.WriteRejected,
                    "Graph did not return a usable upload session URL; nothing was stored.", itemId);
            return uri;
        }
    }

    /// <summary>
    /// The session request body. A replace keeps the item it addresses, guarded by
    /// <c>If-Match</c>; a new name keeps <c>conflictBehavior=fail</c>, so it never overwrites.
    /// </summary>
    private static string UploadSessionBody(string? newName)
    {
        using var buffer = new MemoryStream();
        using (var writer = new Utf8JsonWriter(buffer))
        {
            writer.WriteStartObject();
            writer.WriteStartObject("item");
            writer.WriteString("@microsoft.graph.conflictBehavior", newName is null ? "replace" : "fail");
            if (newName is not null)
                writer.WriteString("name", newName);
            writer.WriteEndObject();
            writer.WriteEndObject();
        }
        return Encoding.UTF8.GetString(buffer.ToArray());
    }

    /// <summary>
    /// Sends one chunk. The upload URL is pre-authenticated, so the bearer token is never sent
    /// to it; Graph can refuse a chunk that carries one.
    /// </summary>
    private async Task<HttpResponseMessage> SendChunkAsync(
        Uri uploadUrl, byte[] buffer, int count, long offset, long length, CancellationToken cancellationToken)
    {
        using var request = new HttpRequestMessage(HttpMethod.Put, uploadUrl);
        var content = new ByteArrayContent(buffer, 0, count);
        content.Headers.ContentRange = new ContentRangeHeaderValue(offset, offset + count - 1, length);
        request.Content = content;
        return await _http.SendAsync(request, HttpCompletionOption.ResponseContentRead, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Asks the session which byte it expects next. When the session cannot say, the upload
    /// resumes where it was; Graph answers 416 if it already holds that range.
    /// </summary>
    private async Task<long> ResumeOffsetAsync(
        Uri uploadUrl, long current, long sent, long length, string? itemId, CancellationToken cancellationToken)
    {
        HttpResponseMessage response;
        try
        {
            using var request = new HttpRequestMessage(HttpMethod.Get, uploadUrl);
            response = await _http.SendAsync(request, HttpCompletionOption.ResponseContentRead, cancellationToken).ConfigureAwait(false);
        }
        catch (Exception ex) when (IsTransient(ex, cancellationToken))
        {
            return current;
        }

        using (response)
        {
            if (response.StatusCode == HttpStatusCode.NotFound)
                throw Error(ProviderErrorCode.WriteRejected,
                    "The upload session expired before the last chunk was uploaded; nothing was stored. Save again.",
                    itemId);
            if (!response.IsSuccessStatusCode)
                return current;
            var next = await ReadNextExpectedAsync(response).ConfigureAwait(false);
            return next is null ? current : CheckResumeOffset(next.Value, sent, length, itemId);
        }
    }

    /// <summary>
    /// Checks a position the session reported: it must be within what was sent and short of the
    /// end, since the session completes when the last byte arrives.
    /// </summary>
    private long CheckResumeOffset(long next, long sent, long length, string? itemId)
    {
        if (next < 0 || next > sent || next >= length)
            throw Error(ProviderErrorCode.WriteRejected,
                "The upload session reported a position the upload never reached; nothing was stored.", itemId);
        return next;
    }

    /// <summary>The start of the earliest range in a session's <c>nextExpectedRanges</c>, if it names one.</summary>
    private static async Task<long?> ReadNextExpectedAsync(HttpResponseMessage response)
    {
        try
        {
            var payload = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            if (string.IsNullOrWhiteSpace(payload)) return null;
            using var json = JsonDocument.Parse(payload);
            if (json.RootElement.ValueKind != JsonValueKind.Object ||
                !json.RootElement.TryGetProperty("nextExpectedRanges", out var ranges) ||
                ranges.ValueKind != JsonValueKind.Array)
                return null;

            long? earliest = null;
            foreach (var range in ranges.EnumerateArray())
            {
                var text = range.ValueKind == JsonValueKind.String ? range.GetString() : null;
                if (string.IsNullOrEmpty(text)) continue;
                var dash = text!.IndexOf('-');
                var start = dash < 0 ? text : text.Substring(0, dash);
                if (long.TryParse(start, System.Globalization.NumberStyles.None,
                        System.Globalization.CultureInfo.InvariantCulture, out var value) &&
                    (earliest is null || value < earliest))
                    earliest = value;
            }
            return earliest;
        }
        catch (JsonException)
        {
            return null;
        }
    }

    /// <summary>Cancels a session so Graph discards what it holds. Best effort: an abandoned session expires.</summary>
    private async Task CancelSessionAsync(Uri uploadUrl)
    {
        try
        {
            using var timeout = new CancellationTokenSource(SessionCancelTimeout);
            using var request = new HttpRequestMessage(HttpMethod.Delete, uploadUrl);
            using var response = await _http.SendAsync(request, timeout.Token).ConfigureAwait(false);
        }
        catch (Exception)
        {
            // The session expires on its own, and it never stores a file without its last chunk.
        }
    }

    /// <summary>Reads one range of the content into the chunk buffer.</summary>
    private async Task ReadRangeAsync(
        Stream stream, long offset, byte[] buffer, int count, string? itemId, CancellationToken cancellationToken)
    {
        if (stream.Position != offset)
            stream.Position = offset;
        var read = 0;
        while (read < count)
        {
            var n = await stream.ReadAsync(buffer, read, count - read, cancellationToken).ConfigureAwait(false);
            if (n == 0)
                throw Error(ProviderErrorCode.WriteRejected,
                    "The document content ended before its stated length while it was being uploaded; nothing was stored.",
                    itemId);
            read += n;
        }
    }

    /// <summary>Refuses content that grew past its stated length while it was being uploaded.</summary>
    private async Task EnsureEndOfContentAsync(Stream stream, string? itemId, CancellationToken cancellationToken)
    {
        var probe = new byte[1];
        if (await stream.ReadAsync(probe, 0, 1, cancellationToken).ConfigureAwait(false) != 0)
            throw Error(ProviderErrorCode.WriteRejected,
                "The document content grew past its stated length while it was being uploaded; nothing was stored.",
                itemId);
    }

    /// <summary>A dropped connection or a timeout the caller did not ask for.</summary>
    private static bool IsTransient(Exception ex, CancellationToken cancellationToken) =>
        ex is HttpRequestException or IOException ||
        (ex is OperationCanceledException && !cancellationToken.IsCancellationRequested);

    /// <summary>Graph's <c>Retry-After</c> when it sends one, else an exponential back-off.</summary>
    private static TimeSpan RetryDelay(HttpResponseMessage response, int failures)
    {
        var retryAfter = response.Headers.RetryAfter;
        var delay = retryAfter?.Delta
            ?? (retryAfter?.Date is { } date ? date - DateTimeOffset.UtcNow : (TimeSpan?)null)
            ?? TimeSpan.FromSeconds(Math.Pow(2, failures - 1));
        if (delay < TimeSpan.Zero) return TimeSpan.Zero;
        return delay > MaximumRetryDelay ? MaximumRetryDelay : delay;
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

    private async Task ThrowOnFailureAsync(HttpResponseMessage response, string? itemId, bool write = false)
    {
        if (response.IsSuccessStatusCode) return;
        throw await FailureAsync(response, itemId, write).ConfigureAwait(false);
    }

    /// <summary>The provider error a failed Graph response stands for.</summary>
    private async Task<DocumentProviderException> FailureAsync(HttpResponseMessage response, string? itemId, bool write)
    {
        var graphCode = await ReadGraphErrorCodeAsync(response).ConfigureAwait(false);
        var suffix = GraphCodeSuffix(graphCode);
        return response.StatusCode switch
        {
            // Retrying cannot help a locked file, so the message says what can.
            _ when write && IsLocked(response.StatusCode, graphCode) => Error(ProviderErrorCode.WriteRejected,
                "The document is locked in SharePoint, for example checked out or open for exclusive editing; " +
                "nothing was stored. Ask whoever holds the lock to check the file in or close it, or save as a " +
                $"new document instead.{suffix}", itemId),
            HttpStatusCode.NotFound => Error(ProviderErrorCode.NotFound,
                $"The referenced SharePoint item no longer exists.{suffix}", itemId),
            HttpStatusCode.Unauthorized or HttpStatusCode.Forbidden => Error(ProviderErrorCode.AccessDenied,
                $"Graph refused access to the item.{suffix}", itemId),
            HttpStatusCode.PreconditionFailed => new DocumentVersionConflictException(
                "(If-Match)", "(changed)", ProviderName, ConnectionId, itemId),
            HttpStatusCode.RequestEntityTooLarge => Error(ProviderErrorCode.ContentTooLarge,
                $"Graph rejected the content as too large.{suffix}", itemId),
            // For an upload, any other refusal status (400, 409, 423 locked, 429 throttled) is
            // Graph declining the request: nothing was stored. A server error is not a refusal;
            // the request may have been processed, so the outcome is unknown.
            _ when write && (int)response.StatusCode is >= 400 and < 500 => Error(ProviderErrorCode.WriteRejected,
                $"Graph refused the upload with status {(int)response.StatusCode}; nothing was stored.{suffix}", itemId),
            _ when write => Error(ProviderErrorCode.OutcomeUnknown,
                $"Graph failed the upload with status {(int)response.StatusCode}; it may have been stored.{suffix}", itemId),
            _ => Error(ProviderErrorCode.IO,
                $"Graph request failed with status {(int)response.StatusCode}.{suffix}", itemId)
        };
    }

    private static string GraphCodeSuffix(string graphCode) =>
        string.IsNullOrEmpty(graphCode) ? string.Empty : $" Graph code: {graphCode}.";

    private static bool IsLocked(HttpStatusCode status, string graphCode) =>
        (int)status == 423 || string.Equals(graphCode, "resourceLocked", StringComparison.OrdinalIgnoreCase);

    private static async Task<string> ReadGraphErrorCodeAsync(HttpResponseMessage response)
    {
        try
        {
            var payload = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            if (string.IsNullOrWhiteSpace(payload)) return string.Empty;
            using var json = JsonDocument.Parse(payload);
            return json.RootElement.TryGetProperty("error", out var error) &&
                   error.TryGetProperty("code", out var code)
                ? code.GetString() ?? string.Empty
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

    private string ValidateCreateName(string name)
    {
        if (!string.IsNullOrWhiteSpace(name) &&
            PortableDocumentName.ContainsInvalidCharacter(name))
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
            $"The configured SharePoint creation destination for '{fileName}' was not found.",
        ProviderErrorCode.AccessDenied =>
            $"The configured SharePoint identity cannot create '{fileName}' in the creation destination.",
        ProviderErrorCode.ContentTooLarge =>
            $"SharePoint rejected '{fileName}' because its content is too large.",
        _ => $"SharePoint could not create document '{fileName}' in the configured destination."
    };

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

    private static string ContentTypeFor(string name) => OfficeContentTypes.ForName(name);

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
    /// Graph confirmed the upload but the registration store failed: carry the file's name and
    /// content hash so it can be found and registered, never its drive or item id.
    /// </summary>
    private DocumentRegistrationFailedException RegistrationFailed(
        string message, string? itemId, string outputName, string outputSha256, Exception innerException) =>
        new(message, ProviderName, ConnectionId, itemId, outputName, outputSha256, innerException);

    private static string Sha256Hex(byte[] bytes)
    {
        using var sha = SHA256.Create();
        return ToHex(sha.ComputeHash(bytes));
    }

    private static string ToHex(byte[] hash) => string.Concat(hash.Select(b => b.ToString("x2")));

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

    /// <summary>
    /// Content ready to upload: a seekable stream of known length, and the bytes themselves when
    /// they were read into memory.
    /// </summary>
    private sealed class UploadPayload
    {
        public UploadPayload(Stream stream, long length, byte[]? bytes)
        {
            Stream = stream;
            Length = length;
            Bytes = bytes;
        }

        public Stream Stream { get; }
        public long Length { get; }
        public byte[]? Bytes { get; }
    }

    /// <summary>The stored item, and the SHA-256 of the bytes uploaded to it.</summary>
    private sealed class UploadResult
    {
        public UploadResult(GraphItem item, string sha256)
        {
            Item = item;
            Sha256 = sha256;
        }

        public GraphItem Item { get; }
        public string Sha256 { get; }
    }
}
