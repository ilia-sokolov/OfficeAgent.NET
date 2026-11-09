using Azure;
using Azure.Identity;
using Azure.Storage.Blobs;
using Azure.Storage.Blobs.Models;
using Azure.Storage.Blobs.Specialized;
using OfficeAgent.Abstractions;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.ObjectStorage;

namespace OfficeAgent.AzureBlob;

/// <summary>
/// Azure Blob Storage document provider: a registry of references to blobs in one container,
/// under an optional prefix, addressed by an opaque item id assigned at registration. Documents
/// are registered by a blob name relative to the prefix (or the blob's URL in the same
/// container); a name that leaves the prefix is refused with <see cref="ProviderErrorCode.AccessDenied"/>.
/// </summary>
/// <remarks>
/// Versions are blob ETags. <see cref="SaveMode.Replace"/> uploads with <c>If-Match</c>, and
/// every new blob - a versioned sibling, a renamed copy, a created document - is uploaded with
/// <c>If-None-Match: *</c>, so the provider never overwrites content it did not open. Each upload
/// carries a random <c>officeagent_write_id</c> metadata value: when a precondition fails after
/// the SDK retried a request whose first attempt landed, the provider reads the blob back and
/// recognises its own write instead of reporting a conflict. Removing a registration never
/// deletes the blob. The agent sees opaque ids and file names, never account URLs or credentials.
/// Listing uses <c>GetBlobsByHierarchy</c> with the <c>/</c> delimiter, one prefix level per
/// request, and never reads outside the prefix.
/// </remarks>
public sealed class AzureBlobDocumentProvider : IDocumentProvider, IDocumentCreatingProvider, IDocumentListingProvider, IConnectionTemplateSource, IConnectionEditingDefaults, IConnectionGranularityDefaults, IConnectionCalculationDefaults
{
    /// <summary>Gets the provider discriminator used in document references.</summary>
    public const string ProviderName = "azureblob";

    /// <summary>The metadata entry each upload carries so a retried write can recognise itself.</summary>
    internal const string WriteMarker = "officeagent_write_id";

    private const int NameRaceAttempts = 5;

    private readonly AzureBlobDocumentProviderOptions _options;
    private readonly BlobContainerClient _container;
    private readonly IAzureBlobRegistrationStore _store;
    private readonly string _prefix;
    private readonly string? _templatesPrefix;
    private readonly string? _templatesFolder;
    private readonly TemplateLocationKind _templateLocation;
    private readonly HashSet<string> _allowedExtensions;

    /// <summary>
    /// Initializes a provider from configuration. With <see cref="AzureBlobDocumentProviderOptions.ServiceUri"/>
    /// it authenticates with <c>DefaultAzureCredential</c>; with
    /// <see cref="AzureBlobDocumentProviderOptions.ConnectionString"/> it uses that instead.
    /// </summary>
    public AzureBlobDocumentProvider(
        AzureBlobDocumentProviderOptions options,
        IAzureBlobRegistrationStore? registrationStore = null)
        : this(options, CreateContainerClient(options), registrationStore)
    {
    }

    /// <summary>
    /// Initializes a provider over a container client the host built, for a custom credential, a
    /// SAS, or the Azurite emulator. <see cref="AzureBlobDocumentProviderOptions.ServiceUri"/>,
    /// <see cref="AzureBlobDocumentProviderOptions.ConnectionString"/> and
    /// <see cref="AzureBlobDocumentProviderOptions.ContainerName"/> are not used.
    /// </summary>
    public AzureBlobDocumentProvider(
        AzureBlobDocumentProviderOptions options,
        BlobContainerClient container,
        IAzureBlobRegistrationStore? registrationStore = null)
    {
        _options = options ?? throw new ArgumentNullException(nameof(options));
        _container = container ?? throw new ArgumentNullException(nameof(container));
        if (string.IsNullOrWhiteSpace(options.ConnectionId))
            throw new ArgumentException("An Azure Blob provider connection id is required.", nameof(options));
        if (options.MaximumBytes <= 0)
            throw new ArgumentOutOfRangeException(nameof(options), "MaximumBytes must be positive.");
        if (options.AllowedExtensions is null || options.AllowedExtensions.Count == 0)
            throw new ArgumentException("At least one allowed document extension is required.", nameof(options));

        _prefix = ObjectKeys.NormalizePrefix(options.Prefix, nameof(options));
        _templatesFolder = ObjectKeys.NormalizeFolder(options.TemplatesFolder, nameof(options));
        _templatesPrefix = _templatesFolder is null ? null : _prefix + _templatesFolder + "/";
        _templateLocation = CloudListing.TemplateLocationFor(_templatesFolder is not null, options.TemplateLocation, nameof(options));
        _allowedExtensions = new HashSet<string>(
            options.AllowedExtensions.Select(ObjectKeys.NormalizeExtension),
            StringComparer.OrdinalIgnoreCase);
        _store = registrationStore ?? (string.IsNullOrWhiteSpace(options.RegistrationIndexPath)
            ? new InMemoryRegistrationIndex()
            : new JsonFileRegistrationIndex(options.RegistrationIndexPath, ProviderName));
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
    /// <remarks>The folder relative to the prefix, such as <c>templates</c>.</remarks>
    public string? TemplatesFolder => _templatesFolder;

    /// <inheritdoc />
    public TemplateLocationKind TemplateLocation => _templateLocation;

    /// <inheritdoc />
    public async Task<DocumentReference> RegisterAsync(
        string source,
        CancellationToken cancellationToken = default)
    {
        var blobName = ResolveSource(source);
        ValidateExtension(ObjectKeys.Leaf(blobName), itemId: null);

        var properties = await GetPropertiesAsync(blobName, itemId: null, beforeWrite: false, cancellationToken).ConfigureAwait(false)
            ?? throw Error(ProviderErrorCode.NotFound, $"No blob named '{RelativeName(blobName)}' exists in this connection.");
        if (properties.ContentLength > _options.MaximumBytes)
            throw Error(ProviderErrorCode.ContentTooLarge,
                $"Document '{RelativeName(blobName)}' exceeds the configured maximum of {_options.MaximumBytes} bytes.");

        var id = await _store.AddAsync(blobName, cancellationToken).ConfigureAwait(false);
        return CreateReference(id, blobName, properties.ETag.ToString());
    }

    /// <inheritdoc />
    public async Task<DocumentReference> CreateAsync(
        string name,
        Stream content,
        CancellationToken cancellationToken = default)
    {
        if (content is null) throw new ArgumentNullException(nameof(content));

        var fileName = ValidateCreateName(name);
        var blobName = _prefix + fileName;
        var bytes = await ReadContentAsync(content, itemId: null, cancellationToken).ConfigureAwait(false);

        string etag;
        try
        {
            etag = await WriteAsync(blobName, bytes, CreateOnly(), NewHeaders(fileName),
                new Dictionary<string, string>(), itemId: null, replace: false, cancellationToken).ConfigureAwait(false);
        }
        catch (DocumentProviderException ex) when (ex.Code == ProviderErrorCode.AlreadyExists)
        {
            throw Error(ProviderErrorCode.AlreadyExists,
                $"A document named '{fileName}' already exists in this connection; choose another name.", itemId: null, ex);
        }

        string id;
        try
        {
            // The blob is stored. Finish registering even if the caller cancels now, so the new
            // blob is not orphaned behind a plain cancellation.
            id = await _store.AddAsync(blobName, CancellationToken.None).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            throw RegistrationFailed(
                $"Document '{fileName}' was created in Azure Blob Storage, but its registration could not be " +
                "persisted. The blob is stored but unregistered. Do not retry creation; register it by name instead.",
                itemId: null, fileName, bytes, ex);
        }
        return CreateReference(id, blobName, etag);
    }

    /// <inheritdoc />
    public async Task<DocumentContent> OpenReadAsync(
        DocumentReference reference,
        CancellationToken cancellationToken = default)
    {
        ValidateReference(reference);
        var blobName = await ResolveRegistrationAsync(reference.ItemId, cancellationToken).ConfigureAwait(false);

        var properties = await GetPropertiesAsync(blobName, reference.ItemId, beforeWrite: false, cancellationToken).ConfigureAwait(false)
            ?? throw Error(ProviderErrorCode.NotFound, "The referenced blob no longer exists.", reference.ItemId);
        if (properties.ContentLength > _options.MaximumBytes)
            throw Error(ProviderErrorCode.ContentTooLarge,
                $"Document exceeds the configured maximum of {_options.MaximumBytes} bytes.", reference.ItemId);

        var etag = properties.ETag.ToString();
        if (!string.IsNullOrEmpty(reference.Version) && !string.Equals(reference.Version, etag, StringComparison.Ordinal))
            throw new DocumentVersionConflictException(reference.Version!, etag, ProviderName, ConnectionId, reference.ItemId);

        // The download is pinned to the ETag just checked, so the bytes are the version reported.
        var bytes = await DownloadAsync(blobName, properties.ETag, reference.ItemId, cancellationToken).ConfigureAwait(false);
        return new DocumentContent(CreateReference(reference.ItemId, blobName, etag), new MemoryStream(bytes, writable: false));
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
                "Azure Blob item ids are opaque and provider-assigned; a destination id cannot be chosen by the caller. Use Mode=NewVersion/NewDocument (optionally with NewName) to mint a new document, or Mode=Replace to overwrite the source.",
                source.ItemId);
        if (options.Mode == SaveMode.Replace && !string.IsNullOrWhiteSpace(options.NewName))
            throw Error(ProviderErrorCode.InvalidArgument,
                "Replace mode overwrites the source in place and cannot rename. Use Mode=NewVersion/NewDocument with NewName to mint a renamed document.",
                source.ItemId);

        var blobName = await ResolveRegistrationAsync(source.ItemId, cancellationToken).ConfigureAwait(false);
        var properties = await GetPropertiesAsync(blobName, source.ItemId, beforeWrite: true, cancellationToken).ConfigureAwait(false)
            ?? throw Error(ProviderErrorCode.NotFound, "The referenced blob no longer exists.", source.ItemId);

        var currentVersion = properties.ETag.ToString();
        var expectedVersion = options.ExpectedVersion ?? source.Version;
        if (!string.IsNullOrEmpty(expectedVersion) && !string.Equals(expectedVersion, currentVersion, StringComparison.Ordinal))
            throw new DocumentVersionConflictException(expectedVersion!, currentVersion, ProviderName, ConnectionId, source.ItemId);

        var bytes = await ReadContentAsync(content, source.ItemId, cancellationToken).ConfigureAwait(false);

        if (options.Mode == SaveMode.Replace)
        {
            // Conditioned on the ETag just checked, so a change between the check and the upload
            // is a conflict, not a silent overwrite. Headers and metadata are carried over.
            var etag = await WriteAsync(
                blobName, bytes, new BlobRequestConditions { IfMatch = properties.ETag },
                CarriedHeaders(properties, ObjectKeys.Leaf(blobName)), CarriedMetadata(properties),
                source.ItemId, replace: true, cancellationToken).ConfigureAwait(false);
            return CreateReference(source.ItemId, blobName, etag);
        }

        var (destination, savedEtag) = string.IsNullOrWhiteSpace(options.NewName)
            ? await WriteNextVersionAsync(blobName, bytes, source.ItemId, cancellationToken).ConfigureAwait(false)
            : await WriteNamedAsync(ObjectKeys.Parent(blobName) + ValidateName(options.NewName!), bytes, source.ItemId, cancellationToken).ConfigureAwait(false);

        string newId;
        try
        {
            // The blob is stored. Finish registering even if the caller cancels now.
            newId = await _store.AddAsync(destination, CancellationToken.None).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            throw RegistrationFailed(
                $"Document '{RelativeName(destination)}' was saved to Azure Blob Storage, but its registration could not be " +
                "persisted. Do not save it again; register the saved blob by name to recover it.",
                source.ItemId, RelativeName(destination), bytes, ex);
        }
        return CreateReference(newId, destination, savedEtag);
    }

    /// <inheritdoc />
    public async Task RemoveAsync(
        DocumentReference reference,
        CancellationToken cancellationToken = default)
    {
        ValidateReference(reference);
        if (string.IsNullOrWhiteSpace(reference.ItemId))
            throw Error(ProviderErrorCode.InvalidArgument, "An item id is required.");
        if (!await _store.RemoveAsync(reference.ItemId, cancellationToken).ConfigureAwait(false))
            throw Error(ProviderErrorCode.NotFound,
                $"The provider document '{reference.ItemId}' is not registered.", reference.ItemId);
    }

    /// <inheritdoc />
    /// <remarks>
    /// Lists with <c>GetBlobsByHierarchy</c> under the prefix (or the templates folder), one prefix
    /// level per request with the <c>/</c> delimiter. Each entry's source is the blob name relative
    /// to the prefix. A page makes at most <see cref="CloudListing.MaximumServiceCalls"/> requests,
    /// and a recursive listing descends at most <see cref="CloudListing.MaximumDepth"/> levels.
    /// </remarks>
    public Task<DocumentListPage> ListAsync(DocumentListQuery query, CancellationToken cancellationToken = default) =>
        ObjectListing.ListAsync(
            query, ConnectionId, _prefix, _templatesPrefix, _allowedExtensions, _options.MaximumBytes,
            (code, message) => Error(code, message), ReadListingPageAsync, cancellationToken);

    // ── Writes ────────────────────────────────────────────────────────────────

    private async Task<(string BlobName, string ETag)> WriteNamedAsync(
        string blobName, byte[] bytes, string itemId, CancellationToken cancellationToken)
    {
        EnsureInsidePrefix(blobName, itemId);
        var etag = await WriteAsync(blobName, bytes, CreateOnly(), NewHeaders(ObjectKeys.Leaf(blobName)),
            new Dictionary<string, string>(), itemId, replace: false, cancellationToken).ConfigureAwait(false);
        return (blobName, etag);
    }

    /// <summary>
    /// Writes the next free versioned sibling (<c>contract.v2.docx</c>, …). Each candidate is
    /// probed first so a taken name costs a HEAD, not an upload; a name another writer takes
    /// between the probe and the upload fails the create-only upload and the next one is tried.
    /// </summary>
    private async Task<(string BlobName, string ETag)> WriteNextVersionAsync(
        string sourceBlobName, byte[] bytes, string itemId, CancellationToken cancellationToken)
    {
        var parent = ObjectKeys.Parent(sourceBlobName);
        var races = 0;
        foreach (var candidate in ObjectKeys.VersionedNames(ObjectKeys.Leaf(sourceBlobName)))
        {
            var blobName = parent + candidate;
            if (blobName.Length - _prefix.Length > ObjectKeys.MaximumKeyLength) break;
            if (await GetPropertiesAsync(blobName, itemId, beforeWrite: true, cancellationToken).ConfigureAwait(false) is not null)
                continue;
            try
            {
                return await WriteNamedAsync(blobName, bytes, itemId, cancellationToken).ConfigureAwait(false);
            }
            catch (DocumentProviderException ex) when (ex.Code == ProviderErrorCode.AlreadyExists && ++races < NameRaceAttempts)
            {
            }
        }
        throw Error(ProviderErrorCode.AlreadyExists,
            "No free versioned name is left for this document; save with NewName instead.", itemId);
    }

    /// <summary>
    /// Uploads one blob under a precondition and returns its new ETag. A refusal the service
    /// reports before storing anything is certain; a server error, a timeout or a lost response is
    /// an unknown outcome and is never reported as nothing written.
    /// </summary>
    private async Task<string> WriteAsync(
        string blobName,
        byte[] bytes,
        BlobRequestConditions conditions,
        BlobHttpHeaders headers,
        IDictionary<string, string> metadata,
        string? itemId,
        bool replace,
        CancellationToken cancellationToken)
    {
        var marker = Guid.NewGuid().ToString("N");
        metadata[WriteMarker] = marker;
        var blob = _container.GetBlockBlobClient(blobName);
        try
        {
            using var body = new MemoryStream(bytes, writable: false);
            var response = await blob.UploadAsync(body, new BlobUploadOptions
            {
                HttpHeaders = headers,
                Metadata = metadata,
                Conditions = conditions
            }, cancellationToken).ConfigureAwait(false);
            return response.Value.ETag.ToString();
        }
        catch (RequestFailedException ex) when (IsPreconditionFailure(ex))
        {
            // The SDK retries timeouts and server errors. If an earlier attempt stored the blob,
            // the retry fails its own precondition: recognise that write before calling it a conflict.
            var landed = await FindOwnWriteAsync(blob, marker, itemId, cancellationToken).ConfigureAwait(false);
            if (landed is not null) return landed;
            if (replace)
                throw new DocumentVersionConflictException("(If-Match)", "(changed)", ProviderName, ConnectionId, itemId);
            throw Error(ProviderErrorCode.AlreadyExists,
                $"A blob named '{RelativeName(blobName)}' already exists; refusing to overwrite it.", itemId);
        }
        catch (RequestFailedException ex)
        {
            throw WriteFailure(ex, itemId);
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            throw;
        }
        catch (Exception ex) when (ex is not DocumentProviderException)
        {
            throw Error(ProviderErrorCode.OutcomeUnknown,
                "The upload to Azure Blob Storage failed before a response arrived; it may have been stored.", itemId, ex);
        }
    }

    private static bool IsPreconditionFailure(RequestFailedException ex) =>
        ex.Status == 412 ||
        (ex.Status == 409 && string.Equals(ex.ErrorCode, "BlobAlreadyExists", StringComparison.OrdinalIgnoreCase));

    private async Task<string?> FindOwnWriteAsync(
        BlockBlobClient blob, string marker, string? itemId, CancellationToken cancellationToken)
    {
        try
        {
            var properties = (await blob.GetPropertiesAsync(cancellationToken: cancellationToken).ConfigureAwait(false)).Value;
            return properties.Metadata is not null &&
                   properties.Metadata.TryGetValue(WriteMarker, out var found) &&
                   string.Equals(found, marker, StringComparison.Ordinal)
                ? properties.ETag.ToString()
                : null;
        }
        catch (RequestFailedException ex) when (ex.Status == 404)
        {
            return null;
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            throw;
        }
        catch (Exception ex)
        {
            throw Error(ProviderErrorCode.OutcomeUnknown,
                "Azure Blob Storage refused the upload's precondition and the blob could not be read back " +
                "to tell whether an earlier attempt stored it.", itemId, ex);
        }
    }

    private DocumentProviderException WriteFailure(RequestFailedException ex, string? itemId)
    {
        var suffix = CodeSuffix(ex);
        return ex.Status switch
        {
            404 => Error(ProviderErrorCode.NotFound,
                $"The blob or its container no longer exists; nothing was stored.{suffix}", itemId, ex),
            401 or 403 => Error(ProviderErrorCode.AccessDenied,
                $"Azure Blob Storage refused access to the blob; nothing was stored.{suffix}", itemId, ex),
            413 => Error(ProviderErrorCode.ContentTooLarge,
                $"Azure Blob Storage rejected the content as too large.{suffix}", itemId, ex),
            // Retrying cannot help a leased blob, so the message says what can.
            409 when ex.ErrorCode is { } code && code.StartsWith("Lease", StringComparison.OrdinalIgnoreCase) =>
                Error(ProviderErrorCode.WriteRejected,
                    "The blob has an active lease, so nothing was stored. Ask whoever holds the lease to release it, " +
                    $"or save as a new document instead.{suffix}", itemId, ex),
            // Any other refusal status is the service declining the request: nothing was stored.
            >= 400 and < 500 => Error(ProviderErrorCode.WriteRejected,
                $"Azure Blob Storage refused the upload with status {ex.Status}; nothing was stored.{suffix}", itemId, ex),
            // A server error, or no response at all, is not a refusal: the blob may have been stored.
            _ => Error(ProviderErrorCode.OutcomeUnknown,
                $"Azure Blob Storage failed the upload with status {ex.Status}; it may have been stored.{suffix}", itemId, ex)
        };
    }

    // ── Reads ─────────────────────────────────────────────────────────────────

    /// <summary>Returns the blob's properties, or <see langword="null"/> when it does not exist.</summary>
    private async Task<BlobProperties?> GetPropertiesAsync(
        string blobName, string? itemId, bool beforeWrite, CancellationToken cancellationToken)
    {
        try
        {
            var response = await _container.GetBlockBlobClient(blobName)
                .GetPropertiesAsync(cancellationToken: cancellationToken).ConfigureAwait(false);
            return response.Value;
        }
        catch (RequestFailedException ex) when (ex.Status == 404)
        {
            return null;
        }
        catch (Exception ex) when (ex is not DocumentProviderException and not OperationCanceledException)
        {
            throw ReadFailure(ex, itemId, beforeWrite);
        }
    }

    private async Task<byte[]> DownloadAsync(string blobName, ETag etag, string itemId, CancellationToken cancellationToken)
    {
        try
        {
            var response = await _container.GetBlockBlobClient(blobName).DownloadStreamingAsync(new BlobDownloadOptions
            {
                Conditions = new BlobRequestConditions { IfMatch = etag }
            }, cancellationToken).ConfigureAwait(false);
            using var result = response.Value;
            return await ObjectKeys.ReadBoundedAsync(result.Content, _options.MaximumBytes, cancellationToken).ConfigureAwait(false)
                ?? throw Error(ProviderErrorCode.ContentTooLarge,
                    $"Document exceeds the configured maximum of {_options.MaximumBytes} bytes.", itemId);
        }
        catch (RequestFailedException ex) when (ex.Status == 412)
        {
            throw new DocumentVersionConflictException(etag.ToString(), "(changed while reading)", ProviderName, ConnectionId, itemId);
        }
        catch (Exception ex) when (ex is not DocumentProviderException and not OperationCanceledException)
        {
            throw ReadFailure(ex, itemId, beforeWrite: false);
        }
    }

    private async Task<ListedObjectPage> ReadListingPageAsync(string prefix, string? continuationToken, CancellationToken cancellationToken)
    {
        var objects = new List<ListedObject>();
        var prefixes = new List<string>();
        string? next = null;
        try
        {
            var pages = _container
                .GetBlobsByHierarchyAsync(new GetBlobsByHierarchyOptions { Delimiter = "/", Prefix = prefix }, cancellationToken)
                .AsPages(continuationToken, CloudListing.ServicePageSize);
            var enumerator = pages.GetAsyncEnumerator(cancellationToken);
            try
            {
                if (await enumerator.MoveNextAsync().ConfigureAwait(false))
                {
                    var page = enumerator.Current;
                    foreach (var item in page.Values)
                    {
                        if (item.IsPrefix)
                            prefixes.Add(item.Prefix);
                        else if (item.IsBlob)
                            objects.Add(new ListedObject
                            {
                                Key = item.Blob.Name,
                                Size = item.Blob.Properties?.ContentLength,
                                LastModified = item.Blob.Properties?.LastModified
                            });
                    }
                    next = page.ContinuationToken;
                }
            }
            finally
            {
                await enumerator.DisposeAsync().ConfigureAwait(false);
            }
        }
        catch (RequestFailedException ex) when (ex.Status == 404)
        {
            throw Error(ProviderErrorCode.NotFound, $"The connection's container does not exist.{CodeSuffix(ex)}", itemId: null, ex);
        }
        catch (Exception ex) when (ex is not DocumentProviderException and not OperationCanceledException)
        {
            throw ReadFailure(ex, itemId: null, beforeWrite: false);
        }
        return new ListedObjectPage { Objects = objects, Prefixes = prefixes, NextToken = string.IsNullOrEmpty(next) ? null : next };
    }

    /// <summary>
    /// Maps a failed read. Before a save writes anything, a failure is certain to have stored
    /// nothing, so an I/O failure there is reported as <see cref="ProviderErrorCode.WriteRejected"/>.
    /// </summary>
    private DocumentProviderException ReadFailure(Exception ex, string? itemId, bool beforeWrite)
    {
        var status = ex is RequestFailedException failed ? failed.Status : 0;
        var suffix = ex is RequestFailedException coded ? CodeSuffix(coded) : string.Empty;
        if (status is 401 or 403)
            return Error(ProviderErrorCode.AccessDenied, $"Azure Blob Storage refused access to the blob.{suffix}", itemId, ex);
        var message = status == 0
            ? "Azure Blob Storage could not be reached."
            : $"Azure Blob Storage request failed with status {status}.{suffix}";
        return beforeWrite
            ? Error(ProviderErrorCode.WriteRejected, message + " Nothing was stored.", itemId, ex)
            : Error(ProviderErrorCode.IO, message, itemId, ex);
    }

    private static string CodeSuffix(RequestFailedException ex) =>
        string.IsNullOrEmpty(ex.ErrorCode) ? string.Empty : $" Storage code: {ex.ErrorCode}.";

    // ── Names and registrations ──────────────────────────────────────────────

    /// <summary>
    /// Resolves a registration source - a blob name relative to the prefix, or the blob's URL in
    /// this connection's container - to a full blob name inside the prefix.
    /// </summary>
    private string ResolveSource(string source)
    {
        if (string.IsNullOrWhiteSpace(source))
            throw Error(ProviderErrorCode.InvalidArgument,
                "An Azure Blob registration requires a blob name relative to the connection's prefix.");

        var trimmed = source.Trim();
        string relative;
        if (Uri.TryCreate(trimmed, UriKind.Absolute, out var uri) &&
            (uri.Scheme == Uri.UriSchemeHttps || uri.Scheme == Uri.UriSchemeHttp))
            relative = RelativeNameFromUrl(uri);
        else if (trimmed.IndexOf("://", StringComparison.Ordinal) >= 0)
            throw Error(ProviderErrorCode.AccessDenied,
                "Only blob names relative to the connection's prefix, or URLs of blobs in its container, can be registered.");
        else
            relative = trimmed;

        CheckRelative(relative, itemId: null);
        return _prefix + relative;
    }

    private string RelativeNameFromUrl(Uri uri)
    {
        const string outside = "The URL does not name a blob inside this connection's container and prefix.";
        var container = new Uri(_container.Uri.GetLeftPart(UriPartial.Path));
        if (!string.Equals(uri.Scheme, container.Scheme, StringComparison.OrdinalIgnoreCase) ||
            !string.Equals(uri.Host, container.Host, StringComparison.OrdinalIgnoreCase) ||
            uri.Port != container.Port)
            throw Error(ProviderErrorCode.AccessDenied, outside);

        var containerPath = container.AbsolutePath.TrimEnd('/') + "/";
        if (!uri.AbsolutePath.StartsWith(containerPath, StringComparison.Ordinal))
            throw Error(ProviderErrorCode.AccessDenied, outside);
        var name = Uri.UnescapeDataString(uri.AbsolutePath.Substring(containerPath.Length));
        if (!name.StartsWith(_prefix, StringComparison.Ordinal))
            throw Error(ProviderErrorCode.AccessDenied, outside);
        return name.Substring(_prefix.Length);
    }

    private void CheckRelative(string relative, string? itemId)
    {
        switch (ObjectKeys.Check(relative))
        {
            case ObjectKeyProblem.Escapes:
                throw Error(ProviderErrorCode.AccessDenied,
                    "Blob names must stay inside the connection's prefix: no leading '/', backslash, '.' or '..' segments.", itemId);
            case ObjectKeyProblem.Invalid:
                throw Error(ProviderErrorCode.InvalidArgument,
                    "A blob name must be 1 to 1024 characters with no empty segments, trailing '/', or control characters.", itemId);
        }
    }

    /// <summary>Revalidates a stored or derived blob name against the current prefix.</summary>
    private void EnsureInsidePrefix(string blobName, string? itemId)
    {
        if (!blobName.StartsWith(_prefix, StringComparison.Ordinal))
            throw Error(ProviderErrorCode.AccessDenied, "The registered blob lies outside this connection's prefix.", itemId);
        CheckRelative(blobName.Substring(_prefix.Length), itemId);
    }

    private async Task<string> ResolveRegistrationAsync(string itemId, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(itemId))
            throw Error(ProviderErrorCode.InvalidArgument, "An item id is required.");
        if (!ObjectKeys.IsOpaqueId(itemId))
            throw Error(ProviderErrorCode.AccessDenied, "Azure Blob item ids are opaque tokens, not blob names.", itemId);

        var blobName = await _store.ResolveAsync(itemId, cancellationToken).ConfigureAwait(false)
            ?? throw Error(ProviderErrorCode.NotFound, $"The provider document '{itemId}' is not registered.", itemId);
        // A stored name does not bypass the boundary: the prefix may have changed since.
        EnsureInsidePrefix(blobName, itemId);
        ValidateExtension(ObjectKeys.Leaf(blobName), itemId);
        return blobName;
    }

    private string RelativeName(string blobName) =>
        blobName.StartsWith(_prefix, StringComparison.Ordinal) ? blobName.Substring(_prefix.Length) : ObjectKeys.Leaf(blobName);

    private void ValidateExtension(string name, string? itemId)
    {
        var extension = ObjectKeys.NormalizeExtension(Path.GetExtension(name));
        if (!_allowedExtensions.Contains(extension))
            throw Error(ProviderErrorCode.ExtensionNotAllowed,
                $"Documents with extension '{extension}' are not allowed by this connection. Allowed: {string.Join(", ", _allowedExtensions)}.",
                itemId);
    }

    private string ValidateName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw Error(ProviderErrorCode.InvalidArgument, "A document name is required.");
        if (name.IndexOfAny(new[] { '/', '\\' }) >= 0 || name == "." || name == ".." || name.Any(char.IsControl))
            throw Error(ProviderErrorCode.InvalidArgument,
                "A document name must be a bare file name without path segments.");
        ValidateExtension(name, itemId: null);
        return name;
    }

    private string ValidateCreateName(string name)
    {
        if (!string.IsNullOrWhiteSpace(name) && PortableDocumentName.ContainsInvalidCharacter(name))
            throw Error(ProviderErrorCode.InvalidArgument,
                "A document name must be a bare file name without path separators, control characters, or other invalid filename characters.");
        if (!string.IsNullOrWhiteSpace(name) && PortableDocumentName.IsWindowsDeviceName(name))
            throw Error(ProviderErrorCode.InvalidArgument, "A document name must not use a reserved Windows device name.");
        if (!string.IsNullOrWhiteSpace(name) &&
            (!string.Equals(name, name.Trim(), StringComparison.Ordinal) || name.EndsWith(".", StringComparison.Ordinal)))
            throw Error(ProviderErrorCode.InvalidArgument,
                "A document name must not begin or end with whitespace or end with a dot.");
        return ValidateName(name);
    }

    private void ValidateReference(DocumentReference reference)
    {
        if (reference is null) throw new ArgumentNullException(nameof(reference));
        if (!string.Equals(reference.Provider, ProviderName, StringComparison.OrdinalIgnoreCase) ||
            !string.Equals(reference.ConnectionId, ConnectionId, StringComparison.Ordinal))
            throw new DocumentProviderException(
                ProviderErrorCode.InvalidArgument,
                $"The document reference (provider='{reference.Provider}', connectionId='{reference.ConnectionId}') does not belong to this Azure Blob connection (provider='{ProviderName}', connectionId='{ConnectionId}').",
                ProviderName, ConnectionId, reference.ItemId);
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private static BlobRequestConditions CreateOnly() => new() { IfNoneMatch = ETag.All };

    private static BlobHttpHeaders NewHeaders(string fileName) => new() { ContentType = OfficeContentTypes.ForName(fileName) };

    private static BlobHttpHeaders CarriedHeaders(BlobProperties properties, string fileName) => new()
    {
        ContentType = string.IsNullOrEmpty(properties.ContentType) ? OfficeContentTypes.ForName(fileName) : properties.ContentType,
        ContentLanguage = properties.ContentLanguage,
        ContentDisposition = properties.ContentDisposition,
        CacheControl = properties.CacheControl
    };

    private static IDictionary<string, string> CarriedMetadata(BlobProperties properties) =>
        properties.Metadata is null
            ? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            : new Dictionary<string, string>(properties.Metadata, StringComparer.OrdinalIgnoreCase);

    private async Task<byte[]> ReadContentAsync(Stream content, string? itemId, CancellationToken cancellationToken)
    {
        if (content.CanSeek) content.Position = 0;
        return await ObjectKeys.ReadBoundedAsync(content, _options.MaximumBytes, cancellationToken).ConfigureAwait(false)
            ?? throw Error(ProviderErrorCode.ContentTooLarge,
                $"Document exceeds the configured maximum of {_options.MaximumBytes} bytes.", itemId);
    }

    private DocumentReference CreateReference(string registrationId, string blobName, string etag)
    {
        var name = ObjectKeys.Leaf(blobName);
        return new DocumentReference
        {
            Provider = ProviderName,
            ConnectionId = ConnectionId,
            ItemId = registrationId,
            Version = etag,
            Name = name,
            ContentType = OfficeContentTypes.ForName(name)
        };
    }

    private DocumentRegistrationFailedException RegistrationFailed(
        string message, string? itemId, string outputName, byte[] bytes, Exception innerException) =>
        new(message, ProviderName, ConnectionId, itemId, outputName, ObjectKeys.Sha256(bytes), innerException);

    private DocumentProviderException Error(
        ProviderErrorCode code, string message, string? itemId = null, Exception? innerException = null) =>
        new(code, message, ProviderName, ConnectionId, itemId, innerException);

    private static BlobContainerClient CreateContainerClient(AzureBlobDocumentProviderOptions options)
    {
        if (options is null) throw new ArgumentNullException(nameof(options));
        if (string.IsNullOrWhiteSpace(options.ContainerName))
            throw new ArgumentException("An Azure Blob container name is required.", nameof(options));

        var hasConnectionString = !string.IsNullOrWhiteSpace(options.ConnectionString);
        var hasServiceUri = !string.IsNullOrWhiteSpace(options.ServiceUri);
        if (hasConnectionString == hasServiceUri)
            throw new ArgumentException(
                "Configure exactly one of ServiceUri (with DefaultAzureCredential) or ConnectionString.", nameof(options));

        if (hasConnectionString)
        {
            try
            {
                return new BlobContainerClient(options.ConnectionString, options.ContainerName);
            }
            catch (Exception ex) when (ex is FormatException or ArgumentException)
            {
                // The parser's message could quote the connection string, which carries the key.
                throw new ArgumentException("The Azure Blob connection string is not valid.", nameof(options));
            }
        }

        if (!Uri.TryCreate(options.ServiceUri.Trim(), UriKind.Absolute, out var service) ||
            (service.Scheme != Uri.UriSchemeHttps && service.Scheme != Uri.UriSchemeHttp))
            throw new ArgumentException("ServiceUri must be an absolute http(s) URL.", nameof(options));

        var credentialOptions = new DefaultAzureCredentialOptions();
        if (!string.IsNullOrWhiteSpace(options.ManagedIdentityClientId))
            credentialOptions.ManagedIdentityClientId = options.ManagedIdentityClientId.Trim();
        var containerUri = new Uri(service.GetLeftPart(UriPartial.Path).TrimEnd('/') + "/" + Uri.EscapeDataString(options.ContainerName.Trim()));
        return new BlobContainerClient(containerUri, new DefaultAzureCredential(credentialOptions));
    }
}
