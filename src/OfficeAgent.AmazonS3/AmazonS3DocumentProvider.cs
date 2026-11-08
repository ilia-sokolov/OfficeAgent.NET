using System.Net;
using Amazon;
using Amazon.Runtime;
using Amazon.S3;
using Amazon.S3.Model;
using OfficeAgent.Abstractions;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.ObjectStorage;

namespace OfficeAgent.AmazonS3;

/// <summary>
/// Amazon S3 document provider: a registry of references to objects in one bucket, under an
/// optional key prefix, addressed by an opaque item id assigned at registration. Documents are
/// registered by a key relative to the prefix (or <c>s3://bucket/key</c> for the same bucket); a
/// key that leaves the prefix is refused with <see cref="ProviderErrorCode.AccessDenied"/>.
/// </summary>
/// <remarks>
/// Saves use S3 conditional writes. <see cref="SaveMode.Replace"/> sends <c>If-Match</c> with the
/// ETag that was checked, and every new object - a versioned sibling, a renamed copy, a created
/// document - is written with <c>If-None-Match: *</c>, so the provider never overwrites content it
/// did not open. When the bucket has versioning enabled, a reference's version is the S3 version
/// id of the object it read or wrote; otherwise it is the ETag. Each write carries a random
/// <c>officeagent-write-id</c> metadata value, so a precondition failure caused by the SDK
/// retrying a request whose first attempt landed is recognised as this provider's own write.
/// Removing a registration never deletes the object.
/// </remarks>
public sealed class AmazonS3DocumentProvider : IDocumentProvider, IDocumentCreatingProvider, IConnectionEditingDefaults, IConnectionGranularityDefaults, IConnectionCalculationDefaults
{
    /// <summary>Gets the provider discriminator used in document references.</summary>
    public const string ProviderName = "amazons3";

    /// <summary>The metadata entry each write carries so a retried write can recognise itself.</summary>
    internal const string WriteMarker = "officeagent-write-id";

    private const int NameRaceAttempts = 5;

    private readonly AmazonS3DocumentProviderOptions _options;
    private readonly IAmazonS3 _s3;
    private readonly IAmazonS3RegistrationStore _store;
    private readonly string _bucket;
    private readonly string _prefix;
    private readonly HashSet<string> _allowedExtensions;

    /// <summary>
    /// Initializes a provider from configuration. Credentials come from the AWS SDK's default
    /// chain: an instance, task or pod role, environment variables, or a shared profile.
    /// </summary>
    public AmazonS3DocumentProvider(
        AmazonS3DocumentProviderOptions options,
        IAmazonS3RegistrationStore? registrationStore = null)
        : this(options, CreateClient(options), registrationStore)
    {
    }

    /// <summary>
    /// Initializes a provider over an S3 client the host built, for custom credentials, endpoint
    /// or retry settings. <see cref="AmazonS3DocumentProviderOptions.Region"/>,
    /// <see cref="AmazonS3DocumentProviderOptions.ServiceUrl"/> and
    /// <see cref="AmazonS3DocumentProviderOptions.ForcePathStyle"/> are not used.
    /// </summary>
    public AmazonS3DocumentProvider(
        AmazonS3DocumentProviderOptions options,
        IAmazonS3 client,
        IAmazonS3RegistrationStore? registrationStore = null)
    {
        _options = options ?? throw new ArgumentNullException(nameof(options));
        _s3 = client ?? throw new ArgumentNullException(nameof(client));
        if (string.IsNullOrWhiteSpace(options.ConnectionId))
            throw new ArgumentException("An Amazon S3 provider connection id is required.", nameof(options));
        if (string.IsNullOrWhiteSpace(options.BucketName))
            throw new ArgumentException("An Amazon S3 bucket name is required.", nameof(options));
        if (options.MaximumBytes <= 0)
            throw new ArgumentOutOfRangeException(nameof(options), "MaximumBytes must be positive.");
        if (options.AllowedExtensions is null || options.AllowedExtensions.Count == 0)
            throw new ArgumentException("At least one allowed document extension is required.", nameof(options));

        _bucket = options.BucketName.Trim();
        _prefix = ObjectKeys.NormalizePrefix(options.KeyPrefix, nameof(options));
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
    public async Task<DocumentReference> RegisterAsync(
        string source,
        CancellationToken cancellationToken = default)
    {
        var key = ResolveSource(source);
        ValidateExtension(ObjectKeys.Leaf(key), itemId: null);

        var head = await HeadAsync(key, itemId: null, beforeWrite: false, cancellationToken).ConfigureAwait(false)
            ?? throw Error(ProviderErrorCode.NotFound, $"No object with key '{RelativeKey(key)}' exists in this connection.");
        if (head.ContentLength > _options.MaximumBytes)
            throw Error(ProviderErrorCode.ContentTooLarge,
                $"Document '{RelativeKey(key)}' exceeds the configured maximum of {_options.MaximumBytes} bytes.");

        var id = await _store.AddAsync(key, cancellationToken).ConfigureAwait(false);
        return CreateReference(id, key, VersionOf(head.VersionId, head.ETag));
    }

    /// <inheritdoc />
    public async Task<DocumentReference> CreateAsync(
        string name,
        Stream content,
        CancellationToken cancellationToken = default)
    {
        if (content is null) throw new ArgumentNullException(nameof(content));

        var fileName = ValidateCreateName(name);
        var key = _prefix + fileName;
        var bytes = await ReadContentAsync(content, itemId: null, cancellationToken).ConfigureAwait(false);

        string version;
        try
        {
            version = await WriteAsync(NewObject(key, bytes), itemId: null, replace: false, cancellationToken).ConfigureAwait(false);
        }
        catch (DocumentProviderException ex) when (ex.Code == ProviderErrorCode.AlreadyExists)
        {
            throw Error(ProviderErrorCode.AlreadyExists,
                $"A document named '{fileName}' already exists in this connection; choose another name.", itemId: null, ex);
        }

        string id;
        try
        {
            // The object is stored. Finish registering even if the caller cancels now.
            id = await _store.AddAsync(key, CancellationToken.None).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            throw RegistrationFailed(
                $"Document '{fileName}' was created in Amazon S3, but its registration could not be persisted. " +
                "The object is stored but unregistered. Do not retry creation; register it by key instead.",
                itemId: null, fileName, bytes, ex);
        }
        return CreateReference(id, key, version);
    }

    /// <inheritdoc />
    public async Task<DocumentContent> OpenReadAsync(
        DocumentReference reference,
        CancellationToken cancellationToken = default)
    {
        ValidateReference(reference);
        var key = await ResolveRegistrationAsync(reference.ItemId, cancellationToken).ConfigureAwait(false);

        var head = await HeadAsync(key, reference.ItemId, beforeWrite: false, cancellationToken).ConfigureAwait(false)
            ?? throw Error(ProviderErrorCode.NotFound, "The referenced object no longer exists.", reference.ItemId);
        if (head.ContentLength > _options.MaximumBytes)
            throw Error(ProviderErrorCode.ContentTooLarge,
                $"Document exceeds the configured maximum of {_options.MaximumBytes} bytes.", reference.ItemId);

        var version = VersionOf(head.VersionId, head.ETag);
        if (!string.IsNullOrEmpty(reference.Version) && !string.Equals(reference.Version, version, StringComparison.Ordinal))
            throw new DocumentVersionConflictException(reference.Version!, version, ProviderName, ConnectionId, reference.ItemId);

        var bytes = await DownloadAsync(key, head, reference.ItemId, cancellationToken).ConfigureAwait(false);
        return new DocumentContent(CreateReference(reference.ItemId, key, version), new MemoryStream(bytes, writable: false));
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
                "Amazon S3 item ids are opaque and provider-assigned; a destination id cannot be chosen by the caller. Use Mode=NewVersion/NewDocument (optionally with NewName) to mint a new document, or Mode=Replace to overwrite the source.",
                source.ItemId);
        if (options.Mode == SaveMode.Replace && !string.IsNullOrWhiteSpace(options.NewName))
            throw Error(ProviderErrorCode.InvalidArgument,
                "Replace mode overwrites the source in place and cannot rename. Use Mode=NewVersion/NewDocument with NewName to mint a renamed document.",
                source.ItemId);

        var key = await ResolveRegistrationAsync(source.ItemId, cancellationToken).ConfigureAwait(false);
        var head = await HeadAsync(key, source.ItemId, beforeWrite: true, cancellationToken).ConfigureAwait(false)
            ?? throw Error(ProviderErrorCode.NotFound, "The referenced object no longer exists.", source.ItemId);

        var currentVersion = VersionOf(head.VersionId, head.ETag);
        var expectedVersion = options.ExpectedVersion ?? source.Version;
        if (!string.IsNullOrEmpty(expectedVersion) && !string.Equals(expectedVersion, currentVersion, StringComparison.Ordinal))
            throw new DocumentVersionConflictException(expectedVersion!, currentVersion, ProviderName, ConnectionId, source.ItemId);

        var bytes = await ReadContentAsync(content, source.ItemId, cancellationToken).ConfigureAwait(false);

        if (options.Mode == SaveMode.Replace)
        {
            // Conditioned on the ETag just checked, so a change between the check and the write
            // is a conflict, not a silent overwrite. Content type, metadata and encryption carry over.
            var version = await WriteAsync(ReplacingObject(key, bytes, head), source.ItemId, replace: true, cancellationToken)
                .ConfigureAwait(false);
            return CreateReference(source.ItemId, key, version);
        }

        var (destination, savedVersion) = string.IsNullOrWhiteSpace(options.NewName)
            ? await WriteNextVersionAsync(key, bytes, source.ItemId, cancellationToken).ConfigureAwait(false)
            : await WriteNamedAsync(ObjectKeys.Parent(key) + ValidateName(options.NewName!), bytes, source.ItemId, cancellationToken).ConfigureAwait(false);

        string newId;
        try
        {
            // The object is stored. Finish registering even if the caller cancels now.
            newId = await _store.AddAsync(destination, CancellationToken.None).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            throw RegistrationFailed(
                $"Document '{RelativeKey(destination)}' was saved to Amazon S3, but its registration could not be " +
                "persisted. Do not save it again; register the saved object by key to recover it.",
                source.ItemId, RelativeKey(destination), bytes, ex);
        }
        return CreateReference(newId, destination, savedVersion);
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

    // ── Writes ────────────────────────────────────────────────────────────────

    private async Task<(string Key, string Version)> WriteNamedAsync(
        string key, byte[] bytes, string itemId, CancellationToken cancellationToken)
    {
        EnsureInsidePrefix(key, itemId);
        var version = await WriteAsync(NewObject(key, bytes), itemId, replace: false, cancellationToken).ConfigureAwait(false);
        return (key, version);
    }

    /// <summary>
    /// Writes the next free versioned sibling (<c>contract.v2.docx</c>, …). Each candidate is
    /// probed first so a taken name costs a HEAD, not an upload; a name another writer takes
    /// between the probe and the write fails the create-only PUT and the next one is tried.
    /// </summary>
    private async Task<(string Key, string Version)> WriteNextVersionAsync(
        string sourceKey, byte[] bytes, string itemId, CancellationToken cancellationToken)
    {
        var parent = ObjectKeys.Parent(sourceKey);
        var races = 0;
        foreach (var candidate in ObjectKeys.VersionedNames(ObjectKeys.Leaf(sourceKey)))
        {
            var key = parent + candidate;
            if (key.Length - _prefix.Length > ObjectKeys.MaximumKeyLength) break;
            if (await HeadAsync(key, itemId, beforeWrite: true, cancellationToken).ConfigureAwait(false) is not null)
                continue;
            try
            {
                return await WriteNamedAsync(key, bytes, itemId, cancellationToken).ConfigureAwait(false);
            }
            catch (DocumentProviderException ex) when (ex.Code == ProviderErrorCode.AlreadyExists && ++races < NameRaceAttempts)
            {
            }
        }
        throw Error(ProviderErrorCode.AlreadyExists,
            "No free versioned name is left for this document; save with NewName instead.", itemId);
    }

    private PutObjectRequest NewObject(string key, byte[] bytes) => new()
    {
        BucketName = _bucket,
        Key = key,
        InputStream = new MemoryStream(bytes, writable: false),
        ContentType = OfficeContentTypes.ForName(ObjectKeys.Leaf(key)),
        IfNoneMatch = "*"
    };

    private PutObjectRequest ReplacingObject(string key, byte[] bytes, GetObjectMetadataResponse head)
    {
        var request = new PutObjectRequest
        {
            BucketName = _bucket,
            Key = key,
            InputStream = new MemoryStream(bytes, writable: false),
            ContentType = string.IsNullOrEmpty(head.Headers?.ContentType)
                ? OfficeContentTypes.ForName(ObjectKeys.Leaf(key))
                : head.Headers!.ContentType,
            IfMatch = head.ETag
        };
        if (head.Metadata is not null)
            foreach (var name in head.Metadata.Keys)
                if (!IsWriteMarker(name))
                    request.Metadata.Add(name, head.Metadata[name]);
        // A PUT replaces the object's encryption settings too; keep a KMS key the object used.
        if (head.ServerSideEncryptionMethod is { } method && method != ServerSideEncryptionMethod.None)
        {
            request.ServerSideEncryptionMethod = method;
            if (!string.IsNullOrEmpty(head.ServerSideEncryptionKeyManagementServiceKeyId))
                request.ServerSideEncryptionKeyManagementServiceKeyId = head.ServerSideEncryptionKeyManagementServiceKeyId;
        }
        return request;
    }

    private static bool IsWriteMarker(string metadataName) =>
        metadataName.EndsWith(WriteMarker, StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Writes one object under its precondition and returns the resulting version. A refusal S3
    /// reports before storing anything is certain; a server error, a timeout or a lost response is
    /// an unknown outcome and is never reported as nothing written.
    /// </summary>
    private async Task<string> WriteAsync(
        PutObjectRequest request, string? itemId, bool replace, CancellationToken cancellationToken)
    {
        var marker = Guid.NewGuid().ToString("N");
        request.Metadata.Add(WriteMarker, marker);
        request.AutoCloseStream = true;
        try
        {
            var response = await _s3.PutObjectAsync(request, cancellationToken).ConfigureAwait(false);
            return VersionOf(response.VersionId, response.ETag);
        }
        catch (AmazonS3Exception ex) when (IsPreconditionFailure(ex))
        {
            // The SDK retries timeouts and server errors. If an earlier attempt stored the object,
            // the retry fails its own precondition: recognise that write before calling it a conflict.
            var landed = await FindOwnWriteAsync(request.Key, marker, itemId, cancellationToken).ConfigureAwait(false);
            if (landed is not null) return landed;
            if (replace)
                throw new DocumentVersionConflictException("(If-Match)", "(changed)", ProviderName, ConnectionId, itemId);
            if (ex.StatusCode == HttpStatusCode.Conflict)
                // S3 reports a concurrent delete or conditional write this way; nothing was stored.
                throw Error(ProviderErrorCode.WriteRejected,
                    $"Amazon S3 refused the write because of a concurrent request on the same key; nothing was stored. Try again.{CodeSuffix(ex)}",
                    itemId, ex);
            throw Error(ProviderErrorCode.AlreadyExists,
                $"An object with key '{RelativeKey(request.Key)}' already exists; refusing to overwrite it.", itemId);
        }
        catch (AmazonServiceException ex) when (ex.StatusCode != 0)
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
                "The write to Amazon S3 failed before a response arrived; it may have been stored.", itemId, ex);
        }
    }

    private static bool IsPreconditionFailure(AmazonS3Exception ex) =>
        ex.StatusCode == HttpStatusCode.PreconditionFailed ||
        (ex.StatusCode == HttpStatusCode.Conflict &&
         string.Equals(ex.ErrorCode, "ConditionalRequestConflict", StringComparison.OrdinalIgnoreCase));

    private async Task<string?> FindOwnWriteAsync(string key, string marker, string? itemId, CancellationToken cancellationToken)
    {
        try
        {
            var head = await _s3.GetObjectMetadataAsync(new GetObjectMetadataRequest { BucketName = _bucket, Key = key }, cancellationToken)
                .ConfigureAwait(false);
            return head.Metadata is not null && string.Equals(head.Metadata[WriteMarker], marker, StringComparison.Ordinal)
                ? VersionOf(head.VersionId, head.ETag)
                : null;
        }
        catch (AmazonServiceException ex) when (ex.StatusCode == HttpStatusCode.NotFound)
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
                "Amazon S3 refused the write's precondition and the object could not be read back " +
                "to tell whether an earlier attempt stored it.", itemId, ex);
        }
    }

    private DocumentProviderException WriteFailure(AmazonServiceException ex, string? itemId)
    {
        var status = (int)ex.StatusCode;
        var suffix = CodeSuffix(ex);
        return status switch
        {
            404 => Error(ProviderErrorCode.NotFound,
                $"The object or its bucket no longer exists; nothing was stored.{suffix}", itemId, ex),
            401 or 403 => Error(ProviderErrorCode.AccessDenied,
                $"Amazon S3 refused access to the object; nothing was stored.{suffix}", itemId, ex),
            413 => Error(ProviderErrorCode.ContentTooLarge,
                $"Amazon S3 rejected the content as too large.{suffix}", itemId, ex),
            // Any other refusal status is S3 declining the request: nothing was stored.
            >= 400 and < 500 => Error(ProviderErrorCode.WriteRejected,
                $"Amazon S3 refused the write with status {status}; nothing was stored.{suffix}", itemId, ex),
            // A server error (including 503 SlowDown) is not proof of a refusal: it may have been stored.
            _ => Error(ProviderErrorCode.OutcomeUnknown,
                $"Amazon S3 failed the write with status {status}; it may have been stored.{suffix}", itemId, ex)
        };
    }

    // ── Reads ─────────────────────────────────────────────────────────────────

    /// <summary>Returns the object's metadata, or <see langword="null"/> when it does not exist.</summary>
    private async Task<GetObjectMetadataResponse?> HeadAsync(
        string key, string? itemId, bool beforeWrite, CancellationToken cancellationToken)
    {
        try
        {
            return await _s3.GetObjectMetadataAsync(new GetObjectMetadataRequest { BucketName = _bucket, Key = key }, cancellationToken)
                .ConfigureAwait(false);
        }
        catch (AmazonServiceException ex) when (ex.StatusCode == HttpStatusCode.NotFound)
        {
            return null;
        }
        catch (Exception ex) when (ex is not DocumentProviderException and not OperationCanceledException)
        {
            throw ReadFailure(ex, itemId, beforeWrite);
        }
    }

    private async Task<byte[]> DownloadAsync(string key, GetObjectMetadataResponse head, string itemId, CancellationToken cancellationToken)
    {
        // Pinned to what was just checked: the version id when versioning is on, the ETag otherwise.
        var request = new GetObjectRequest { BucketName = _bucket, Key = key };
        if (IsRealVersionId(head.VersionId))
            request.VersionId = head.VersionId;
        else
            request.EtagToMatch = head.ETag;
        try
        {
            using var response = await _s3.GetObjectAsync(request, cancellationToken).ConfigureAwait(false);
            return await ObjectKeys.ReadBoundedAsync(response.ResponseStream, _options.MaximumBytes, cancellationToken).ConfigureAwait(false)
                ?? throw Error(ProviderErrorCode.ContentTooLarge,
                    $"Document exceeds the configured maximum of {_options.MaximumBytes} bytes.", itemId);
        }
        catch (AmazonServiceException ex) when (ex.StatusCode == HttpStatusCode.PreconditionFailed)
        {
            throw new DocumentVersionConflictException(head.ETag ?? "(unknown)", "(changed while reading)", ProviderName, ConnectionId, itemId);
        }
        catch (Exception ex) when (ex is not DocumentProviderException and not OperationCanceledException)
        {
            throw ReadFailure(ex, itemId, beforeWrite: false);
        }
    }

    /// <summary>
    /// Maps a failed read. Before a save writes anything, a failure is certain to have stored
    /// nothing, so an I/O failure there is reported as <see cref="ProviderErrorCode.WriteRejected"/>.
    /// </summary>
    private DocumentProviderException ReadFailure(Exception ex, string? itemId, bool beforeWrite)
    {
        var service = ex as AmazonServiceException;
        var status = service is null ? 0 : (int)service.StatusCode;
        var suffix = service is null ? string.Empty : CodeSuffix(service);
        if (status is 401 or 403)
            return Error(ProviderErrorCode.AccessDenied, $"Amazon S3 refused access to the object.{suffix}", itemId, ex);
        var message = status == 0
            ? "Amazon S3 could not be reached."
            : $"Amazon S3 request failed with status {status}.{suffix}";
        return beforeWrite
            ? Error(ProviderErrorCode.WriteRejected, message + " Nothing was stored.", itemId, ex)
            : Error(ProviderErrorCode.IO, message, itemId, ex);
    }

    private static string CodeSuffix(AmazonServiceException ex) =>
        string.IsNullOrEmpty(ex.ErrorCode) ? string.Empty : $" S3 code: {ex.ErrorCode}.";

    // ── Names and registrations ──────────────────────────────────────────────

    /// <summary>
    /// Resolves a registration source - a key relative to the prefix, or <c>s3://bucket/key</c>
    /// for this connection's bucket - to a full key inside the prefix.
    /// </summary>
    private string ResolveSource(string source)
    {
        if (string.IsNullOrWhiteSpace(source))
            throw Error(ProviderErrorCode.InvalidArgument,
                "An Amazon S3 registration requires an object key relative to the connection's prefix.");

        var trimmed = source.Trim();
        string relative;
        const string scheme = "s3://";
        if (trimmed.StartsWith(scheme, StringComparison.OrdinalIgnoreCase))
        {
            var rest = trimmed.Substring(scheme.Length);
            var slash = rest.IndexOf('/');
            var bucket = slash < 0 ? rest : rest.Substring(0, slash);
            var key = slash < 0 ? string.Empty : rest.Substring(slash + 1);
            if (!string.Equals(bucket, _bucket, StringComparison.Ordinal) || !key.StartsWith(_prefix, StringComparison.Ordinal))
                throw Error(ProviderErrorCode.AccessDenied,
                    "The URI does not name an object inside this connection's bucket and prefix.");
            relative = key.Substring(_prefix.Length);
        }
        else if (trimmed.IndexOf("://", StringComparison.Ordinal) >= 0)
            throw Error(ProviderErrorCode.AccessDenied,
                "Only keys relative to the connection's prefix, or s3:// URIs in its bucket, can be registered.");
        else
            relative = trimmed;

        CheckRelative(relative, itemId: null);
        return _prefix + relative;
    }

    private void CheckRelative(string relative, string? itemId)
    {
        switch (ObjectKeys.Check(relative))
        {
            case ObjectKeyProblem.Escapes:
                throw Error(ProviderErrorCode.AccessDenied,
                    "Object keys must stay inside the connection's prefix: no leading '/', backslash, '.' or '..' segments.", itemId);
            case ObjectKeyProblem.Invalid:
                throw Error(ProviderErrorCode.InvalidArgument,
                    "An object key must be 1 to 1024 characters with no empty segments, trailing '/', or control characters.", itemId);
        }
    }

    /// <summary>Revalidates a stored or derived key against the current prefix.</summary>
    private void EnsureInsidePrefix(string key, string? itemId)
    {
        if (!key.StartsWith(_prefix, StringComparison.Ordinal))
            throw Error(ProviderErrorCode.AccessDenied, "The registered object lies outside this connection's prefix.", itemId);
        CheckRelative(key.Substring(_prefix.Length), itemId);
    }

    private async Task<string> ResolveRegistrationAsync(string itemId, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(itemId))
            throw Error(ProviderErrorCode.InvalidArgument, "An item id is required.");
        if (!ObjectKeys.IsOpaqueId(itemId))
            throw Error(ProviderErrorCode.AccessDenied, "Amazon S3 item ids are opaque tokens, not object keys.", itemId);

        var key = await _store.ResolveAsync(itemId, cancellationToken).ConfigureAwait(false)
            ?? throw Error(ProviderErrorCode.NotFound, $"The provider document '{itemId}' is not registered.", itemId);
        // A stored key does not bypass the boundary: the prefix may have changed since.
        EnsureInsidePrefix(key, itemId);
        ValidateExtension(ObjectKeys.Leaf(key), itemId);
        return key;
    }

    private string RelativeKey(string key) =>
        key.StartsWith(_prefix, StringComparison.Ordinal) ? key.Substring(_prefix.Length) : ObjectKeys.Leaf(key);

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
                $"The document reference (provider='{reference.Provider}', connectionId='{reference.ConnectionId}') does not belong to this Amazon S3 connection (provider='{ProviderName}', connectionId='{ConnectionId}').",
                ProviderName, ConnectionId, reference.ItemId);
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    /// <summary>
    /// The version a reference carries: the S3 version id on a versioning-enabled bucket, so it
    /// names exactly the bytes read or written, and the ETag otherwise.
    /// </summary>
    private static string VersionOf(string? versionId, string? etag) =>
        IsRealVersionId(versionId) ? versionId! : etag ?? string.Empty;

    // An unversioned or version-suspended bucket reports no version id, or the literal "null".
    private static bool IsRealVersionId(string? versionId) =>
        !string.IsNullOrEmpty(versionId) && !string.Equals(versionId, "null", StringComparison.Ordinal);

    private async Task<byte[]> ReadContentAsync(Stream content, string? itemId, CancellationToken cancellationToken)
    {
        if (content.CanSeek) content.Position = 0;
        return await ObjectKeys.ReadBoundedAsync(content, _options.MaximumBytes, cancellationToken).ConfigureAwait(false)
            ?? throw Error(ProviderErrorCode.ContentTooLarge,
                $"Document exceeds the configured maximum of {_options.MaximumBytes} bytes.", itemId);
    }

    private DocumentReference CreateReference(string registrationId, string key, string version)
    {
        var name = ObjectKeys.Leaf(key);
        return new DocumentReference
        {
            Provider = ProviderName,
            ConnectionId = ConnectionId,
            ItemId = registrationId,
            Version = version,
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

    private static IAmazonS3 CreateClient(AmazonS3DocumentProviderOptions options)
    {
        if (options is null) throw new ArgumentNullException(nameof(options));
        var config = new AmazonS3Config { ForcePathStyle = options.ForcePathStyle };
        var region = string.IsNullOrWhiteSpace(options.Region) ? null : options.Region.Trim();
        if (!string.IsNullOrWhiteSpace(options.ServiceUrl))
        {
            config.ServiceURL = options.ServiceUrl.Trim();
            if (region is not null) config.AuthenticationRegion = region;
        }
        else if (region is not null)
        {
            config.RegionEndpoint = RegionEndpoint.GetBySystemName(region);
        }
        // No credentials are passed: the client resolves them through the SDK's default chain.
        return new AmazonS3Client(config);
    }
}
