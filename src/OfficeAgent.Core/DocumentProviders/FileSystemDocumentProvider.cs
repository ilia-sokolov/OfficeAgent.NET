using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Serialization;
using OfficeAgent.Abstractions;

namespace OfficeAgent.Core.DocumentProviders;

/// <summary>Configures one rooted filesystem document-provider connection.</summary>
public sealed class FileSystemDocumentProviderOptions
{
    /// <summary>Gets or sets the provider connection identifier.</summary>
    public string ConnectionId { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the filesystem root visible through this connection. Registrations
    /// must point at a file under this root (relative paths are resolved against it,
    /// absolute paths are rejected when they escape it). New revisions written by
    /// <see cref="FileSystemDocumentProvider.SaveAsync"/> in NewVersion/NewDocument
    /// mode are placed beside the source file under this root.
    /// The root is a host-managed trust boundary: untrusted operating-system
    /// principals must not be able to create, rename, or replace directory entries
    /// beneath it. The provider rejects traversal and reparse points before each
    /// access, but path validation and I/O are not one atomic filesystem operation.
    /// </summary>
    public string RootPath { get; set; } = string.Empty;

    /// <summary>Gets or sets the maximum accepted document size in bytes.</summary>
    public long MaximumBytes { get; set; } = 100 * 1024 * 1024;

    /// <summary>Gets or sets the allowed document extensions.</summary>
    public IReadOnlyCollection<string> AllowedExtensions { get; set; } = new[] { ".docx" };

    /// <summary>
    /// Gets or sets the change mode used when an operation does not state one. The
    /// default is <see cref="ChangeMode.Tracked"/>, which suits documents a person
    /// reviews. Set <see cref="ChangeMode.Direct"/> for a connection over generated or
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
    /// Gets or sets the folder, relative to <see cref="RootPath"/>, that holds this connection's
    /// templates, such as <c>Templates</c>. It must stay inside the root. Listing with
    /// <see cref="DocumentListQuery.Templates"/> lists it, and every document in it is reported as
    /// a template. Empty for none.
    /// </summary>
    public string? TemplatesFolder { get; set; }

    /// <summary>
    /// Gets or sets what kind of template location the connection is. Left at
    /// <see cref="TemplateLocationKind.None"/>, it becomes <see cref="TemplateLocationKind.Connection"/>
    /// when <see cref="TemplatesFolder"/> is set. Any other value without a
    /// <see cref="TemplatesFolder"/> makes the whole root the templates folder, as for Office's
    /// personal and workgroup template folders.
    /// </summary>
    public TemplateLocationKind TemplateLocation { get; set; }

    /// <summary>
    /// Gets or sets whether the connection refuses every write: saving, creating and removing a
    /// registration fail with <see cref="ProviderErrorCode.AccessDenied"/> before anything is
    /// changed. A read-only connection also creates no directory and keeps its registrations in
    /// memory instead of writing <c>.officeagent/index.json</c> under the root, so a templates
    /// folder people share is never written to. Defaults to <see langword="false"/>.
    /// </summary>
    public bool ReadOnly { get; set; }
}

/// <summary>
/// Filesystem document provider: a registry of references to files the host already
/// owns, addressed by an opaque item id assigned at registration time. The provider
/// never copies, mirrors, or owns content - it only persists the id → path mapping
/// under <c>{root}/.officeagent/index.json</c> and routes open/save back to the
/// referenced path. Default edit tools use opaque ids; opt-in registration tools may
/// accept a connection-relative source path. New documents created through
/// <see cref="CreateAsync"/> are the one case
/// where the provider owns the bytes it writes: they land directly under the root.
/// </summary>
/// <remarks>
/// <see cref="ListAsync"/> lists documents under the root without registering them. It never
/// leaves the root: links, junctions and other reparse points are skipped rather than followed,
/// and hidden or system entries, dot-prefixed names (the index directory and staging files) and
/// Office lock files (<c>~$*</c>) are not listed. A recursive listing descends at most
/// <see cref="MaximumListingDepth"/> folders below the one listed, examines at most
/// <see cref="MaximumEntriesPerListing"/> entries per call (a page cut short by that bound carries
/// a continuation token), and reads at most <see cref="MaximumEntriesPerFolder"/> entries of any
/// one folder, in name order.
/// </remarks>
public sealed class FileSystemDocumentProvider : IDocumentProvider, IDocumentCreatingProvider, IDocumentListingProvider, IConnectionTemplateSource, IConnectionEditingDefaults, IConnectionGranularityDefaults, IConnectionCalculationDefaults
{
    /// <summary>Gets the provider discriminator used in document references.</summary>
    public const string ProviderName = "filesystem";

    /// <summary>How many folders below the listed one a recursive listing descends.</summary>
    public const int MaximumListingDepth = 8;

    /// <summary>How many files and folders one listing call examines before it returns a page.</summary>
    public const int MaximumEntriesPerListing = 5000;

    /// <summary>How many entries of one folder a listing reads; the rest of a larger folder is not listed.</summary>
    public const int MaximumEntriesPerFolder = 10000;

    private const string IndexDirectoryName = ".officeagent";
    private const string IndexFileName = "index.json";
    private const string TokenVersion = "fs1";

    private readonly string _root;
    private readonly string _rootWithSeparator;
    private readonly string? _templatesDirectory;
    private readonly string _indexPath;
    private readonly long _maximumBytes;
    private readonly HashSet<string> _allowedExtensions;
    private readonly SemaphoreSlim _indexLock = new(1, 1);
    // One gate per item, so a contested write to one document serializes its own
    // check-and-publish while writes to other documents proceed in parallel. A single
    // shared gate would make every save in the connection queue behind every other.
    private readonly Dictionary<string, SemaphoreSlim> _itemLocks = new(StringComparer.Ordinal);
    private readonly Dictionary<string, string> _index;

    /// <summary>Initializes a rooted filesystem provider.</summary>
    public FileSystemDocumentProvider(FileSystemDocumentProviderOptions options)
    {
        if (options is null) throw new ArgumentNullException(nameof(options));
        if (string.IsNullOrWhiteSpace(options.ConnectionId))
            throw new ArgumentException("A filesystem provider connection id is required.", nameof(options));
        if (string.IsNullOrWhiteSpace(options.RootPath))
            throw new ArgumentException("A filesystem provider root path is required.", nameof(options));
        if (options.MaximumBytes <= 0)
            throw new ArgumentOutOfRangeException(nameof(options), "MaximumBytes must be positive.");

        ConnectionId = options.ConnectionId;
        DefaultChangeMode = options.DefaultChangeMode;
        DefaultChangeGranularity = options.DefaultChangeGranularity;
        CalculateFormulas = options.CalculateFormulas;
        RejectFormulaErrors = options.RejectFormulaErrors;
        ReadOnly = options.ReadOnly;
        _root = Path.GetFullPath(options.RootPath);
        var separator = Path.DirectorySeparatorChar.ToString();
        _rootWithSeparator = _root.EndsWith(separator, StringComparison.Ordinal) ? _root : _root + separator;
        if (!ReadOnly)
        {
            Directory.CreateDirectory(_root);
            Directory.CreateDirectory(Path.Combine(_root, IndexDirectoryName));
        }
        _indexPath = Path.Combine(_root, IndexDirectoryName, IndexFileName);

        (TemplatesFolder, _templatesDirectory) = ResolveTemplatesFolder(options.TemplatesFolder);
        TemplateLocation = options.TemplateLocation == TemplateLocationKind.None && TemplatesFolder is not null
            ? TemplateLocationKind.Connection
            : options.TemplateLocation;
        // A template location with no folder named is a template location as a whole.
        if (_templatesDirectory is null && TemplateLocation != TemplateLocationKind.None)
            _templatesDirectory = _root;
        _maximumBytes = options.MaximumBytes;
        if (options.AllowedExtensions is null || options.AllowedExtensions.Count == 0)
            throw new ArgumentException("At least one allowed document extension is required.", nameof(options));
        _allowedExtensions = new HashSet<string>(
            options.AllowedExtensions.Select(NormalizeExtension),
            StringComparer.OrdinalIgnoreCase);

        _index = LoadIndex();
    }

    /// <inheritdoc />
    public string Provider => ProviderName;

    /// <inheritdoc />
    public string ConnectionId { get; }

    /// <inheritdoc />
    public ChangeMode DefaultChangeMode { get; }

    /// <inheritdoc />
    public ChangeGranularity DefaultChangeGranularity { get; }

    /// <inheritdoc />
    public bool CalculateFormulas { get; }

    /// <inheritdoc />
    public bool RejectFormulaErrors { get; }

    /// <summary>
    /// Gets whether the connection refuses every write; see
    /// <see cref="FileSystemDocumentProviderOptions.ReadOnly"/>.
    /// </summary>
    public bool ReadOnly { get; }

    /// <inheritdoc />
    /// <remarks>The folder relative to the root, with <c>/</c> separators, or <see langword="null"/>.</remarks>
    public string? TemplatesFolder { get; }

    /// <inheritdoc />
    public TemplateLocationKind TemplateLocation { get; }

    /// <inheritdoc />
    public async Task<DocumentReference> RegisterAsync(
        string source,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(source))
            throw Error(ProviderErrorCode.InvalidArgument, "A filesystem registration requires a non-empty path.");

        var fullPath = ResolveAndValidateSource(source);

        var info = new FileInfo(fullPath);
        if (info.Length > _maximumBytes)
            throw Error(ProviderErrorCode.ContentTooLarge,
                $"Document at '{fullPath}' exceeds the configured maximum of {_maximumBytes} bytes.");

        var version = await ComputeVersionStreamingAsync(fullPath, cancellationToken).ConfigureAwait(false);
        var id = await MintAndPersistAsync(fullPath, cancellationToken).ConfigureAwait(false);
        return CreateReference(id, version, Path.GetFileName(fullPath));
    }

    /// <inheritdoc />
    public async Task<DocumentReference> CreateAsync(
        string name,
        Stream content,
        CancellationToken cancellationToken = default)
    {
        if (content is null) throw new ArgumentNullException(nameof(content));
        RefuseWhenReadOnly(itemId: null);

        var fileName = ValidateCreateName(name);
        var destinationPath = Path.Combine(_root, fileName);
        if (File.Exists(destinationPath) || Directory.Exists(destinationPath))
            throw Error(ProviderErrorCode.AlreadyExists,
                $"A document named '{fileName}' already exists in this connection; choose another name.");

        var bytes = await ReadStreamAsync(content, itemId: null, cancellationToken).ConfigureAwait(false);

        try
        {
            // WriteAtomicallyAsync stages beside the destination and publishes with a
            // no-overwrite File.Move, so concurrent creators cannot displace each other.
            await WriteAtomicallyAsync(
                destinationPath, bytes, replace: false, cancellationToken).ConfigureAwait(false);
        }
        // WriteAtomicallyAsync reports every failed publish as WriteRejected, wrapping the
        // framework exception. These handlers once caught IOException directly, which never
        // arrives here, so a lost creation race was not reported as AlreadyExists and the
        // framework message, which carries the absolute path, was echoed unchanged.
        catch (DocumentProviderException ex) when (
            ex.Code == ProviderErrorCode.WriteRejected &&
            (File.Exists(destinationPath) || Directory.Exists(destinationPath)))
        {
            throw Error(ProviderErrorCode.AlreadyExists,
                $"A document named '{fileName}' already exists in this connection; choose another name.",
                itemId: null, ex);
        }
        catch (DocumentProviderException ex) when (ex.Code == ProviderErrorCode.WriteRejected)
        {
            throw Error(ProviderErrorCode.WriteRejected,
                $"Could not create '{fileName}': the connection's storage rejected the write; nothing was created.",
                itemId: null, ex);
        }

        string id;
        try
        {
            // Once the file exists, finish registration even if the caller cancels. Stopping
            // between publication and index persistence would manufacture an avoidable
            // unregistered document.
            id = await MintAndPersistAsync(destinationPath, CancellationToken.None).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            // Do not compensate by deleting the published path: another actor could have
            // opened, changed, or replaced it after publication. The caller knows the name
            // and can register that file explicitly instead of risking data loss.
            throw RegistrationFailed(
                $"Document '{fileName}' was created in this connection, but its registration " +
                "could not be persisted. The file exists but is unregistered. Do not retry " +
                "creation; recover the known filename through registration when available, " +
                "or report it to the host/operator.",
                itemId: null, fileName, bytes, ex);
        }

        return CreateReference(id, ComputeVersion(bytes), fileName);
    }

    /// <inheritdoc />
    public async Task<DocumentContent> OpenReadAsync(
        DocumentReference reference,
        CancellationToken cancellationToken = default)
    {
        ValidateReference(reference);
        // Revalidate on every access. A directory inside the root could have been
        // replaced with a link after registration; stored mappings do not bypass the
        // connection boundary.
        var path = ResolveAndValidateSource(LookupPath(reference.ItemId));

        var info = new FileInfo(path);
        if (!info.Exists)
            throw Error(ProviderErrorCode.NotFound, $"The referenced file for '{reference.ItemId}' no longer exists.", reference.ItemId);
        if (info.Length > _maximumBytes)
            throw Error(ProviderErrorCode.ContentTooLarge,
                $"Document exceeds the configured maximum of {_maximumBytes} bytes.", reference.ItemId);

        // A file another process holds exclusively - a document open in Word - fails here,
        // before any edit is attempted. Reported as an IO error naming the file, because a
        // raw IOException is outside the error vocabulary this contract defines and reaches
        // callers as an unexplained internal failure.
        byte[] bytes;
        try
        {
            bytes = await ReadAllBytesAsync(path, cancellationToken).ConfigureAwait(false);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            throw Error(
                ProviderErrorCode.IO,
                $"The document could not be read: {ex.Message} " +
                $"'{Path.GetFileName(path)}' is most likely open in another process.",
                reference.ItemId, ex);
        }

        var version = ComputeVersion(bytes);
        if (!string.IsNullOrEmpty(reference.Version) &&
            !string.Equals(reference.Version, version, StringComparison.Ordinal))
            throw new DocumentVersionConflictException(reference.Version!, version, ProviderName, ConnectionId, reference.ItemId);

        var canonical = CreateReference(reference.ItemId, version, Path.GetFileName(path));
        return new DocumentContent(canonical, new MemoryStream(bytes, writable: false));
    }

    /// <inheritdoc />
    public async Task<DocumentReference> SaveAsync(
        DocumentReference source,
        Stream content,
        SaveDocumentOptions options,
        CancellationToken cancellationToken = default)
    {
        ValidateReference(source);
        RefuseWhenReadOnly(source.ItemId);

        // The optimistic check reads the file and the publish writes it. Without holding
        // the item's gate across both, two callers that read the same version both pass
        // the check and both write, and the first caller's edit disappears while it is
        // told the save succeeded. The gate makes check-and-publish one step.
        var gate = ItemLock(source.ItemId);
        await gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            return await SaveCoreAsync(source, content, options, cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            gate.Release();
        }
    }

    private SemaphoreSlim ItemLock(string itemId)
    {
        lock (_itemLocks)
        {
            if (!_itemLocks.TryGetValue(itemId, out var gate))
            {
                gate = new SemaphoreSlim(1, 1);
                _itemLocks[itemId] = gate;
            }

            return gate;
        }
    }

    private async Task<DocumentReference> SaveCoreAsync(
        DocumentReference source,
        Stream content,
        SaveDocumentOptions options,
        CancellationToken cancellationToken)
    {
        if (content is null) throw new ArgumentNullException(nameof(content));
        if (options is null) throw new ArgumentNullException(nameof(options));
        if (!string.IsNullOrEmpty(options.DestinationItemId))
            throw Error(ProviderErrorCode.InvalidArgument,
                "Filesystem item ids are opaque and provider-assigned; a destination id cannot be chosen by the caller. Use Mode=NewVersion/NewDocument (optionally with NewName) to mint a new document, or Mode=Replace to overwrite the source.",
                source.ItemId);
        if (options.Mode == SaveMode.Replace && !string.IsNullOrWhiteSpace(options.NewName))
            throw Error(ProviderErrorCode.InvalidArgument,
                "Replace mode overwrites the source in place and cannot rename. Use Mode=NewVersion/NewDocument with NewName to mint a renamed document.",
                source.ItemId);

        // Revalidate the stored path immediately before reading or publishing output.
        var sourcePath = ResolveAndValidateSource(LookupPath(source.ItemId));
        var sourceName = Path.GetFileName(sourcePath);
        var sourceDir = Path.GetDirectoryName(sourcePath)
            ?? throw Error(ProviderErrorCode.WriteRejected, "The referenced file has no parent directory; nothing was changed.", source.ItemId);

        var expectedVersion = options.ExpectedVersion ?? source.Version;
        if (!string.IsNullOrEmpty(expectedVersion))
        {
            // The optimistic check reads the file, so it meets the same locks the write
            // does - a document open in Word fails here rather than at the replace. It is
            // reported the same way for the same reason: a raw IOException is outside the
            // error vocabulary this contract defines.
            string actualVersion;
            try
            {
                actualVersion = await ComputeVersionStreamingAsync(sourcePath, cancellationToken).ConfigureAwait(false);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                // Refused before anything was written, so the outcome is certain.
                throw Error(
                    ProviderErrorCode.WriteRejected,
                    $"The document could not be read to check its version: {ex.Message} " +
                    $"'{sourceName}' is most likely open in another process; nothing was changed.",
                    source.ItemId, ex);
            }

            if (!string.Equals(expectedVersion, actualVersion, StringComparison.Ordinal))
                throw new DocumentVersionConflictException(expectedVersion!, actualVersion, ProviderName, ConnectionId, source.ItemId);
        }

        var bytes = await ReadStreamAsync(content, source.ItemId, cancellationToken).ConfigureAwait(false);
        if (bytes.LongLength > _maximumBytes)
            throw Error(ProviderErrorCode.ContentTooLarge,
                $"Document exceeds the configured maximum of {_maximumBytes} bytes.", source.ItemId);

        if (options.Mode == SaveMode.Replace)
        {
            await WriteAtomicallyAsync(sourcePath, bytes, replace: true, cancellationToken).ConfigureAwait(false);
            return CreateReference(source.ItemId, ComputeVersion(bytes), sourceName);
        }

        var destinationName = string.IsNullOrWhiteSpace(options.NewName)
            ? NextVersionedName(sourceDir, sourceName)
            : ValidateName(options.NewName!);
        var destinationPath = Path.Combine(sourceDir, destinationName);
        // A taken name is refused before anything is written, so it is AlreadyExists, which
        // callers may treat as certain, rather than IO, which they must treat as uncertain.
        if (File.Exists(destinationPath))
            throw Error(ProviderErrorCode.AlreadyExists,
                "A file already exists at the destination path; refusing to overwrite without Replace mode.",
                source.ItemId);
        await WriteAtomicallyAsync(destinationPath, bytes, replace: false, cancellationToken).ConfigureAwait(false);

        string newId;
        try
        {
            // The file is published. Finish registering it even if the caller cancels now:
            // stopping here would leave a written, unregistered document behind a plain
            // cancellation, which a caller would read as nothing written.
            newId = await MintAndPersistAsync(destinationPath, CancellationToken.None).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            throw RegistrationFailed(
                $"Document '{destinationName}' was written, but its registration could not be persisted. " +
                "Do not save it again; register the file by name to recover it.",
                source.ItemId, destinationName, bytes, ex);
        }
        return CreateReference(newId, ComputeVersion(bytes), destinationName);
    }

    /// <inheritdoc />
    public async Task RemoveAsync(
        DocumentReference reference,
        CancellationToken cancellationToken = default)
    {
        ValidateReference(reference);
        RefuseWhenReadOnly(reference.ItemId);
        await _indexLock.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            if (!_index.Remove(reference.ItemId))
                throw Error(ProviderErrorCode.NotFound, $"The provider document '{reference.ItemId}' is not registered.", reference.ItemId);
            await PersistIndexAsync(cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            _indexLock.Release();
        }
    }

    // ── Listing ───────────────────────────────────────────────────────────────

    /// <inheritdoc />
    /// <remarks>
    /// Entries are returned in name order, folder by folder, and each <see cref="DocumentListEntry.Source"/>
    /// is the path relative to the root with <c>/</c> separators. A file the connection would not
    /// register is not listed: one whose extension it does not allow, a macro-enabled one, or one
    /// over its size limit. The continuation token names the last entry examined, relative to the
    /// listed folder; it is only ever compared with names found under that folder, never opened,
    /// so a forged one cannot reach outside the root.
    /// </remarks>
    public Task<DocumentListPage> ListAsync(DocumentListQuery query, CancellationToken cancellationToken = default)
    {
        if (query is null) throw new ArgumentNullException(nameof(query));
        try
        {
            return Task.FromResult(List(query, cancellationToken));
        }
        catch (Exception ex)
        {
            return Task.FromException<DocumentListPage>(ex);
        }
    }

    private DocumentListPage List(DocumentListQuery query, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        string listingRoot;
        if (query.Templates)
        {
            if (_templatesDirectory is null) return new DocumentListPage();
            listingRoot = _templatesDirectory;
        }
        else
        {
            listingRoot = _root;
        }

        var fingerprint = DocumentListingRules.Fingerprint(query, ProviderName);
        var resume = ParseListingToken(query.ContinuationToken, fingerprint);
        var folder = ResolveListingFolder(listingRoot, query.Folder);
        // Every component below the root, the templates folder included, is checked before the
        // folder is opened: enumerating through a link would list whatever it points at.
        if (folder.StartsWith(_rootWithSeparator, StringComparison.Ordinal))
            EnsureNoReparsePoints(folder, _rootWithSeparator);
        // A folder that is not there lists nothing rather than failing: a templates folder the
        // host named may simply not have been created yet.
        if (!Directory.Exists(folder)) return new DocumentListPage();

        var walk = new ListingWalk(this, query, listingRoot, cancellationToken);
        walk.Visit(new DirectoryInfo(folder), new List<string>(), depth: 0, resume, isListedFolder: true);

        string? token = null;
        if (walk.More)
        {
            var position = walk.LastVisited ?? (resume is null ? null : string.Join("/", resume));
            if (position is not null)
                token = DocumentListingRules.EncodeToken($"{TokenVersion}\n{fingerprint}\n{position}");
        }

        return new DocumentListPage { Items = walk.Items, ContinuationToken = token };
    }

    /// <summary>The path segments a token resumes after, or <see langword="null"/> for the first page.</summary>
    private string[]? ParseListingToken(string? token, string fingerprint)
    {
        if (string.IsNullOrWhiteSpace(token)) return null;

        var parts = DocumentListingRules.DecodeToken(token!)?.Split(new[] { '\n' }, 3);
        if (parts is not { Length: 3 } || parts[0] != TokenVersion || parts[2].Length == 0)
            throw Error(ProviderErrorCode.InvalidArgument,
                "The continuation token is not one this connection issued. Start the listing again without it.");
        if (!string.Equals(parts[1], fingerprint, StringComparison.Ordinal))
            throw Error(ProviderErrorCode.InvalidArgument,
                "The continuation token belongs to a listing with a different folder or filter. Repeat that query, or start again without the token.");

        return parts[2].Split('/');
    }

    /// <summary>
    /// Resolves the folder to list under <paramref name="listingRoot"/>, with the same boundary
    /// checks registration applies to a source: it must stay under the listing root, outside the
    /// index directory, and reach it through no link.
    /// </summary>
    private string ResolveListingFolder(string listingRoot, string? folder)
    {
        if (string.IsNullOrWhiteSpace(folder)) return listingRoot;

        var trimmed = folder!.Trim().Replace('/', Path.DirectorySeparatorChar).Replace('\\', Path.DirectorySeparatorChar);
        var fullPath = Path.GetFullPath(Path.IsPathRooted(trimmed) ? trimmed : Path.Combine(listingRoot, trimmed))
            .TrimEnd(Path.DirectorySeparatorChar);
        var listingRootWithSeparator = listingRoot.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal)
            ? listingRoot
            : listingRoot + Path.DirectorySeparatorChar;

        if (!string.Equals(fullPath, listingRoot.TrimEnd(Path.DirectorySeparatorChar), StringComparison.Ordinal) &&
            !fullPath.StartsWith(listingRootWithSeparator, StringComparison.Ordinal))
            throw Error(ProviderErrorCode.AccessDenied,
                "The folder must lie under the connection's root" + (listingRoot == _root ? "." : " templates folder."));

        var indexDirectory = Path.Combine(_root, IndexDirectoryName);
        if (string.Equals(fullPath, indexDirectory, StringComparison.Ordinal) ||
            fullPath.StartsWith(indexDirectory + Path.DirectorySeparatorChar, StringComparison.Ordinal))
            throw Error(ProviderErrorCode.AccessDenied, "The provider's own index directory cannot be listed.");

        if (File.Exists(fullPath))
            throw Error(ProviderErrorCode.InvalidArgument, "The folder names a file, not a folder.");
        return fullPath;
    }

    /// <summary>The entry's path relative to the root, with <c>/</c> separators: what registration accepts.</summary>
    private string RelativeToRoot(string fullPath) =>
        fullPath.Substring(_rootWithSeparator.Length).Replace(Path.DirectorySeparatorChar, '/');

    /// <summary>
    /// One listing call: a depth-first walk in name order that stops when the page is full or
    /// the work bound is reached, remembering the last entry it examined.
    /// </summary>
    private sealed class ListingWalk
    {
        private readonly FileSystemDocumentProvider _provider;
        private readonly DocumentListQuery _query;
        private readonly string _listingRoot;
        private readonly HashSet<string> _extensions;
        private readonly int _pageSize;
        private readonly CancellationToken _cancellationToken;
        private int _budget = MaximumEntriesPerListing;

        public ListingWalk(FileSystemDocumentProvider provider, DocumentListQuery query, string listingRoot, CancellationToken cancellationToken)
        {
            _provider = provider;
            _query = query;
            _listingRoot = listingRoot;
            _extensions = DocumentListingRules.Extensions(query, provider._allowedExtensions);
            _pageSize = DocumentListingRules.PageSize(query);
            _cancellationToken = cancellationToken;
        }

        public List<DocumentListEntry> Items { get; } = new();

        /// <summary>The last entry examined, relative to the listed folder.</summary>
        public string? LastVisited { get; private set; }

        /// <summary>Whether the walk stopped before it ran out of entries.</summary>
        public bool More { get; private set; }

        /// <summary>
        /// Visits one folder's entries in name order. Returns <see langword="false"/> once the
        /// walk has stopped. <paramref name="resume"/>, when set, skips everything up to and
        /// including the entry it names, so the walk continues exactly where a token left it.
        /// </summary>
        public bool Visit(DirectoryInfo directory, List<string> prefix, int depth, string[]? resume, bool isListedFolder)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            var entries = Entries(directory, isListedFolder);
            var level = prefix.Count;

            foreach (var entry in entries)
            {
                var isFolder = entry is DirectoryInfo;
                var descend = isFolder && _query.Recursive && depth < MaximumListingDepth;
                if (isFolder && !_query.Recursive) continue;

                if (resume is not null && level < resume.Length)
                {
                    var order = CompareNames(entry.Name, resume[level]);
                    if (order < 0) continue;
                    if (order == 0)
                    {
                        // The token names this entry or something beneath it. A file it names was
                        // already examined; a folder it names was entered, so continue inside it.
                        if (descend &&
                            !Visit((DirectoryInfo)entry, Append(prefix, entry.Name), depth + 1,
                                level < resume.Length - 1 ? resume : null, isListedFolder: false))
                            return false;
                        continue;
                    }
                }

                if (_budget == 0)
                {
                    More = true;
                    return false;
                }

                var path = string.Join("/", Append(prefix, entry.Name));
                if (entry is FileInfo file && Lists(file))
                {
                    if (Items.Count == _pageSize)
                    {
                        // Another match exists, so the page is not the last one. The token points
                        // just before it, so the next page starts here.
                        More = true;
                        return false;
                    }
                    Items.Add(_provider.EntryFor(file, _listingRoot));
                }

                _budget--;
                LastVisited = path;
                if (descend && !Visit((DirectoryInfo)entry, Append(prefix, entry.Name), depth + 1, null, isListedFolder: false))
                    return false;
            }

            return true;
        }

        private bool Lists(FileInfo file) =>
            DocumentListingRules.Matches(file.Name, _extensions, _query) &&
            file.Length <= _provider._maximumBytes;

        /// <summary>
        /// A folder's listable entries in name order. A subfolder that cannot be read is
        /// skipped rather than failing the whole listing; the listed folder itself reports why.
        /// </summary>
        private List<FileSystemInfo> Entries(DirectoryInfo directory, bool isListedFolder)
        {
            var entries = new List<FileSystemInfo>();
            try
            {
                foreach (var entry in directory.EnumerateFileSystemInfos())
                {
                    if (!Listable(entry)) continue;
                    entries.Add(entry);
                    if (entries.Count == MaximumEntriesPerFolder) break;
                }
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Security.SecurityException)
            {
                if (!isListedFolder) return new List<FileSystemInfo>();
                // The framework message names the absolute path, so it stays on the inner exception.
                throw _provider.Error(ProviderErrorCode.IO, "The folder could not be listed.", itemId: null, ex);
            }

            entries.Sort((left, right) => CompareNames(left.Name, right.Name));
            return entries;
        }

        /// <summary>
        /// Whether an entry may appear at all. Links and junctions are never followed, so a
        /// listing cannot leave the root; hidden, system and dot-prefixed entries are the host's
        /// and the provider's own; <c>~$</c> files are Office's lock files for open documents.
        /// </summary>
        private static bool Listable(FileSystemInfo entry)
        {
            var attributes = entry.Attributes;
            if ((attributes & (FileAttributes.ReparsePoint | FileAttributes.Hidden | FileAttributes.System)) != 0)
                return false;
            return !entry.Name.StartsWith(".", StringComparison.Ordinal) &&
                   !entry.Name.StartsWith("~$", StringComparison.Ordinal);
        }

        private static List<string> Append(List<string> prefix, string name) => new(prefix) { name };

        /// <summary>Name order: case-insensitive, with an ordinal tie-break so the order is total.</summary>
        private static int CompareNames(string left, string right)
        {
            var order = StringComparer.OrdinalIgnoreCase.Compare(left, right);
            return order != 0 ? order : StringComparer.Ordinal.Compare(left, right);
        }
    }

    private DocumentListEntry EntryFor(FileInfo file, string listingRoot)
    {
        var directory = file.DirectoryName ?? listingRoot;
        var trimmedRoot = listingRoot.TrimEnd(Path.DirectorySeparatorChar);
        string? folder = string.Equals(directory, trimmedRoot, StringComparison.Ordinal)
            ? null
            : directory.Substring(trimmedRoot.Length + 1).Replace(Path.DirectorySeparatorChar, '/');

        return new DocumentListEntry
        {
            Source = RelativeToRoot(file.FullName),
            Name = file.Name,
            Folder = folder,
            ContentType = ContentTypeFor(file.Name),
            Length = file.Length,
            LastModified = new DateTimeOffset(file.LastWriteTimeUtc, TimeSpan.Zero),
            IsTemplate = DocumentListingRules.IsTemplateName(file.Name) || IsInTemplatesDirectory(file.FullName)
        };
    }

    private bool IsInTemplatesDirectory(string fullPath) =>
        _templatesDirectory is not null &&
        (string.Equals(_templatesDirectory, _root, StringComparison.Ordinal) ||
         fullPath.StartsWith(_templatesDirectory + Path.DirectorySeparatorChar, StringComparison.Ordinal));

    /// <summary>
    /// Validates the configured templates folder: relative, inside the root and outside the
    /// index directory. It need not exist yet; listing a missing one lists nothing.
    /// </summary>
    private (string? Display, string? FullPath) ResolveTemplatesFolder(string? configured)
    {
        if (string.IsNullOrWhiteSpace(configured)) return (null, null);

        var trimmed = configured!.Trim().Replace('\\', '/').Trim('/');
        if (Path.IsPathRooted(configured.Trim()))
            throw new ArgumentException(
                "TemplatesFolder must be a folder relative to the connection root, not an absolute path.", "options");

        var fullPath = Path.GetFullPath(Path.Combine(_root, trimmed.Replace('/', Path.DirectorySeparatorChar)))
            .TrimEnd(Path.DirectorySeparatorChar);
        if (string.Equals(fullPath, _root.TrimEnd(Path.DirectorySeparatorChar), StringComparison.Ordinal))
            throw new ArgumentException(
                "TemplatesFolder must name a folder below the connection root. To make the whole connection " +
                "a template location, leave TemplatesFolder empty and set TemplateLocation instead.", "options");
        if (!fullPath.StartsWith(_rootWithSeparator, StringComparison.Ordinal))
            throw new ArgumentException("TemplatesFolder must stay inside the connection root.", "options");
        var indexDirectory = Path.Combine(_root, IndexDirectoryName);
        if (string.Equals(fullPath, indexDirectory, StringComparison.Ordinal) ||
            fullPath.StartsWith(indexDirectory + Path.DirectorySeparatorChar, StringComparison.Ordinal))
            throw new ArgumentException("TemplatesFolder cannot be the provider's own index directory.", "options");

        return (RelativeToRoot(fullPath), fullPath);
    }

    /// <summary>The refusal every write on a read-only connection meets before anything is touched.</summary>
    private void RefuseWhenReadOnly(string? itemId)
    {
        if (ReadOnly)
            throw Error(ProviderErrorCode.AccessDenied,
                "This connection is read-only: documents in it can be listed, opened and used as templates, " +
                "but not saved, created or removed. Nothing was changed.",
                itemId);
    }

    // ── Registry plumbing ─────────────────────────────────────────────────────

    private async Task<string> MintAndPersistAsync(string fullPath, CancellationToken cancellationToken)
    {
        await _indexLock.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            var id = NewItemId();
            _index[id] = fullPath;
            // A read-only connection keeps its registrations in memory: writing the index would
            // be a write under a root the host said must not be written.
            if (ReadOnly) return id;
            try
            {
                await PersistIndexAsync(cancellationToken).ConfigureAwait(false);
            }
            catch
            {
                // Keep memory and disk agreeing: an id that did not survive the write must
                // not answer lookups for the rest of the process's life.
                _index.Remove(id);
                throw;
            }
            return id;
        }
        finally
        {
            _indexLock.Release();
        }
    }

    private string LookupPath(string itemId)
    {
        if (string.IsNullOrWhiteSpace(itemId))
            throw Error(ProviderErrorCode.InvalidArgument, "An item id is required.");
        if (Path.IsPathRooted(itemId) ||
            itemId.IndexOfAny(new[] { '/', '\\' }) >= 0 ||
            itemId.Contains(".."))
            throw Error(ProviderErrorCode.AccessDenied,
                "Filesystem item ids are opaque tokens, not paths.", itemId);

        _indexLock.Wait();
        try
        {
            if (!_index.TryGetValue(itemId, out var path))
                throw Error(ProviderErrorCode.NotFound, $"The provider document '{itemId}' is not registered.", itemId);
            return path;
        }
        finally
        {
            _indexLock.Release();
        }
    }

    private Dictionary<string, string> LoadIndex()
    {
        if (!File.Exists(_indexPath))
            return new Dictionary<string, string>(StringComparer.Ordinal);

        try
        {
            var json = File.ReadAllText(_indexPath);
            var model = JsonSerializer.Deserialize<IndexFile>(json, IndexJsonOptions);
            return model?.Items is { Count: > 0 }
                ? new Dictionary<string, string>(model.Items, StringComparer.Ordinal)
                : new Dictionary<string, string>(StringComparer.Ordinal);
        }
        catch (Exception ex) when (ex is IOException or JsonException or UnauthorizedAccessException)
        {
            throw new DocumentProviderException(
                ProviderErrorCode.ConfigurationError,
                $"Could not load the filesystem provider index at '{_indexPath}': {ex.Message}",
                ProviderName, ConnectionId, itemId: null, ex);
        }
    }

    private async Task PersistIndexAsync(CancellationToken cancellationToken)
    {
        var model = new IndexFile { Version = 1, Items = new Dictionary<string, string>(_index, StringComparer.Ordinal) };
        var payload = JsonSerializer.SerializeToUtf8Bytes(model, IndexJsonOptions);
        await WriteAtomicallyAsync(_indexPath, payload, replace: true, cancellationToken).ConfigureAwait(false);
    }

    private static readonly JsonSerializerOptions IndexJsonOptions = new()
    {
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    private sealed class IndexFile
    {
        public int Version { get; set; }
        public Dictionary<string, string> Items { get; set; } = new(StringComparer.Ordinal);
    }

    // ── Path validation & naming ──────────────────────────────────────────────

    /// <summary>
    /// Resolves a host-supplied source string to a full filesystem path under the
    /// configured root. Accepts an absolute path that is already under root, or any
    /// relative path (resolved against root). Rejects traversal, symlinks/reparse
    /// points, non-existent files, disallowed extensions, and the provider's own
    /// index directory.
    /// </summary>
    private string ResolveAndValidateSource(string source)
    {
        var fullPath = Path.IsPathRooted(source)
            ? Path.GetFullPath(source)
            : Path.GetFullPath(Path.Combine(_root, source));

        var separator = Path.DirectorySeparatorChar.ToString();
        var rootWithSeparator = _root.EndsWith(separator, StringComparison.Ordinal)
            ? _root
            : _root + separator;
        if (!fullPath.StartsWith(rootWithSeparator, StringComparison.Ordinal) &&
            !string.Equals(fullPath, _root, StringComparison.Ordinal))
            throw Error(ProviderErrorCode.AccessDenied,
                $"The source path must lie under the connection's configured root '{_root}'.");

        if (fullPath.StartsWith(Path.Combine(_root, IndexDirectoryName), StringComparison.Ordinal))
            throw Error(ProviderErrorCode.AccessDenied,
                "The provider's own index directory cannot be registered as a document.");

        // The boundary is enforced before existence is probed. Asking whether the file
        // exists first means asking *through* whatever links are in the path: the answer
        // then depends on the target rather than on the link, so the same escape attempt
        // reports NotFound or AccessDenied according to what happens to be on the other
        // side. It also answers an existence question about a location outside the root.
        // The boundary is enforced before existence is probed. Asking whether the file
        // exists first means asking *through* whatever links are in the path: the answer
        // then depends on the target rather than on the link, so the same escape attempt
        // reports NotFound or AccessDenied according to what happens to be on the other
        // side. It also answers an existence question about a location outside the root.
        EnsureNoReparsePoints(fullPath, rootWithSeparator);

        if (!File.Exists(fullPath))
            throw Error(ProviderErrorCode.NotFound, $"The source file '{fullPath}' does not exist.");

        var extension = NormalizeExtension(Path.GetExtension(fullPath));
        if (!_allowedExtensions.Contains(extension))
            throw Error(ProviderErrorCode.ExtensionNotAllowed,
                $"Documents with extension '{extension}' are not allowed by this connection. Allowed: {string.Join(", ", _allowedExtensions)}.");

        return fullPath;
    }

    private void EnsureNoReparsePoints(string fullPath, string rootWithSeparator)
    {
        // Check every component beneath the configured root, not only the final file.
        // Otherwise root/link/file.docx can lexically remain under root while link
        // resolves to an arbitrary directory outside it.
        var relative = fullPath.Substring(rootWithSeparator.Length);
        var current = _root;
        foreach (var component in relative.Split(
                     new[] { Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar },
                     StringSplitOptions.RemoveEmptyEntries))
        {
            current = Path.Combine(current, component);

            // Stop at the first component that is not there. Whether the path exists is
            // decided by the caller afterwards; asking for the attributes of a missing
            // entry would throw an exception that says nothing useful. A link whose target
            // cannot be followed still fails here, on the link itself, which is the point.
            if (!File.Exists(current) && !Directory.Exists(current)) return;

            var attributes = File.GetAttributes(current);
            if ((attributes & FileAttributes.ReparsePoint) == FileAttributes.ReparsePoint)
                throw Error(
                    ProviderErrorCode.AccessDenied,
                    "Symlinks and reparse points are not accepted as document sources.");
        }
    }

    /// <summary>Validates a caller-supplied display name (used only by NewName saves).</summary>
    private string ValidateName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw Error(ProviderErrorCode.InvalidArgument, "A document name is required.");
        var leaf = Path.GetFileName(name);
        if (string.IsNullOrWhiteSpace(leaf) || !string.Equals(leaf, name, StringComparison.Ordinal))
            throw Error(ProviderErrorCode.InvalidArgument,
                "A document name must be a bare file name without path segments.");
        var extension = NormalizeExtension(Path.GetExtension(leaf));
        if (!_allowedExtensions.Contains(extension))
            throw Error(ProviderErrorCode.ExtensionNotAllowed,
                $"Documents with extension '{extension}' are not allowed by this connection. Allowed: {string.Join(", ", _allowedExtensions)}.");
        return leaf;
    }

    private string ValidateCreateName(string name)
    {
        if (!string.IsNullOrWhiteSpace(name) &&
            (name.IndexOf('/') >= 0 || name.IndexOf('\\') >= 0 || name == "." || name == ".."))
            throw Error(ProviderErrorCode.InvalidArgument,
                "A document name must be a bare file name without path segments.");
        if (!string.IsNullOrWhiteSpace(name) &&
            PortableDocumentName.ContainsInvalidCharacter(name))
            throw Error(ProviderErrorCode.InvalidArgument,
                "A document name must not contain characters that are invalid in a file name.");
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

    /// <summary>
    /// Picks the next unused versioned sibling file name in <paramref name="directory"/>.
    /// For <c>contract.docx</c> the sequence is <c>contract.v2.docx</c>, <c>contract.v3.docx</c>, …
    /// An already-versioned source name keeps the same stem and just bumps the suffix.
    /// </summary>
    private static string NextVersionedName(string directory, string sourceName)
    {
        var extension = Path.GetExtension(sourceName);
        var stem = Path.GetFileNameWithoutExtension(sourceName);
        var (baseStem, startVersion) = ParseVersionedStem(stem);

        for (var n = Math.Max(2, startVersion + 1); n < int.MaxValue; n++)
        {
            var candidate = $"{baseStem}.v{n}{extension}";
            if (!File.Exists(Path.Combine(directory, candidate)))
                return candidate;
        }
        throw new InvalidOperationException("Exhausted version suffixes for the source file.");
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

    private void ValidateReference(DocumentReference reference)
    {
        if (reference is null) throw new ArgumentNullException(nameof(reference));
        if (!string.Equals(reference.Provider, ProviderName, StringComparison.OrdinalIgnoreCase) ||
            !string.Equals(reference.ConnectionId, ConnectionId, StringComparison.Ordinal))
            throw new DocumentProviderException(
                ProviderErrorCode.InvalidArgument,
                $"The document reference (provider='{reference.Provider}', connectionId='{reference.ConnectionId}') does not belong to this filesystem connection (provider='{ProviderName}', connectionId='{ConnectionId}').",
                ProviderName, ConnectionId, reference.ItemId);
    }

    private DocumentReference CreateReference(string itemId, string version, string name) => new()
    {
        Provider = ProviderName,
        ConnectionId = ConnectionId,
        ItemId = itemId,
        Version = version,
        Name = name,
        ContentType = ContentTypeFor(name)
    };

    private static string ContentTypeFor(string name) => OfficeContentTypes.ForName(name);

    private DocumentProviderException Error(
        ProviderErrorCode code,
        string message,
        string? itemId = null,
        Exception? innerException = null) =>
        new(code, message, ProviderName, ConnectionId, itemId, innerException);

    private static string NewItemId() => Guid.NewGuid().ToString("N");

    // ── Hash / IO helpers ─────────────────────────────────────────────────────

    /// <summary>
    /// The published file exists but has no id: carry its connection-relative name and content
    /// hash so the caller can find and register it, never its path.
    /// </summary>
    private DocumentRegistrationFailedException RegistrationFailed(
        string message, string? itemId, string outputName, byte[] bytes, Exception innerException) =>
        new(message, ProviderName, ConnectionId, itemId, outputName,
            ComputeVersion(bytes).Substring("sha256:".Length), innerException);

    private static string ComputeVersion(byte[] bytes)
    {
        using var sha = SHA256.Create();
        var hash = sha.ComputeHash(bytes);
        return "sha256:" + string.Concat(hash.Select(b => b.ToString("x2")));
    }

    private static async Task<string> ComputeVersionStreamingAsync(string path, CancellationToken cancellationToken)
    {
        using var sha = SHA256.Create();
        using var stream = new FileStream(
            path,
            FileMode.Open,
            FileAccess.Read,
            FileShare.Read,
            bufferSize: 81920,
            useAsync: true);

        var buffer = new byte[81920];
        while (true)
        {
            var read = await stream.ReadAsync(buffer, 0, buffer.Length, cancellationToken).ConfigureAwait(false);
            if (read == 0) break;
            sha.TransformBlock(buffer, 0, read, null, 0);
        }
        sha.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
        return "sha256:" + string.Concat(sha.Hash!.Select(b => b.ToString("x2")));
    }

    private async Task<byte[]> ReadStreamAsync(Stream content, string? itemId, CancellationToken cancellationToken)
    {
        if (content.CanSeek) content.Position = 0;
        using var copy = new MemoryStream();
        var buffer = new byte[81920];
        while (true)
        {
            var read = await content.ReadAsync(buffer, 0, buffer.Length, cancellationToken).ConfigureAwait(false);
            if (read == 0) break;
            await copy.WriteAsync(buffer, 0, read, cancellationToken).ConfigureAwait(false);
            if (copy.Length > _maximumBytes)
                throw Error(ProviderErrorCode.ContentTooLarge,
                    $"Document exceeds the configured maximum of {_maximumBytes} bytes.", itemId);
        }
        return copy.ToArray();
    }

    private static Task<byte[]> ReadAllBytesAsync(string path, CancellationToken cancellationToken)
    {
#if NET8_0_OR_GREATER
        return File.ReadAllBytesAsync(path, cancellationToken);
#else
        cancellationToken.ThrowIfCancellationRequested();
        return Task.Run(() => File.ReadAllBytes(path), cancellationToken);
#endif
    }

    /// <remarks>
    /// A failure that outlives the retry is reported as a provider IO error naming the
    /// destination. Letting the raw <see cref="IOException"/> escape put it outside every
    /// error code the contract defines, and it reached callers as an unexplained internal
    /// error - the one outcome a caller can do nothing with.
    /// </remarks>
    private async Task WriteAtomicallyAsync(
        string destinationPath,
        byte[] bytes,
        bool replace,
        CancellationToken cancellationToken)
    {
        try
        {
            await WriteAtomicallyCoreAsync(destinationPath, bytes, replace, cancellationToken).ConfigureAwait(false);
        }
        catch (OperationCanceledException ex)
        {
            // Cancellation is observed only while the temporary file is written or between
            // publish attempts, never after a successful rename, so the destination is as it was.
            throw Error(
                ProviderErrorCode.WriteRejected,
                $"The write to '{Path.GetFileName(destinationPath)}' was cancelled before it was published; nothing was changed.",
                itemId: null,
                innerException: ex);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            // Publication is one rename or replace of a fully written temporary file, so a
            // failure leaves the destination as it was. That certainty is what WriteRejected
            // means, and what lets a caller retry without reconciling first.
            throw Error(
                ProviderErrorCode.WriteRejected,
                $"The document could not be written to '{Path.GetFileName(destinationPath)}': {ex.Message} " +
                "The destination is most likely held open by another process; nothing was changed.",
                itemId: null,
                innerException: ex);
        }
    }

    private static async Task WriteAtomicallyCoreAsync(
        string destinationPath,
        byte[] bytes,
        bool replace,
        CancellationToken cancellationToken)
    {
        var directory = Path.GetDirectoryName(destinationPath)
            ?? throw new InvalidOperationException("The destination has no parent directory.");
        Directory.CreateDirectory(directory);
        var temporaryPath = Path.Combine(directory, "." + Path.GetFileName(destinationPath) + "." + Guid.NewGuid().ToString("N") + ".tmp");

        try
        {
#if NET8_0_OR_GREATER
            await File.WriteAllBytesAsync(temporaryPath, bytes, cancellationToken).ConfigureAwait(false);
#else
            cancellationToken.ThrowIfCancellationRequested();
            await Task.Run(() => File.WriteAllBytes(temporaryPath, bytes), cancellationToken).ConfigureAwait(false);
#endif
            await PublishAsync(temporaryPath, destinationPath, replace, cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            if (File.Exists(temporaryPath)) File.Delete(temporaryPath);
        }
    }

    /// <summary>
    /// Moves the finished temporary file into place, retrying briefly while the
    /// destination is held by something else.
    /// </summary>
    /// <remarks>
    /// <para>
    /// On Windows this is where a save fails for a reason that has nothing to do with the
    /// caller: an antivirus scanner, the search indexer, or a preview handler opens the
    /// destination for a moment after it was last written, and <see cref="File.Replace(string, string, string)"/>
    /// answers <c>"Unable to remove the file to be replaced"</c>. It is transient by
    /// definition - the interfering process lets go - so the fix is to wait and try again
    /// rather than to fail an edit the engine had already completed.
    /// </para>
    /// <para>
    /// The retry is bounded and short. A destination held open by something that is not
    /// going to let go - a document open in Word, say - has to surface as an error, and
    /// promptly, rather than stalling a request that cannot succeed.
    /// </para>
    /// </remarks>
    private static async Task PublishAsync(
        string temporaryPath, string destinationPath, bool replace, CancellationToken cancellationToken)
    {
        const int attempts = 5;

        for (var attempt = 1; ; attempt++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            try
            {
                // File.Replace preserves the destination's identity, which File.Move would
                // not; it is only usable when there is a destination to replace.
                if (replace && File.Exists(destinationPath))
                    File.Replace(temporaryPath, destinationPath, null);
                else
                    File.Move(temporaryPath, destinationPath);
                return;
            }
            catch (Exception ex) when (attempt < attempts && IsTransient(ex))
            {
                // 20ms, 40ms, 80ms, 160ms: long enough for a scanner to finish with the
                // file, short enough that a genuine lock is reported quickly.
                await Task.Delay(TimeSpan.FromMilliseconds(10 * (1 << attempt)), cancellationToken)
                    .ConfigureAwait(false);
            }
        }
    }

    /// <summary>
    /// Whether the failure is the kind another attempt can get past: the file was busy,
    /// not missing, not forbidden by its permissions, and not on a full disk.
    /// </summary>
    private static bool IsTransient(Exception exception) => exception switch
    {
        FileNotFoundException => false,
        DirectoryNotFoundException => false,
        PathTooLongException => false,
        IOException io => !IsDiskFull(io),
        UnauthorizedAccessException => true,
        _ => false
    };

    /// <summary>A full disk does not empty itself, so retrying only delays the report.</summary>
    private static bool IsDiskFull(IOException exception)
    {
        const int ErrorDiskFull = 0x70;
        const int ErrorHandleDiskFull = 0x27;

        var code = exception.HResult & 0xFFFF;
        return code == ErrorDiskFull || code == ErrorHandleDiskFull;
    }
}
