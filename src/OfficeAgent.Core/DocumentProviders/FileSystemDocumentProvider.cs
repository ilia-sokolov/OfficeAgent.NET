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
    /// </summary>
    public string RootPath { get; set; } = string.Empty;

    /// <summary>Gets or sets the maximum accepted document size in bytes.</summary>
    public long MaximumBytes { get; set; } = 100 * 1024 * 1024;

    /// <summary>Gets or sets the allowed document extensions.</summary>
    public IReadOnlyCollection<string> AllowedExtensions { get; set; } = new[] { ".docx" };
}

/// <summary>
/// Filesystem document provider: a registry of references to files the host already
/// owns, addressed by an opaque item id assigned at registration time. The provider
/// never copies, mirrors, or owns content - it only persists the id → path mapping
/// under <c>{root}/.officeagent/index.json</c> and routes open/save back to the
/// referenced path. The agent only sees opaque ids and never receives a filesystem
/// path.
/// </summary>
public sealed class FileSystemDocumentProvider : IDocumentProvider
{
    /// <summary>Gets the provider discriminator used in document references.</summary>
    public const string ProviderName = "filesystem";

    private const string IndexDirectoryName = ".officeagent";
    private const string IndexFileName = "index.json";

    private readonly string _root;
    private readonly string _indexPath;
    private readonly long _maximumBytes;
    private readonly HashSet<string> _allowedExtensions;
    private readonly SemaphoreSlim _indexLock = new(1, 1);
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
        _root = Path.GetFullPath(options.RootPath);
        Directory.CreateDirectory(_root);
        Directory.CreateDirectory(Path.Combine(_root, IndexDirectoryName));
        _indexPath = Path.Combine(_root, IndexDirectoryName, IndexFileName);
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
    public async Task<DocumentContent> OpenReadAsync(
        DocumentReference reference,
        CancellationToken cancellationToken = default)
    {
        ValidateReference(reference);
        var path = LookupPath(reference.ItemId);

        var info = new FileInfo(path);
        if (!info.Exists)
            throw Error(ProviderErrorCode.NotFound, $"The referenced file for '{reference.ItemId}' no longer exists.", reference.ItemId);
        if (info.Length > _maximumBytes)
            throw Error(ProviderErrorCode.ContentTooLarge,
                $"Document exceeds the configured maximum of {_maximumBytes} bytes.", reference.ItemId);

        var bytes = await ReadAllBytesAsync(path, cancellationToken).ConfigureAwait(false);
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

        var sourcePath = LookupPath(source.ItemId);
        var sourceName = Path.GetFileName(sourcePath);
        var sourceDir = Path.GetDirectoryName(sourcePath)
            ?? throw Error(ProviderErrorCode.IO, "The referenced file has no parent directory.", source.ItemId);

        var expectedVersion = options.ExpectedVersion ?? source.Version;
        if (!string.IsNullOrEmpty(expectedVersion))
        {
            var actualVersion = await ComputeVersionStreamingAsync(sourcePath, cancellationToken).ConfigureAwait(false);
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
        if (File.Exists(destinationPath))
            throw Error(ProviderErrorCode.IO,
                $"A file already exists at the destination path; refusing to overwrite without Replace mode.",
                source.ItemId);
        await WriteAtomicallyAsync(destinationPath, bytes, replace: false, cancellationToken).ConfigureAwait(false);

        var newId = await MintAndPersistAsync(destinationPath, cancellationToken).ConfigureAwait(false);
        return CreateReference(newId, ComputeVersion(bytes), destinationName);
    }

    /// <inheritdoc />
    public async Task RemoveAsync(
        DocumentReference reference,
        CancellationToken cancellationToken = default)
    {
        ValidateReference(reference);
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

    // ── Registry plumbing ─────────────────────────────────────────────────────

    private async Task<string> MintAndPersistAsync(string fullPath, CancellationToken cancellationToken)
    {
        await _indexLock.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            var id = NewItemId();
            _index[id] = fullPath;
            await PersistIndexAsync(cancellationToken).ConfigureAwait(false);
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

        if (!File.Exists(fullPath))
            throw Error(ProviderErrorCode.NotFound, $"The source file '{fullPath}' does not exist.");

        var attributes = File.GetAttributes(fullPath);
        if ((attributes & FileAttributes.ReparsePoint) == FileAttributes.ReparsePoint)
            throw Error(ProviderErrorCode.AccessDenied,
                "Symlinks and reparse points are not accepted as document sources.");

        var extension = NormalizeExtension(Path.GetExtension(fullPath));
        if (!_allowedExtensions.Contains(extension))
            throw Error(ProviderErrorCode.ExtensionNotAllowed,
                $"Documents with extension '{extension}' are not allowed by this connection. Allowed: {string.Join(", ", _allowedExtensions)}.");

        return fullPath;
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

    private static string ContentTypeFor(string name) =>
        string.Equals(Path.GetExtension(name), ".docx", StringComparison.OrdinalIgnoreCase)
            ? "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            : "application/octet-stream";

    private DocumentProviderException Error(
        ProviderErrorCode code,
        string message,
        string? itemId = null,
        Exception? innerException = null) =>
        new(code, message, ProviderName, ConnectionId, itemId, innerException);

    private static string NewItemId() => Guid.NewGuid().ToString("N");

    // ── Hash / IO helpers ─────────────────────────────────────────────────────

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

    private static async Task WriteAtomicallyAsync(
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
            if (replace && File.Exists(destinationPath))
            {
                File.Replace(temporaryPath, destinationPath, null);
            }
            else
            {
                File.Move(temporaryPath, destinationPath);
            }
        }
        finally
        {
            if (File.Exists(temporaryPath)) File.Delete(temporaryPath);
        }
    }
}
