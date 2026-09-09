using System.Collections.Concurrent;
using System.Security.Cryptography;
using OfficeAgent.Abstractions;

namespace OfficeAgent.Core.DocumentProviders;

/// <summary>
/// Holds documents in the server's own memory for the life of the process, so an agent can
/// create and edit one without any storage being configured and without the bytes ever
/// travelling through the conversation.
/// </summary>
/// <remarks>
/// <para>
/// This is the answer to the one thing passing documents inline cannot do: chaining. An
/// agent that must hand the whole document back to make a second edit has to reproduce it
/// exactly, and a model reproducing several thousand characters of base64 gets a character
/// wrong often enough to matter - after which every later call fails on content that is no
/// longer a package. Here the agent passes a short opaque id instead, and the bytes stay on
/// this side of the conversation.
/// </para>
/// <para>
/// Nothing is persisted. The documents live as long as the provider instance does, which
/// for a stdio MCP server is the length of one session; when the process ends they are
/// gone. That is the intended lifetime, not a limitation to work around: a host that needs
/// the result keeps it by exporting the bytes or by configuring real storage.
/// </para>
/// <para>
/// Because the store is memory, its cap is a real one. <see cref="MemoryDocumentProviderOptions.MaximumTotalBytes"/>
/// bounds everything the connection holds at once, so a long session cannot grow until the
/// host runs out of memory.
/// </para>
/// </remarks>
public sealed class MemoryDocumentProvider : IDocumentProvider, IDocumentCreatingProvider, IConnectionEditingDefaults
{
    /// <summary>The provider name reported on references and errors.</summary>
    public const string ProviderName = "memory";

    private readonly ConcurrentDictionary<string, Entry> _documents = new(StringComparer.Ordinal);
    private readonly MemoryDocumentProviderOptions _options;

    /// <summary>Initializes a provider for one connection.</summary>
    public MemoryDocumentProvider(string connectionId, MemoryDocumentProviderOptions? options = null)
    {
        if (string.IsNullOrWhiteSpace(connectionId))
            throw new ArgumentException("A connection id is required.", nameof(connectionId));

        ConnectionId = connectionId;
        _options = options ?? new MemoryDocumentProviderOptions();
    }

    /// <inheritdoc />
    public string Provider => ProviderName;

    /// <inheritdoc />
    public string ConnectionId { get; }

    /// <inheritdoc />
    public ChangeMode DefaultChangeMode => _options.DefaultChangeMode;

    /// <summary>Gets the number of documents currently held.</summary>
    public int Count => _documents.Count;

    /// <summary>Gets the total size of everything currently held, in bytes.</summary>
    public long TotalBytes => _documents.Values.Sum(entry => (long)entry.Bytes.Length);

    /// <summary>
    /// Adds a document the caller already holds the bytes of, and returns its reference.
    /// This is how content enters an ephemeral connection - there is no external store to
    /// register a path against.
    /// </summary>
    public DocumentReference Add(string name, byte[] content)
    {
        if (content is null) throw new ArgumentNullException(nameof(content));

        var safeName = ValidateName(name);
        Require(content.LongLength <= _options.MaximumBytes, ProviderErrorCode.ContentTooLarge,
            $"Document exceeds the configured maximum of {_options.MaximumBytes} bytes.", itemId: null);
        RequireRoom(content.LongLength, replacing: 0);

        var itemId = Guid.NewGuid().ToString("N");
        var entry = new Entry(safeName, content, Version(content));
        _documents[itemId] = entry;

        return Reference(itemId, entry);
    }

    /// <summary>Returns the bytes held under an item id.</summary>
    public byte[] Read(string itemId) => Locate(itemId).Bytes;

    /// <summary>Returns the current reference for an item id, without reading its bytes.</summary>
    public DocumentReference Describe(string itemId) => Reference(itemId, Locate(itemId));

    /// <summary>Lists what the connection currently holds, newest first is not implied.</summary>
    public IReadOnlyList<DocumentReference> List() =>
        _documents.Select(pair => Reference(pair.Key, pair.Value)).ToList();

    /// <inheritdoc />
    /// <remarks>
    /// An ephemeral connection has nothing to register <em>against</em>: there is no path
    /// or URL that names a document the host already has. Content arrives through
    /// <see cref="Add"/> or <see cref="CreateAsync"/> instead, and saying so plainly is
    /// better than accepting a source and inventing a meaning for it.
    /// </remarks>
    public Task<DocumentReference> RegisterAsync(string source, CancellationToken cancellationToken = default) =>
        throw Error(ProviderErrorCode.InvalidArgument,
            $"The '{ConnectionId}' connection keeps documents in memory and has no external source to register. " +
            "Create a document in it, or supply the content directly, instead of naming a path or URL.",
            itemId: null);

    /// <inheritdoc />
    public Task<DocumentContent> OpenReadAsync(
        DocumentReference reference, CancellationToken cancellationToken = default)
    {
        if (reference is null) throw new ArgumentNullException(nameof(reference));
        cancellationToken.ThrowIfCancellationRequested();

        var entry = Locate(reference.ItemId);
        if (!string.IsNullOrEmpty(reference.Version) &&
            !string.Equals(reference.Version, entry.Version, StringComparison.Ordinal))
            throw new DocumentVersionConflictException(
                reference.Version!, entry.Version, ProviderName, ConnectionId, reference.ItemId);

        return Task.FromResult(new DocumentContent(
            Reference(reference.ItemId, entry),
            new MemoryStream(entry.Bytes, writable: false)));
    }

    /// <inheritdoc />
    public async Task<DocumentReference> SaveAsync(
        DocumentReference source,
        Stream content,
        SaveDocumentOptions options,
        CancellationToken cancellationToken = default)
    {
        if (source is null) throw new ArgumentNullException(nameof(source));
        if (content is null) throw new ArgumentNullException(nameof(content));
        if (options is null) throw new ArgumentNullException(nameof(options));

        Require(string.IsNullOrEmpty(options.DestinationItemId), ProviderErrorCode.InvalidArgument,
            "Item ids are opaque and provider-assigned; a destination id cannot be chosen by the caller.",
            source.ItemId);
        Require(options.Mode != SaveMode.Replace || string.IsNullOrWhiteSpace(options.NewName),
            ProviderErrorCode.InvalidArgument,
            "Replace mode overwrites the source in place and cannot rename.", source.ItemId);

        var existing = Locate(source.ItemId);

        var expectedVersion = options.ExpectedVersion ?? source.Version;
        if (!string.IsNullOrEmpty(expectedVersion) &&
            !string.Equals(expectedVersion, existing.Version, StringComparison.Ordinal))
            throw new DocumentVersionConflictException(
                expectedVersion!, existing.Version, ProviderName, ConnectionId, source.ItemId);

        var bytes = await ReadAllAsync(content, cancellationToken).ConfigureAwait(false);
        Require(bytes.LongLength <= _options.MaximumBytes, ProviderErrorCode.ContentTooLarge,
            $"Document exceeds the configured maximum of {_options.MaximumBytes} bytes.", source.ItemId);

        if (options.Mode == SaveMode.Replace)
        {
            RequireRoom(bytes.LongLength, replacing: existing.Bytes.LongLength);
            var replaced = new Entry(existing.Name, bytes, Version(bytes));
            _documents[source.ItemId] = replaced;
            return Reference(source.ItemId, replaced);
        }

        RequireRoom(bytes.LongLength, replacing: 0);
        var name = string.IsNullOrWhiteSpace(options.NewName)
            ? NextVersionedName(existing.Name)
            : ValidateName(options.NewName!);

        var itemId = Guid.NewGuid().ToString("N");
        var created = new Entry(name, bytes, Version(bytes));
        _documents[itemId] = created;
        return Reference(itemId, created);
    }

    /// <inheritdoc />
    /// <remarks>
    /// Unlike a provider over external storage, removing here really does discard the
    /// content: the store is the only copy. That is the point of an ephemeral connection,
    /// and it is what lets an agent clear a working document it is finished with.
    /// </remarks>
    public Task RemoveAsync(DocumentReference reference, CancellationToken cancellationToken = default)
    {
        if (reference is null) throw new ArgumentNullException(nameof(reference));
        _documents.TryRemove(reference.ItemId, out _);
        return Task.CompletedTask;
    }

    /// <inheritdoc />
    public async Task<DocumentReference> CreateAsync(
        string name, Stream content, CancellationToken cancellationToken = default)
    {
        if (content is null) throw new ArgumentNullException(nameof(content));

        var safeName = ValidateName(name);
        Require(!_documents.Values.Any(entry => string.Equals(entry.Name, safeName, StringComparison.OrdinalIgnoreCase)),
            ProviderErrorCode.AlreadyExists,
            $"A document named '{safeName}' is already open in this connection.", itemId: null);

        var bytes = await ReadAllAsync(content, cancellationToken).ConfigureAwait(false);
        return Add(safeName, bytes);
    }

    // ── Internals ────────────────────────────────────────────────────────

    private sealed record Entry(string Name, byte[] Bytes, string Version);

    private Entry Locate(string itemId) =>
        _documents.TryGetValue(itemId, out var entry)
            ? entry
            : throw Error(ProviderErrorCode.NotFound,
                $"No document with id '{itemId}' is open in this connection. " +
                "Documents here live only for the session, so an id from an earlier run no longer resolves.",
                itemId);

    private DocumentReference Reference(string itemId, Entry entry) => new()
    {
        Provider = ProviderName,
        ConnectionId = ConnectionId,
        ItemId = itemId,
        Name = entry.Name,
        ContentType = OfficeContentTypes.ForName(entry.Name),
        Version = entry.Version
    };

    private string ValidateName(string name)
    {
        Require(!string.IsNullOrWhiteSpace(name), ProviderErrorCode.InvalidArgument,
            "A document name is required, including its extension (for example 'report.docx').", itemId: null);

        var trimmed = name.Trim();
        Require(
            trimmed.IndexOfAny(InvalidNameChars) < 0 && !trimmed.Any(char.IsControl),
            ProviderErrorCode.InvalidArgument,
            "A document name must be a bare file name without path separators or control characters.",
            itemId: null);

        var extension = Path.GetExtension(trimmed);
        Require(
            _options.AllowedExtensions.Count == 0 ||
            _options.AllowedExtensions.Any(allowed => string.Equals(allowed, extension, StringComparison.OrdinalIgnoreCase)),
            ProviderErrorCode.InvalidArgument,
            $"'{extension}' is not an allowed extension for this connection. Allowed: {string.Join(", ", _options.AllowedExtensions)}.",
            itemId: null);

        return trimmed;
    }

    /// <summary>
    /// Keeps the connection inside its total budget. The store is memory, so an
    /// unbounded one is a leak with a friendly name.
    /// </summary>
    private void RequireRoom(long incoming, long replacing)
    {
        var projected = TotalBytes - replacing + incoming;
        Require(projected <= _options.MaximumTotalBytes, ProviderErrorCode.ContentTooLarge,
            $"The '{ConnectionId}' connection holds documents in memory and is limited to " +
            $"{_options.MaximumTotalBytes} bytes in total; this would take it to {projected}. " +
            "Remove a document you have finished with, or configure a larger limit.",
            itemId: null);
    }

    /// <summary>A sibling name for a NewVersion save, matching the filesystem provider's shape.</summary>
    private string NextVersionedName(string name)
    {
        var stem = Path.GetFileNameWithoutExtension(name);
        var extension = Path.GetExtension(name);

        for (var version = 2; ; version++)
        {
            var candidate = $"{stem}.v{version}{extension}";
            if (!_documents.Values.Any(entry => string.Equals(entry.Name, candidate, StringComparison.OrdinalIgnoreCase)))
                return candidate;
        }
    }

    private static async Task<byte[]> ReadAllAsync(Stream content, CancellationToken cancellationToken)
    {
        if (content is MemoryStream ready) return ready.ToArray();

        using var buffer = new MemoryStream();
        await content.CopyToAsync(buffer, 81920, cancellationToken).ConfigureAwait(false);
        return buffer.ToArray();
    }

    private static string Version(byte[] bytes)
    {
        using var sha = SHA256.Create();
        return BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", string.Empty);
    }

    private void Require(bool condition, ProviderErrorCode code, string message, string? itemId)
    {
        if (!condition) throw Error(code, message, itemId);
    }

    private DocumentProviderException Error(ProviderErrorCode code, string message, string? itemId) =>
        new(code, message, ProviderName, ConnectionId, itemId);

    private static readonly char[] InvalidNameChars =
        { '<', '>', ':', '"', '|', '?', '*', '/', '\\', '\0' };
}

/// <summary>Options for a <see cref="MemoryDocumentProvider"/> connection.</summary>
public sealed class MemoryDocumentProviderOptions
{
    /// <summary>Gets or sets the largest single document the connection accepts. Defaults to 25 MB.</summary>
    public long MaximumBytes { get; set; } = 25L * 1024 * 1024;

    /// <summary>
    /// Gets or sets the total the connection may hold at once. Defaults to 100 MB - the
    /// store is process memory, so this bound is what keeps a long session from growing
    /// without limit.
    /// </summary>
    public long MaximumTotalBytes { get; set; } = 100L * 1024 * 1024;

    /// <summary>Gets or sets the extensions the connection accepts. Empty means any.</summary>
    public IList<string> AllowedExtensions { get; set; } = new List<string> { ".docx", ".pptx", ".xlsx" };

    /// <summary>
    /// Gets or sets the change mode for an operation that does not state one. Defaults to
    /// <see cref="ChangeMode.Tracked"/>, as every other connection does.
    /// </summary>
    public ChangeMode DefaultChangeMode { get; set; } = ChangeMode.Tracked;
}
