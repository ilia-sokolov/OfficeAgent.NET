using System.Diagnostics.CodeAnalysis;
using OfficeAgent.Abstractions;

namespace OfficeAgent.Core.DocumentProviders;

/// <summary>Contains readable document content and its canonical provider reference.</summary>
public sealed class DocumentContent : IDisposable
{
    /// <summary>Initializes readable provider content.</summary>
    public DocumentContent(DocumentReference reference, Stream stream)
    {
        Reference = reference ?? throw new ArgumentNullException(nameof(reference));
        Stream = stream ?? throw new ArgumentNullException(nameof(stream));
    }

    /// <summary>Gets the canonical reference, including the current provider version.</summary>
    public DocumentReference Reference { get; }

    /// <summary>Gets the readable content stream owned by this instance.</summary>
    public Stream Stream { get; }

    /// <inheritdoc />
    public void Dispose() => Stream.Dispose();
}

/// <summary>Opens and saves documents for one configured storage provider.</summary>
public interface IDocumentProvider
{
    /// <summary>Gets the provider type handled by this instance.</summary>
    string Provider { get; }

    /// <summary>Gets the configured connection identifier handled by this instance.</summary>
    string ConnectionId { get; }

    /// <summary>
    /// Registers an existing document with this connection and returns its canonical
    /// reference. The provider mints an opaque item id that callers use from then on;
    /// the provider stores only the reference (path, URL, drive id, …), not the bytes.
    /// <paramref name="source"/> is provider-specific - a filesystem path for the
    /// filesystem provider, a sharing URL or drive id for cloud providers.
    /// </summary>
    Task<DocumentReference> RegisterAsync(
        string source,
        CancellationToken cancellationToken = default);

    /// <summary>Opens a document and returns its canonical current reference.</summary>
    Task<DocumentContent> OpenReadAsync(
        DocumentReference reference,
        CancellationToken cancellationToken = default);

    /// <summary>Saves transformed content and returns the resulting canonical reference.</summary>
    Task<DocumentReference> SaveAsync(
        DocumentReference source,
        Stream content,
        SaveDocumentOptions options,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Removes the registration for a document from this connection. The underlying
    /// content the provider only referenced is left untouched; the host owns the
    /// file's lifecycle.
    /// </summary>
    Task RemoveAsync(
        DocumentReference reference,
        CancellationToken cancellationToken = default);
}

/// <summary>Thrown when a provider item changed after the caller selected or inspected it.</summary>
public sealed class DocumentVersionConflictException : DocumentProviderException
{
    /// <summary>Initializes a version-conflict exception.</summary>
    public DocumentVersionConflictException(string expectedVersion, string actualVersion)
        : this(expectedVersion, actualVersion, provider: "", connectionId: "", itemId: null)
    {
    }

    /// <summary>Initializes a version-conflict exception with provider context.</summary>
    public DocumentVersionConflictException(
        string expectedVersion,
        string actualVersion,
        string provider,
        string connectionId,
        string? itemId)
        : base(
            ProviderErrorCode.VersionConflict,
            $"Document version conflict: expected '{expectedVersion}', current version is '{actualVersion}'.",
            provider, connectionId, itemId)
    {
        ExpectedVersion = expectedVersion;
        ActualVersion = actualVersion;
    }

    /// <summary>Gets the version supplied by the caller.</summary>
    public string ExpectedVersion { get; }

    /// <summary>Gets the current provider version.</summary>
    public string ActualVersion { get; }
}

/// <summary>Resolves configured providers without exposing credentials to agent tool calls.</summary>
public sealed class DocumentProviderRegistry
{
    private readonly IReadOnlyList<IDocumentProvider> _providers;

    /// <summary>Initializes the registry from dependency-injection registrations.</summary>
    public DocumentProviderRegistry(IEnumerable<IDocumentProvider> providers) =>
        _providers = providers.ToArray();

    /// <summary>Returns the provider configured for a reference.</summary>
    public IDocumentProvider Resolve(DocumentReference reference)
    {
        if (reference is null) throw new ArgumentNullException(nameof(reference));

        var matches = _providers.Where(p =>
            string.Equals(p.Provider, reference.Provider, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(p.ConnectionId, reference.ConnectionId, StringComparison.Ordinal)).ToArray();

        return matches.Length switch
        {
            1 => matches[0],
            0 => throw new DocumentProviderException(
                ProviderErrorCode.ConfigurationError,
                $"No document provider is registered for '{reference.Provider}' connection '{reference.ConnectionId}'.",
                reference.Provider, reference.ConnectionId, reference.ItemId),
            _ => throw new DocumentProviderException(
                ProviderErrorCode.ConfigurationError,
                $"Multiple document providers are registered for '{reference.Provider}' connection '{reference.ConnectionId}'.",
                reference.Provider, reference.ConnectionId, reference.ItemId)
        };
    }

    /// <summary>
    /// Returns the provider configured for a connection id, regardless of provider type.
    /// Connection ids are host-chosen and must be unique across providers for callers
    /// that address documents by <c>(connectionId, documentId)</c> alone.
    /// </summary>
    public IDocumentProvider ResolveConnection(string connectionId)
    {
        if (string.IsNullOrWhiteSpace(connectionId))
            throw new DocumentProviderException(
                ProviderErrorCode.InvalidArgument,
                "A connection id is required.",
                provider: "", connectionId: connectionId ?? "", itemId: null);

        var matches = _providers.Where(p =>
            string.Equals(p.ConnectionId, connectionId, StringComparison.Ordinal)).ToArray();

        return matches.Length switch
        {
            1 => matches[0],
            0 => throw new DocumentProviderException(
                ProviderErrorCode.ConfigurationError,
                $"No document provider is registered for connection '{connectionId}'.",
                provider: "", connectionId, itemId: null),
            _ => throw new DocumentProviderException(
                ProviderErrorCode.ConfigurationError,
                $"Multiple document providers share connection id '{connectionId}'; address documents by full reference instead.",
                provider: "", connectionId, itemId: null)
        };
    }

    /// <summary>Returns the set of registered <c>(Provider, ConnectionId)</c> pairs for host enumeration.</summary>
    public IReadOnlyList<(string Provider, string ConnectionId)> Connections =>
        _providers.Select(p => (p.Provider, p.ConnectionId)).ToList();

    /// <summary>Returns <see langword="true"/> when a provider matches the supplied pair.</summary>
    public bool Contains(string provider, string connectionId) =>
        _providers.Any(p =>
            string.Equals(p.Provider, provider, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(p.ConnectionId, connectionId, StringComparison.Ordinal));
}

/// <summary>Contains a provider-backed apply result and the saved document reference.</summary>
public sealed class ProviderApplyResult
{
    /// <summary>Gets the deterministic OfficeAgent change report.</summary>
    public ChangeReport Report { get; init; } = new();

    /// <summary>
    /// Gets whether the plan committed and the provider save completed. When
    /// <see langword="true"/>, <see cref="Document"/> is non-null.
    /// </summary>
    [MemberNotNullWhen(true, nameof(Document))]
    public bool Committed { get; init; }

    /// <summary>Gets the saved provider reference when committed.</summary>
    public DocumentReference? Document { get; init; }
}
