namespace OfficeAgent.Core.DocumentProviders;

/// <summary>
/// Strongly typed category of a provider boundary failure. Wire-stable for
/// agents and logs; exhaustive for switch statements.
/// </summary>
public enum ProviderErrorCode
{
    /// <summary>Unspecified failure.</summary>
    Unknown,

    /// <summary>The requested item does not exist at the provider.</summary>
    NotFound,

    /// <summary>The reference is well-formed but the provider refused access (path traversal, symlink, wrong connection, etc.).</summary>
    AccessDenied,

    /// <summary>The document is larger than the provider's configured limit.</summary>
    ContentTooLarge,

    /// <summary>The item extension is not in the provider's allow-list.</summary>
    ExtensionNotAllowed,

    /// <summary>An optimistic-concurrency check failed; the item changed under the caller.</summary>
    VersionConflict,

    /// <summary>The supplied reference or options were structurally invalid for this provider.</summary>
    InvalidArgument,

    /// <summary>The provider registry is misconfigured (none or multiple matches).</summary>
    ConfigurationError,

    /// <summary>An underlying IO error occurred.</summary>
    IO
}

/// <summary>
/// Base exception for any failure at the document-provider boundary. Carries the
/// <see cref="Code"/>, <see cref="Provider"/>, <see cref="ConnectionId"/>, and (when
/// known) <see cref="ItemId"/> so a single <c>catch (DocumentProviderException ex)</c>
/// can route to a structured handler.
/// </summary>
public class DocumentProviderException : Exception
{
    /// <summary>Initializes a provider exception with full context.</summary>
    public DocumentProviderException(
        ProviderErrorCode code,
        string message,
        string provider = "",
        string connectionId = "",
        string? itemId = null,
        Exception? innerException = null)
        : base(message, innerException)
    {
        Code = code;
        Provider = provider;
        ConnectionId = connectionId;
        ItemId = itemId;
    }

    /// <summary>Gets the strongly typed error category.</summary>
    public ProviderErrorCode Code { get; }

    /// <summary>Gets the provider type (e.g. <c>filesystem</c>) when known.</summary>
    public string Provider { get; }

    /// <summary>Gets the configured connection id when known.</summary>
    public string ConnectionId { get; }

    /// <summary>Gets the offending item id when known.</summary>
    public string? ItemId { get; }
}
