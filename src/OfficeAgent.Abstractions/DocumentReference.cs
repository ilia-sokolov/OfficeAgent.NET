namespace OfficeAgent.Abstractions;

/// <summary>
/// Identifies a document managed by a configured provider connection.
/// </summary>
/// <remarks>
/// References contain stable provider identifiers, not credentials. <see cref="Version"/>
/// is provider-defined and is used for optimistic concurrency checks.
/// </remarks>
public sealed class DocumentReference
{
    /// <summary>Gets the provider type, for example <c>filesystem</c> or <c>sharepoint</c>.</summary>
    public string Provider { get; init; } = string.Empty;

    /// <summary>Gets the host-configured connection identifier.</summary>
    public string ConnectionId { get; init; } = string.Empty;

    /// <summary>Gets the provider-defined item identifier.</summary>
    public string ItemId { get; init; } = string.Empty;

    /// <summary>Gets the provider-defined version, ETag, or content hash.</summary>
    public string? Version { get; init; }

    /// <summary>Gets the display file name when known.</summary>
    public string? Name { get; init; }

    /// <summary>Gets the media type when known.</summary>
    public string? ContentType { get; init; }

    /// <summary>
    /// Creates a reference to a document inside a filesystem provider connection.
    /// </summary>
    public static DocumentReference ForFileSystem(string connectionId, string itemId, string? version = null) => new()
    {
        Provider = "filesystem",
        ConnectionId = connectionId,
        ItemId = itemId,
        Version = version
    };

    /// <summary>
    /// Creates a reference for an arbitrary provider type. Prefer
    /// <see cref="ForFileSystem(string, string, string?)"/> when you know the provider.
    /// </summary>
    public static DocumentReference For(string provider, string connectionId, string itemId, string? version = null) => new()
    {
        Provider = provider,
        ConnectionId = connectionId,
        ItemId = itemId,
        Version = version
    };
}

/// <summary>Specifies how a provider saves transformed document content.</summary>
public enum SaveMode
{
    /// <summary>Create a non-destructive new version or sibling document.</summary>
    NewVersion,

    /// <summary>Create a separate document at the requested destination.</summary>
    NewDocument,

    /// <summary>Replace the source item after an optimistic concurrency check.</summary>
    Replace
}

/// <summary>Provides provider-independent options for saving transformed content.</summary>
public sealed class SaveDocumentOptions
{
    /// <summary>Gets the save behavior. The default is non-destructive.</summary>
    public SaveMode Mode { get; init; } = SaveMode.NewVersion;

    /// <summary>
    /// Gets the source version expected by the caller. When omitted, providers use the
    /// version carried by the source <see cref="DocumentReference"/>.
    /// </summary>
    public string? ExpectedVersion { get; init; }

    /// <summary>Gets an optional file name for a new version or document.</summary>
    public string? NewName { get; init; }

    /// <summary>
    /// Gets an optional provider-relative destination item id. Providers reject destinations
    /// outside their configured connection boundary.
    /// </summary>
    public string? DestinationItemId { get; init; }
}
