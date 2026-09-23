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
    IO,

    /// <summary>
    /// A document with the requested name already exists and the provider refuses to
    /// overwrite it. Distinct from <see cref="IO"/> because the caller fixes it by
    /// choosing another name rather than by retrying.
    /// </summary>
    AlreadyExists,

    /// <summary>
    /// The storage definitely did not accept the write: nothing was changed. A provider
    /// reports this only when it can guarantee it, for example after an atomic publish failed.
    /// Safe to retry once the cause is fixed.
    /// </summary>
    WriteRejected,

    /// <summary>
    /// The write may or may not have been stored: the provider failed, timed out or was
    /// cancelled after the write began. Do not retry blindly and do not report the document as
    /// unchanged; reopen the destination and compare it with the intended output first.
    /// </summary>
    OutcomeUnknown,

    /// <summary>
    /// The document was stored but its registration could not be persisted, so it exists
    /// without a document id. Do not create it again; register the known name instead.
    /// </summary>
    RegistrationFailed
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

/// <summary>
/// A failure after a storage write began, when the output may exist. Carries the path-free
/// locator a caller reconciles with before doing anything else.
/// </summary>
/// <remarks>
/// Two outcomes derive from it: <see cref="DocumentWriteOutcomeUnknownException"/>, where the
/// engine cannot tell whether the bytes landed, and <see cref="DocumentRegistrationFailedException"/>,
/// where they were stored but not registered. Catch this type to handle both. Neither is ever safe
/// to answer by writing again: find the output by <see cref="OutputName"/> and compare its content
/// hash with <see cref="OutputSha256"/> first.
/// </remarks>
public abstract class DocumentWriteRecoveryException : DocumentProviderException
{
    /// <summary>Initializes the locator shared by both post-write outcomes.</summary>
    private protected DocumentWriteRecoveryException(
        ProviderErrorCode code,
        string message,
        string provider,
        string connectionId,
        string? itemId,
        string? outputName,
        string outputSha256,
        Exception? innerException)
        : base(code, message, provider, connectionId, itemId, innerException)
    {
        // No validation here: DocumentWriteOutcomeUnknownException, frozen before this base
        // existed, never validated its arguments, and a base class must not change that.
        OutputName = outputName;
        OutputSha256 = outputSha256;
    }

    /// <summary>
    /// The name of the output in its connection, when the write was a creation or a new version.
    /// It is the name the caller asked for or the one the provider chose, never a storage path.
    /// </summary>
    public string? OutputName { get; }

    /// <summary>Lowercase SHA-256 of the exact bytes the write was carrying.</summary>
    public string OutputSha256 { get; }
}

/// <summary>
/// A storage write whose outcome is unknown: the provider failed, timed out or was cancelled
/// after the write call began. Carries what a caller needs to reconcile before retrying.
/// </summary>
/// <remarks>
/// The engine raises this for every failure inside a provider's save or create call that the
/// provider did not classify as certain. It never reports such a failure as nothing written.
/// To recover, reopen the destination and compare its content hash with
/// <see cref="DocumentWriteRecoveryException.OutputSha256"/>: equal means the write landed; the
/// source's prior version means it did not; anything else means another writer intervened and the
/// document must be inspected again.
/// </remarks>
public sealed class DocumentWriteOutcomeUnknownException : DocumentWriteRecoveryException
{
    /// <summary>Initializes the exception for one uncertain write.</summary>
    public DocumentWriteOutcomeUnknownException(
        string message,
        string provider,
        string connectionId,
        string? itemId,
        string? outputName,
        string outputSha256,
        Exception? innerException = null)
        : base(ProviderErrorCode.OutcomeUnknown, message, provider, connectionId, itemId,
            outputName, outputSha256, innerException)
    {
    }
}

/// <summary>
/// A storage write that landed but could not be registered: the bytes are stored under
/// <see cref="DocumentWriteRecoveryException.OutputName"/> and the connection cannot yet address
/// them by id. Its <see cref="DocumentProviderException.Code"/> is
/// <see cref="ProviderErrorCode.RegistrationFailed"/>.
/// </summary>
/// <remarks>
/// Recover by registering the named output, after confirming its content hash equals
/// <see cref="DocumentWriteRecoveryException.OutputSha256"/>. Writing it again would create a
/// duplicate beside the stored file.
/// </remarks>
public sealed class DocumentRegistrationFailedException : DocumentWriteRecoveryException
{
    /// <summary>Initializes the exception for one stored but unregistered output.</summary>
    public DocumentRegistrationFailedException(
        string message,
        string provider,
        string connectionId,
        string? itemId,
        string outputName,
        string outputSha256,
        Exception? innerException = null)
        : base(ProviderErrorCode.RegistrationFailed, message, provider, connectionId, itemId,
            outputName ?? throw new ArgumentNullException(nameof(outputName)),
            outputSha256 ?? throw new ArgumentNullException(nameof(outputSha256)), innerException)
    {
    }
}
