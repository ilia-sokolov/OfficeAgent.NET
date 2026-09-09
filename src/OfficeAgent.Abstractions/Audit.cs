namespace OfficeAgent.Abstractions;

/// <summary>Identifies the author displayed on tracked document revisions.</summary>
public sealed class RevisionMetadata
{
    /// <summary>Gets the author written into revision markup.</summary>
    public string Author { get; init; } = "OfficeAgent";

    /// <summary>
    /// Gets the optional revision timestamp. When omitted, the engine resolves it once
    /// from the format module's clock and uses that value for the whole apply.
    /// </summary>
    public DateTimeOffset? TimestampUtc { get; init; }
}

/// <summary>
/// Identifies the authenticated actor on whose behalf a host applied a plan.
/// The host is responsible for deriving this value from its authentication context.
/// </summary>
public sealed class AuditActor
{
    /// <summary>Gets the stable actor subject, such as an OIDC <c>sub</c> claim.</summary>
    public string Subject { get; init; } = string.Empty;

    /// <summary>Gets the authority that issued the subject, when available.</summary>
    public string? Issuer { get; init; }

    /// <summary>Gets an optional human-readable actor name.</summary>
    public string? DisplayName { get; init; }
}

/// <summary>Describes how far an apply attempt progressed.</summary>
public enum ApplyOutcome
{
    /// <summary>The plan was valid and previewed without writing output.</summary>
    Previewed,

    /// <summary>The plan was rejected before an output document was produced.</summary>
    Rejected,

    /// <summary>The plan was applied and output bytes were produced.</summary>
    Committed
}

/// <summary>
/// Provides a tamper-evident summary of an apply attempt. Hashes are lowercase SHA-256
/// values over the effective plan JSON and the exact input/output document bytes.
/// </summary>
public sealed class ApplyReceipt
{
    /// <summary>Gets the receipt schema version.</summary>
    public string ReceiptVersion { get; init; } = "1";

    /// <summary>Gets the SHA-256 hash of the effective plan JSON.</summary>
    public string PlanSha256 { get; init; } = string.Empty;

    /// <summary>Gets the SHA-256 hash of the input document bytes.</summary>
    public string InputSha256 { get; init; } = string.Empty;

    /// <summary>Gets the SHA-256 hash of committed output bytes, if any.</summary>
    public string? OutputSha256 { get; init; }

    /// <summary>Gets the trusted engine timestamp for the apply attempt.</summary>
    public DateTimeOffset TimestampUtc { get; init; }

    /// <summary>Gets the apply outcome.</summary>
    public ApplyOutcome Outcome { get; init; }

    /// <summary>Gets the author and timestamp resolved for tracked revisions.</summary>
    public RevisionMetadata Revision { get; init; } = new();

    /// <summary>
    /// Gets the authenticated actor supplied by the host. This is deliberately separate
    /// from <see cref="Revision"/>, whose author is display metadata from the plan.
    /// </summary>
    public AuditActor? Actor { get; init; }

    /// <summary>Gets the provider reference after a provider save completes.</summary>
    public DocumentReference? OutputDocument { get; init; }
}
