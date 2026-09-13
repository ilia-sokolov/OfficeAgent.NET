namespace OfficeAgent.Abstractions;

/// <summary>Host-enforced resource ceilings for Word assembly; never taken from a merge plan.</summary>
public sealed class DocumentMergeLimits
{
    /// <summary>Gets the maximum number of ordered input documents.</summary>
    public int MaximumSources { get; init; } = 20;
    /// <summary>Gets the maximum compressed size of one input.</summary>
    public long MaximumInputBytes { get; init; } = 64L * 1024 * 1024;
    /// <summary>Gets the maximum combined compressed input size.</summary>
    public long MaximumTotalInputBytes { get; init; } = 128L * 1024 * 1024;
    /// <summary>Gets the maximum combined expanded package size.</summary>
    public long MaximumExpandedBytes { get; init; } = 256L * 1024 * 1024;
    /// <summary>Gets the maximum combined input or output package entry count.</summary>
    public int MaximumParts { get; init; } = 4000;
    /// <summary>Gets the maximum compressed assembled output size.</summary>
    public long MaximumOutputBytes { get; init; } = 128L * 1024 * 1024;
}

/// <summary>Output metadata. Source formatting and next-page source boundaries are fixed policies.</summary>
public sealed class DocumentMergeOptions
{
    /// <summary>Gets an optional output title override.</summary>
    public string? Title { get; init; }
    /// <summary>Gets an optional output creator override.</summary>
    public string? Author { get; init; }
}

/// <summary>Ordered provider sources for an assembly preview.</summary>
public sealed class DocumentMergeRequest
{
    /// <summary>Gets provider sources in output order.</summary>
    public IReadOnlyList<DocumentReference> Sources { get; init; } = Array.Empty<DocumentReference>();
    /// <summary>Gets output metadata options.</summary>
    public DocumentMergeOptions Options { get; init; } = new();
}

/// <summary>One input, bound to exact bytes and its ordered position.</summary>
public sealed class DocumentMergeInput
{
    /// <summary>Gets the zero-based output order.</summary>
    public int Index { get; init; }
    /// <summary>Gets the provider source, or null for an in-memory workflow.</summary>
    public DocumentReference? Document { get; init; }
    /// <summary>Gets the exact lowercase SHA-256 input hash.</summary>
    public string Sha256 { get; init; } = string.Empty;
}

/// <summary>Serializable assembly intent. Its hash detects accidental changes, not authorization.</summary>
public sealed class DocumentMergePlan
{
    /// <summary>Gets the merge-plan schema version supported by this build.</summary>
    public const string CurrentVersion = "1";

    /// <summary>Gets the merge-plan schema version.</summary>
    public string Version { get; init; } = CurrentVersion;
    /// <summary>Gets ordered, hash-bound inputs.</summary>
    public IReadOnlyList<DocumentMergeInput> Inputs { get; init; } = Array.Empty<DocumentMergeInput>();
    /// <summary>Gets output metadata options bound into the plan hash.</summary>
    public DocumentMergeOptions Options { get; init; } = new();
    /// <summary>Gets the lowercase SHA-256 hash of the effective plan.</summary>
    public string PlanSha256 { get; init; } = string.Empty;
}

/// <summary>Source coverage and identifier transformations reported by preview.</summary>
public sealed class DocumentMergeSourceReport
{
    /// <summary>Gets the zero-based input order.</summary>
    public int Index { get; init; }
    /// <summary>Gets the paragraph count across imported content parts.</summary>
    public int Paragraphs { get; init; }
    /// <summary>Gets the main-document table count.</summary>
    public int Tables { get; init; }
    /// <summary>Gets the embedded raster image count.</summary>
    public int Images { get; init; }
    /// <summary>Gets the source section count.</summary>
    public int Sections { get; init; }
    /// <summary>Gets reported remapping and boundary decisions.</summary>
    public IReadOnlyList<string> Decisions { get; init; } = Array.Empty<string>();
}

/// <summary>Compatibility assessment after assembling and validating a candidate in memory.</summary>
public sealed class DocumentMergePreview
{
    /// <summary>Gets whether the candidate was fully supported and validated.</summary>
    public bool IsValid => Plan is not null && Diagnostics.Count == 0;
    /// <summary>Gets an executable plan when the candidate is valid.</summary>
    public DocumentMergePlan? Plan { get; init; }
    /// <summary>Gets source coverage reports.</summary>
    public IReadOnlyList<DocumentMergeSourceReport> Sources { get; init; } = Array.Empty<DocumentMergeSourceReport>();
    /// <summary>Gets blocking compatibility and resource diagnostics.</summary>
    public IReadOnlyList<WorkflowDiagnostic> Diagnostics { get; init; } = Array.Empty<WorkflowDiagnostic>();
}

/// <summary>Audit record for successful production of assembled bytes or a provider document.</summary>
public sealed class DocumentMergeReceipt
{
    /// <summary>Gets the merge-receipt schema version emitted by this build.</summary>
    public const string CurrentReceiptVersion = "1";

    /// <summary>Gets the receipt schema version.</summary>
    public string ReceiptVersion { get; init; } = CurrentReceiptVersion;
    /// <summary>Gets the exact effective plan hash.</summary>
    public string PlanSha256 { get; init; } = string.Empty;
    /// <summary>Gets ordered input references and hashes.</summary>
    public IReadOnlyList<DocumentMergeInput> Inputs { get; init; } = Array.Empty<DocumentMergeInput>();
    /// <summary>Gets the exact lowercase SHA-256 output hash.</summary>
    public string OutputSha256 { get; init; } = string.Empty;
    /// <summary>Gets the trusted assembly timestamp.</summary>
    public DateTimeOffset TimestampUtc { get; init; }
    /// <summary>Gets the host-authenticated actor, independently of output metadata.</summary>
    public AuditActor? Actor { get; init; }
    /// <summary>Gets the created provider document, when applicable.</summary>
    public DocumentReference? OutputDocument { get; init; }
}

/// <summary>Assembly result. Provider failures propagate with the provider's recovery semantics.</summary>
public sealed class DocumentMergeResult
{
    /// <summary>Gets whether assembled bytes were produced and, if applicable, saved.</summary>
    public bool Committed { get; init; }
    /// <summary>Gets the created provider document, when applicable.</summary>
    public DocumentReference? Document { get; init; }
    /// <summary>In-memory output only; provider results never include document bytes.</summary>
    public byte[]? Content { get; init; }
    /// <summary>Gets the multi-source receipt when committed.</summary>
    public DocumentMergeReceipt? Receipt { get; init; }
    /// <summary>Gets blocking diagnostics.</summary>
    public IReadOnlyList<WorkflowDiagnostic> Diagnostics { get; init; } = Array.Empty<WorkflowDiagnostic>();
}
