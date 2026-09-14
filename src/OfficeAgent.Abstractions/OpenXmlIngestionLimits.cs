namespace OfficeAgent.Abstractions;

/// <summary>
/// Host-owned ceilings applied to every OOXML package this engine opens.
/// </summary>
/// <remarks>
/// <para>
/// These are resource ceilings, not a security sandbox. They bound what a single
/// malformed or hostile package can spend inside this process: how many bytes are read,
/// how far a package is allowed to expand, how many parts it may contain, and how deep
/// or how large its XML may be. They do not isolate the process, cap total memory across
/// concurrent calls, or replace operating-system limits.
/// </para>
/// <para>
/// The host owns these values. A caller may ask for something stricter with
/// <see cref="Restrict"/>, and every field is then taken as the smaller of the two, so a
/// request can lower an effective limit but can never raise a host ceiling.
/// </para>
/// </remarks>
public sealed class OpenXmlIngestionLimits
{
    /// <summary>The ceilings applied when a host configures nothing.</summary>
    public static OpenXmlIngestionLimits Default { get; } = new();

    /// <summary>
    /// Gets the maximum compressed package size accepted, in bytes. A stream is read up
    /// to this many bytes and refused if more arrive, so a non-seekable or endless source
    /// cannot be drained into memory.
    /// </summary>
    public long MaximumCompressedBytes { get; init; } = 128L * 1024 * 1024;

    /// <summary>
    /// Gets the maximum total expanded size of all package parts, in bytes. Measured
    /// while reading rather than taken from the archive's own metadata, which a hostile
    /// package controls.
    /// </summary>
    public long MaximumExpandedBytes { get; init; } = 512L * 1024 * 1024;

    /// <summary>Gets the maximum expanded size of any single package part, in bytes.</summary>
    public long MaximumPartBytes { get; init; } = 128L * 1024 * 1024;

    /// <summary>Gets the maximum number of entries a package may contain.</summary>
    public int MaximumParts { get; init; } = 4000;

    /// <summary>Gets the maximum element nesting depth accepted in package XML.</summary>
    public int MaximumXmlDepth { get; init; } = 256;

    /// <summary>Gets the maximum number of characters accepted in a single XML part.</summary>
    public long MaximumXmlCharacters { get; init; } = 64L * 1024 * 1024;

    /// <summary>
    /// Gets the maximum ratio of expanded to compressed bytes tolerated before a package
    /// is treated as a decompression bomb. Ordinary Office documents sit far below this.
    /// </summary>
    public int MaximumExpansionRatio { get; init; } = 200;

    /// <summary>
    /// Returns the effective limits for a request that asks for something stricter. Each
    /// field is the smaller of the two, so this never raises a host ceiling. A null
    /// request returns the host ceilings unchanged.
    /// </summary>
    public OpenXmlIngestionLimits Restrict(OpenXmlIngestionLimits? requested)
    {
        if (requested is null) return this;

        return new OpenXmlIngestionLimits
        {
            MaximumCompressedBytes = Math.Min(MaximumCompressedBytes, requested.MaximumCompressedBytes),
            MaximumExpandedBytes = Math.Min(MaximumExpandedBytes, requested.MaximumExpandedBytes),
            MaximumPartBytes = Math.Min(MaximumPartBytes, requested.MaximumPartBytes),
            MaximumParts = Math.Min(MaximumParts, requested.MaximumParts),
            MaximumXmlDepth = Math.Min(MaximumXmlDepth, requested.MaximumXmlDepth),
            MaximumXmlCharacters = Math.Min(MaximumXmlCharacters, requested.MaximumXmlCharacters),
            MaximumExpansionRatio = Math.Min(MaximumExpansionRatio, requested.MaximumExpansionRatio)
        };
    }

    /// <summary>Throws when any ceiling is not positive.</summary>
    public void Validate()
    {
        Positive(MaximumCompressedBytes, nameof(MaximumCompressedBytes));
        Positive(MaximumExpandedBytes, nameof(MaximumExpandedBytes));
        Positive(MaximumPartBytes, nameof(MaximumPartBytes));
        Positive(MaximumParts, nameof(MaximumParts));
        Positive(MaximumXmlDepth, nameof(MaximumXmlDepth));
        Positive(MaximumXmlCharacters, nameof(MaximumXmlCharacters));
        Positive(MaximumExpansionRatio, nameof(MaximumExpansionRatio));
    }

    private static void Positive(long value, string name)
    {
        if (value <= 0)
            throw new ArgumentOutOfRangeException(name, value, $"{name} must be positive.");
    }
}

/// <summary>
/// Raised when a package exceeds a host ingestion ceiling. Carries the limit that was hit
/// and the value that breached it so an operator can tune the ceiling or reject the input.
/// </summary>
public sealed class OpenXmlIngestionLimitException : Exception
{
    /// <summary>The stable wire code reported to agents and tools.</summary>
    public const string Code = "input-too-large";

    /// <summary>Initializes the exception for a named limit.</summary>
    public OpenXmlIngestionLimitException(string limit, long allowed, long observed, string? detail = null)
        : base(Describe(limit, allowed, observed, detail))
    {
        Limit = limit;
        Allowed = allowed;
        Observed = observed;
    }

    /// <summary>Gets the name of the ceiling that was exceeded.</summary>
    public string Limit { get; }

    /// <summary>Gets the configured ceiling.</summary>
    public long Allowed { get; }

    /// <summary>
    /// Gets the observed value, or -1 when the breach was detected before the true size
    /// was known, which is the point of refusing early.
    /// </summary>
    public long Observed { get; }

    private static string Describe(string limit, long allowed, long observed, string? detail)
    {
        var seen = observed < 0 ? "more" : observed.ToString();
        var suffix = string.IsNullOrEmpty(detail) ? string.Empty : $" {detail}";
        return $"The document exceeds the host ingestion limit {limit} ({allowed}); observed {seen}." +
               $"{suffix} Nothing was read beyond the limit and no output was produced.";
    }
}

/// <summary>
/// Raised when a package is structurally unacceptable: a traversing or duplicated entry
/// name, a corrupt archive, or XML this engine refuses to process.
/// </summary>
public sealed class OpenXmlPackageRejectedException : Exception
{
    /// <summary>The stable wire code reported to agents and tools.</summary>
    public const string Code = "malformed-package";

    /// <summary>Initializes the exception with the reason the package was refused.</summary>
    public OpenXmlPackageRejectedException(string message) : base(message)
    {
    }

    /// <summary>Initializes the exception with a reason and the underlying failure.</summary>
    public OpenXmlPackageRejectedException(string message, Exception inner) : base(message, inner)
    {
    }
}
