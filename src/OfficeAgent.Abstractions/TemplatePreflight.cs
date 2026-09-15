namespace OfficeAgent.Abstractions;

/// <summary>
/// What a template offers to bind, found by inspecting it rather than guessed.
/// </summary>
public sealed class TemplateDiscoveryResult
{
    /// <summary>Gets the scalar slots the template exposes, sorted by name.</summary>
    public IReadOnlyList<TemplateSlot> Slots { get; init; } = Array.Empty<TemplateSlot>();

    /// <summary>Gets the repeating-row candidates found in the template.</summary>
    public IReadOnlyList<RepeatingRowCandidate> RepeatingRows { get; init; } =
        Array.Empty<RepeatingRowCandidate>();

    /// <summary>
    /// Gets problems that would make a binding ambiguous, such as a tag used twice.
    /// A template with any of these can still be inspected, but binding it will fail.
    /// </summary>
    public IReadOnlyList<WorkflowDiagnostic> Diagnostics { get; init; } =
        Array.Empty<WorkflowDiagnostic>();

    /// <summary>Gets the SHA-256 of the template bytes this result describes.</summary>
    public string TemplateSha256 { get; init; } = string.Empty;
}

/// <summary>One scalar slot a template exposes.</summary>
public sealed class TemplateSlot
{
    /// <summary>Gets the content-control tag or PowerPoint shape name to bind by.</summary>
    public string Name { get; init; } = string.Empty;

    /// <summary>Gets the slot kind, <c>contentControl</c> or <c>shapeName</c>.</summary>
    public string Kind { get; init; } = string.Empty;

    /// <summary>Gets how many times this name occurs. More than one makes it unbindable.</summary>
    public int Occurrences { get; init; }

    /// <summary>
    /// Gets whether this slot can actually be bound. A duplicated name is reported so a
    /// caller can see it and fix the template, not so they can bind it anyway.
    /// </summary>
    public bool Bindable => Occurrences == 1;

    /// <summary>Gets the text currently in the slot, which is usually placeholder text.</summary>
    public string? CurrentText { get; init; }
}

/// <summary>
/// A table row that looks like a repeating template row.
/// </summary>
/// <remarks>
/// These are candidates, not confirmed structures: a row is nominated because it carries
/// <c>{{Field}}</c> placeholders, which is a convention rather than something the file
/// format declares. <see cref="Confirmed"/> says which is which, so a caller never
/// mistakes an inference for a guarantee.
/// </remarks>
public sealed class RepeatingRowCandidate
{
    /// <summary>Gets the inspected table path, such as <c>table#0</c>.</summary>
    public string TablePath { get; init; } = string.Empty;

    /// <summary>Gets the zero-based row index to use as the template row.</summary>
    public int TemplateRowIndex { get; init; }

    /// <summary>Gets the placeholder field names found in that row, sorted.</summary>
    public IReadOnlyList<string> Fields { get; init; } = Array.Empty<string>();

    /// <summary>
    /// Gets whether the row is a confirmed repeating structure rather than an inference
    /// from placeholder text. Nothing in WordprocessingML declares a repeating row, so
    /// this is false for placeholder-derived candidates.
    /// </summary>
    public bool Confirmed { get; init; }

    /// <summary>Gets the number of rows the table has, for context.</summary>
    public int TableRowCount { get; init; }
}

/// <summary>
/// Host-owned budgets for template batches.
/// </summary>
/// <remarks>
/// A request may ask for something stricter through <see cref="Restrict"/>; it can never
/// raise a host budget. Without this the request's own <c>MaximumDocuments</c> would be
/// the only ceiling, which is a caller setting its own limit.
/// </remarks>
public sealed class TemplateBatchLimits
{
    /// <summary>The budgets applied when a host configures nothing.</summary>
    public static TemplateBatchLimits Default { get; } = new();

    /// <summary>Gets the maximum outputs one batch may produce.</summary>
    public int MaximumDocuments { get; init; } = 100;

    /// <summary>Gets the maximum repeating rows one output may create.</summary>
    public int MaximumRowsPerDocument { get; init; } = 2_000;

    /// <summary>Gets the maximum scalar fields one output may bind.</summary>
    public int MaximumFieldsPerDocument { get; init; } = 500;

    /// <summary>Gets the maximum repeating rows the whole batch may create.</summary>
    public int MaximumTotalRows { get; init; } = 20_000;

    /// <summary>Gets the maximum characters any single bound value may carry.</summary>
    public int MaximumValueLength { get; init; } = 100_000;

    /// <summary>
    /// Gets the maximum diagnostics reported per item. Preflight reports every item's
    /// problems, but one pathological item must not produce an unbounded payload.
    /// </summary>
    public int MaximumDiagnosticsPerItem { get; init; } = 50;

    /// <summary>Returns the stricter of the host budgets and a request's own.</summary>
    public TemplateBatchLimits Restrict(TemplateBatchLimits? requested)
    {
        if (requested is null) return this;

        return new TemplateBatchLimits
        {
            MaximumDocuments = Math.Min(MaximumDocuments, requested.MaximumDocuments),
            MaximumRowsPerDocument = Math.Min(MaximumRowsPerDocument, requested.MaximumRowsPerDocument),
            MaximumFieldsPerDocument = Math.Min(MaximumFieldsPerDocument, requested.MaximumFieldsPerDocument),
            MaximumTotalRows = Math.Min(MaximumTotalRows, requested.MaximumTotalRows),
            MaximumValueLength = Math.Min(MaximumValueLength, requested.MaximumValueLength),
            MaximumDiagnosticsPerItem =
                Math.Min(MaximumDiagnosticsPerItem, requested.MaximumDiagnosticsPerItem)
        };
    }

    /// <summary>Throws when any budget is not positive.</summary>
    public void Validate()
    {
        Positive(MaximumDocuments, nameof(MaximumDocuments));
        Positive(MaximumRowsPerDocument, nameof(MaximumRowsPerDocument));
        Positive(MaximumFieldsPerDocument, nameof(MaximumFieldsPerDocument));
        Positive(MaximumTotalRows, nameof(MaximumTotalRows));
        Positive(MaximumValueLength, nameof(MaximumValueLength));
        Positive(MaximumDiagnosticsPerItem, nameof(MaximumDiagnosticsPerItem));
    }

    private static void Positive(int value, string name)
    {
        if (value <= 0)
            throw new ArgumentOutOfRangeException(name, value, $"{name} must be positive.");
    }
}

/// <summary>How far one batch item got.</summary>
public enum TemplateItemOutcome
{
    /// <summary>Validated only; no write was attempted.</summary>
    Previewed,

    /// <summary>Refused before any write. Nothing was created for this item.</summary>
    Failed,

    /// <summary>Not attempted, because an earlier item failed and the batch stops on error.</summary>
    Skipped,

    /// <summary>The output was created and registered.</summary>
    Committed,

    /// <summary>
    /// The provider was given the bytes and then failed to confirm. An output may or may
    /// not exist under this name. Do not retry it blindly; reconcile the destination.
    /// </summary>
    Uncertain
}

/// <summary>The result of validating a whole batch without writing anything.</summary>
public sealed class TemplateBatchPreview
{
    /// <summary>Gets whether every item validated.</summary>
    public bool IsValid { get; init; }

    /// <summary>Gets one entry per requested item, in request order.</summary>
    public IReadOnlyList<TemplateBatchPreviewItem> Items { get; init; } =
        Array.Empty<TemplateBatchPreviewItem>();

    /// <summary>Gets diagnostics about the batch as a whole, such as a duplicate name.</summary>
    public IReadOnlyList<WorkflowDiagnostic> Diagnostics { get; init; } =
        Array.Empty<WorkflowDiagnostic>();

    /// <summary>
    /// Gets the token binding this preview to the exact template bytes and the exact
    /// normalized batch. Pass it to the commit to refuse anything that has since changed.
    /// </summary>
    public TemplateBatchToken Token { get; init; } = new();

    /// <summary>Gets the effective budgets this preview was validated under.</summary>
    public TemplateBatchLimits Limits { get; init; } = TemplateBatchLimits.Default;
}

/// <summary>One item's preview outcome.</summary>
public sealed class TemplateBatchPreviewItem
{
    /// <summary>Gets the requested output name.</summary>
    public string OutputName { get; init; } = string.Empty;

    /// <summary>Gets whether this item validated.</summary>
    public bool IsValid { get; init; }

    /// <summary>Gets the operations the item would apply, for review.</summary>
    public int OperationCount { get; init; }

    /// <summary>Gets the repeating rows the item would create.</summary>
    public int RowCount { get; init; }

    /// <summary>Gets this item's diagnostics, capped by the effective budget.</summary>
    public IReadOnlyList<WorkflowDiagnostic> Diagnostics { get; init; } =
        Array.Empty<WorkflowDiagnostic>();

    /// <summary>
    /// Gets whether diagnostics were dropped to stay inside the cap, so a reader knows
    /// the list is partial rather than complete.
    /// </summary>
    public bool DiagnosticsTruncated { get; init; }
}

/// <summary>
/// Binds a validated batch to the bytes and intent it was validated against.
/// </summary>
/// <remarks>
/// Both halves matter. The template hash catches a template edited between preview and
/// commit; the batch hash catches a changed value, a renamed output, an added item or a
/// reordering. A commit presented with a token that no longer matches is refused rather
/// than silently applying intent nobody reviewed.
/// </remarks>
public sealed class TemplateBatchToken
{
    /// <summary>Gets the SHA-256 of the template bytes the preview validated.</summary>
    public string TemplateSha256 { get; init; } = string.Empty;

    /// <summary>Gets the SHA-256 of the deterministically normalized batch.</summary>
    public string BatchSha256 { get; init; } = string.Empty;
}
