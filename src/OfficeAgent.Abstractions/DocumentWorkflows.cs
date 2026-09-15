namespace OfficeAgent.Abstractions;

/// <summary>Controls how template slots omitted from one binding are handled.</summary>
public enum MissingTemplateValueBehavior
{
    /// <summary>Reject the binding.</summary>
    Fail,
    /// <summary>Leave the template slot or placeholder unchanged.</summary>
    Ignore,
    /// <summary>Replace the template slot or placeholder with an empty string.</summary>
    Empty
}

/// <summary>One repeating Word table binding.</summary>
public sealed class RepeatingTableBinding
{
    /// <summary>Gets the inspected Word table path, such as <c>table#0</c>.</summary>
    public string TablePath { get; init; } = string.Empty;
    /// <summary>Gets the zero-based row used as the template.</summary>
    public int TemplateRowIndex { get; init; }
    /// <summary>Gets the records used to create rows.</summary>
    public IReadOnlyList<IReadOnlyDictionary<string, string?>> Records { get; init; }
        = Array.Empty<IReadOnlyDictionary<string, string?>>();
}

/// <summary>Values applied to one template document.</summary>
public sealed class TemplateBinding
{
    /// <summary>Gets scalar values keyed by content-control tag or PowerPoint shape name.</summary>
    public IReadOnlyDictionary<string, string?> Values { get; init; }
        = new Dictionary<string, string?>(StringComparer.Ordinal);
    /// <summary>
    /// Gets typed values keyed by slot name: text, an image, or native chart data.
    /// </summary>
    /// <remarks>
    /// These sit beside <see cref="Values"/> rather than replacing them. A name may
    /// appear in one or the other, not both. A typed value aimed at a slot that cannot
    /// hold it is refused, never stringified and never dropped.
    /// </remarks>
    public IReadOnlyDictionary<string, TemplateValue> TypedValues { get; init; }
        = new Dictionary<string, TemplateValue>(StringComparer.Ordinal);

    /// <summary>Gets repeating Word table bindings.</summary>
    public IReadOnlyList<RepeatingTableBinding> RepeatingTables { get; init; }
        = Array.Empty<RepeatingTableBinding>();
    /// <summary>Gets the displayed Word revision identity for generated operations.</summary>
    public RevisionMetadata? Revision { get; init; }
    /// <summary>Gets how missing scalar slots and repeating placeholders are handled.</summary>
    public MissingTemplateValueBehavior MissingValueBehavior { get; init; }
        = MissingTemplateValueBehavior.Fail;
    /// <summary>Gets whether supplied scalar and row fields absent from the template are rejected.</summary>
    public bool RejectUnknownValues { get; init; } = true;
    /// <summary>Gets how Word changes are recorded. Template generation defaults to direct edits.</summary>
    public ChangeMode Mode { get; init; } = ChangeMode.Direct;
}

/// <summary>One requested output in a template batch.</summary>
public sealed class TemplateBatchItem
{
    /// <summary>Gets the new document name.</summary>
    public string OutputName { get; init; } = string.Empty;
    /// <summary>Gets the values for this output.</summary>
    public TemplateBinding Binding { get; init; } = new();
}

/// <summary>A bounded request to populate several outputs from one template.</summary>
public sealed class TemplateBatchRequest
{
    /// <summary>Gets the requested outputs.</summary>
    public IReadOnlyList<TemplateBatchItem> Items { get; init; } = Array.Empty<TemplateBatchItem>();
    /// <summary>Gets whether later items run after one item fails.</summary>
    public bool ContinueOnError { get; init; } = true;
    /// <summary>
    /// Gets the maximum accepted item count requested by the caller.
    /// </summary>
    /// <remarks>
    /// This is a request, not a ceiling. The effective limit is the stricter of this and
    /// the host's <see cref="TemplateBatchLimits"/>, so raising it here cannot raise the
    /// host's budget.
    /// </remarks>
    public int MaximumDocuments { get; init; } = 100;

    /// <summary>
    /// Gets budgets this request asks to be held to, which are intersected with the
    /// host's. Omit it to run under the host budgets alone.
    /// </summary>
    public TemplateBatchLimits? Limits { get; init; }
}

/// <summary>A stable diagnostic returned by a higher-level document workflow.</summary>
public sealed class WorkflowDiagnostic
{
    /// <summary>Gets the stable machine-readable code.</summary>
    public string Code { get; init; } = string.Empty;
    /// <summary>Gets the human-readable explanation.</summary>
    public string Message { get; init; } = string.Empty;
    /// <summary>Gets the affected slot, table, paragraph, or document area.</summary>
    public string? Path { get; init; }
}

/// <summary>Result of resolving template values into a document plan.</summary>
public sealed class TemplatePlanResult
{
    /// <summary>Gets whether every binding resolved unambiguously.</summary>
    public bool IsValid => Diagnostics.Count == 0;
    /// <summary>Gets the applicable plan when valid.</summary>
    public DocumentPlan? Plan { get; init; }
    /// <summary>Gets binding diagnostics.</summary>
    public IReadOnlyList<WorkflowDiagnostic> Diagnostics { get; init; } = Array.Empty<WorkflowDiagnostic>();
}

/// <summary>Result for one output in a template batch.</summary>
public sealed class TemplateBatchItemResult
{
    /// <summary>Gets the requested output name.</summary>
    public string OutputName { get; init; } = string.Empty;
    /// <summary>Gets whether the provider saved the output.</summary>
    public bool Committed { get; init; }

    /// <summary>
    /// Gets how far this item got. <see cref="Committed"/> stays the simple question
    /// "is there an output"; this distinguishes a refusal from an item never attempted
    /// and from a write whose fate is unknown.
    /// </summary>
    public TemplateItemOutcome Outcome { get; init; } = TemplateItemOutcome.Failed;
    /// <summary>Gets the saved output reference.</summary>
    public DocumentReference? Document { get; init; }
    /// <summary>Gets the apply report when a plan reached preview or commit.</summary>
    public ChangeReport? Report { get; init; }
    /// <summary>Gets the apply receipt when one was produced.</summary>
    public ApplyReceipt? Receipt { get; init; }
    /// <summary>Gets binding or provider diagnostics.</summary>
    public IReadOnlyList<WorkflowDiagnostic> Diagnostics { get; init; } = Array.Empty<WorkflowDiagnostic>();
}

/// <summary>Result of a template batch.</summary>
public sealed class TemplateBatchResult
{
    /// <summary>Gets one result per attempted item.</summary>
    public IReadOnlyList<TemplateBatchItemResult> Items { get; init; } = Array.Empty<TemplateBatchItemResult>();
    /// <summary>Gets whether every requested item committed.</summary>
    public bool Committed => Items.Count > 0 && Items.All(item => item.Committed);
}

/// <summary>What a comparison was able to say about one area of the documents.</summary>
public enum ComparisonAreaState
{
    /// <summary>The area was compared and is the same in both documents.</summary>
    Unchanged,

    /// <summary>The area was compared, differs, and the difference is represented in the plan.</summary>
    Changed,

    /// <summary>
    /// The area differs and this comparison cannot represent the difference, so no plan is
    /// offered for the whole comparison.
    /// </summary>
    Blocked,

    /// <summary>
    /// The area was not examined, so nothing is known about it. This is never the same as
    /// unchanged: silence about an area is not evidence that it matches.
    /// </summary>
    NotCompared
}

/// <summary>One area of the documents and what the comparison could say about it.</summary>
public sealed class ComparisonArea
{
    /// <summary>Gets the area name, for example <c>bodyParagraphs</c> or <c>images</c>.</summary>
    public string Name { get; init; } = string.Empty;

    /// <summary>Gets what the comparison established about this area.</summary>
    public ComparisonAreaState State { get; init; }

    /// <summary>
    /// Gets the diagnostic code explaining a blocked area, or null when there is nothing
    /// to explain.
    /// </summary>
    public string? Code { get; init; }

    /// <summary>Gets a short human-readable note, when one helps.</summary>
    public string? Detail { get; init; }
}

/// <summary>Kind of paragraph difference found by a comparison.</summary>
public enum DocumentDifferenceKind
{
    /// <summary>A paragraph exists only in the revised document.</summary>
    Added,
    /// <summary>A paragraph exists only in the original document.</summary>
    Removed,
    /// <summary>An aligned paragraph has different text.</summary>
    Changed
}

/// <summary>One structured paragraph difference.</summary>
public sealed class DocumentDifference
{
    /// <summary>Gets the difference kind.</summary>
    public DocumentDifferenceKind Kind { get; init; }
    /// <summary>Gets the zero-based original body-paragraph index, when present.</summary>
    public int? OriginalIndex { get; init; }
    /// <summary>Gets the zero-based revised body-paragraph index, when present.</summary>
    public int? RevisedIndex { get; init; }
    /// <summary>Gets the original text.</summary>
    public string? Before { get; init; }
    /// <summary>Gets the revised text.</summary>
    public string? After { get; init; }
}

/// <summary>Bounds and authorship for Word comparison.</summary>
public sealed class DocumentComparisonOptions
{
    /// <summary>Gets the displayed author and timestamp used by the proposed redline.</summary>
    public RevisionMetadata? Revision { get; init; }
    /// <summary>Gets the maximum body paragraph count accepted per document.</summary>
    public int MaximumParagraphs { get; init; } = 2000;
    /// <summary>Gets the maximum exact input size per document.</summary>
    public long MaximumDocumentBytes { get; init; } = 64L * 1024 * 1024;
    /// <summary>Gets the maximum differences returned.</summary>
    public int MaximumDifferences { get; init; } = 1000;
}

/// <summary>A read-only Word comparison and, when coverage is complete, an applicable redline plan.</summary>
public sealed class DocumentComparisonResult
{
    /// <summary>Gets whether every detected change is covered by the proposed plan.</summary>
    public bool IsComplete => Diagnostics.Count == 0;
    /// <summary>Gets the SHA-256 hash of the exact original bytes.</summary>
    public string OriginalSha256 { get; init; } = string.Empty;
    /// <summary>Gets the SHA-256 hash of the exact revised bytes.</summary>
    public string RevisedSha256 { get; init; } = string.Empty;
    /// <summary>Gets structured paragraph differences.</summary>
    public IReadOnlyList<DocumentDifference> Differences { get; init; } = Array.Empty<DocumentDifference>();
    /// <summary>Gets unsupported-area and resource diagnostics.</summary>
    public IReadOnlyList<WorkflowDiagnostic> Diagnostics { get; init; } = Array.Empty<WorkflowDiagnostic>();

    /// <summary>
    /// Gets what the comparison established about each area of the documents, in a stable
    /// order.
    /// </summary>
    /// <remarks>
    /// This is the difference between "these documents match" and "this comparison did not
    /// look". An area reported <see cref="ComparisonAreaState.NotCompared"/> is unknown,
    /// and treating it as unchanged is exactly the inference this result refuses to make
    /// on a caller's behalf.
    /// </remarks>
    public IReadOnlyList<ComparisonArea> Coverage { get; init; } = Array.Empty<ComparisonArea>();

    /// <summary>
    /// Gets a snapshot-bound tracked-change plan when coverage is complete.
    /// </summary>
    /// <remarks>
    /// Null whenever any area is blocked. There is no partial plan and no way to ask for
    /// one: applying a subset of detected changes would silently drop the rest.
    /// </remarks>
    public DocumentPlan? Plan { get; init; }
}
