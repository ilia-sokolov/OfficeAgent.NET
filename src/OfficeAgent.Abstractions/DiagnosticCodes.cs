namespace OfficeAgent.Abstractions;

// Stable failure and diagnostic codes, catalogued so hosts can switch on constants rather than
// strings and so the public API baseline freezes every value. Each value is the exact string
// already on the wire; cataloguing moved no bytes. Plan validation codes live in
// ValidationErrorCodes, and ingestion refusals in OpenXmlIngestionLimitException.Code and
// OpenXmlPackageRejectedException.Code.

/// <summary>
/// Codes an agent tool returns in its error envelope. Provider failures surface here with the
/// code that corresponds to their <c>OfficeAgent.Core.DocumentProviders.ProviderErrorCode</c>.
/// </summary>
public static class ToolErrorCodes
{
    /// <summary>The caller lacks permission for the document or connection.</summary>
    public const string AccessDenied = "access-denied";

    /// <summary>A document with that name already exists where it would have been created.</summary>
    public const string AlreadyExists = "already-exists";

    /// <summary>The host's connection or provider configuration is invalid.</summary>
    public const string ConfigurationError = "configuration-error";

    /// <summary>
    /// The call was cancelled. It does not say nothing was written: a cancellation that arrives
    /// after storage accepted the bytes leaves the write in place.
    /// </summary>
    public const string Cancelled = "cancelled";

    /// <summary>The connection access policy refused the capability the tool needs.</summary>
    public const string ConnectionForbidden = "connection-forbidden";

    /// <summary>The document or supplied content exceeds a size limit.</summary>
    public const string ContentTooLarge = "content-too-large";

    /// <summary>The connection does not accept documents with that extension.</summary>
    public const string ExtensionNotAllowed = "extension-not-allowed";

    /// <summary>An unexpected internal failure. Details are withheld from the caller.</summary>
    public const string InternalError = "internal-error";

    /// <summary>An argument was missing, empty or out of range.</summary>
    public const string InvalidArgument = "invalid-argument";

    /// <summary>
    /// A JSON argument could not be read, or carried an unknown property, operation or enum value.
    /// </summary>
    public const string InvalidJson = "invalid-json";

    /// <summary>The provider failed to read or write storage.</summary>
    public const string IOError = "io-error";

    /// <summary>The connection or document does not exist.</summary>
    public const string NotFound = "not-found";

    /// <summary>A provider failure that has no more specific code.</summary>
    public const string ProviderError = "provider-error";

    /// <summary>A regular expression search exceeded its time limit.</summary>
    public const string RegexTimeout = "regex-timeout";

    /// <summary>The document changed since the version the caller supplied.</summary>
    public const string VersionConflict = "version-conflict";
}

/// <summary>Diagnostics from template discovery, batch preflight and population.</summary>
public static class TemplateDiagnosticCodes
{
    /// <summary>A template slot occurs more than once, so a binding to it is ambiguous.</summary>
    public const string AmbiguousTemplateSlot = "ambiguous-template-slot";

    /// <summary>The batch requests more outputs than the effective limit.</summary>
    public const string BatchTooLarge = "batch-too-large";

    /// <summary>A chart binding carries more data points than the effective limit.</summary>
    public const string ChartTooLarge = "chart-too-large";

    /// <summary>Two batch items request the same output name.</summary>
    public const string DuplicateOutputName = "duplicate-output-name";

    /// <summary>A slot is bound both as a scalar value and as a typed value.</summary>
    public const string DuplicateTemplateBinding = "duplicate-template-binding";

    /// <summary>The image bound to a slot could not be read.</summary>
    public const string ImageNotAvailable = "image-not-available";

    /// <summary>An image binding exceeds the effective per-image size limit.</summary>
    public const string ImageTooLarge = "image-too-large";

    /// <summary>The bound bytes are not the declared image type.</summary>
    public const string ImageTypeMismatch = "image-type-mismatch";

    /// <summary>An image binding needs exactly one of inline bytes or a provider document id.</summary>
    public const string InvalidImageBinding = "invalid-image-binding";

    /// <summary>
    /// The item was never attempted because an earlier item failed in a batch that stops on
    /// error. The output does not exist.
    /// </summary>
    public const string ItemSkipped = "item-skipped";

    /// <summary>An image slot sits inside a repeating table row, which is not supported.</summary>
    public const string MediaInRepeatingRow = "media-in-repeating-row";

    /// <summary>An image binding has no alt text.</summary>
    public const string MissingAltText = "missing-alt-text";

    /// <summary>A batch item has no output name.</summary>
    public const string MissingOutputName = "missing-output-name";

    /// <summary>No value was supplied for a template slot and missing values are refused.</summary>
    public const string MissingTemplateValue = "missing-template-value";

    /// <summary>
    /// The template or the batch changed since the preview that issued the supplied token.
    /// </summary>
    public const string StaleBatchPreview = "stale-batch-preview";

    /// <summary>No paragraph in the template can host the bound image.</summary>
    public const string TemplateAnchorNotFound = "template-anchor-not-found";

    /// <summary>No inspected table has the requested repeating-table path.</summary>
    public const string TemplateTableNotFound = "template-table-not-found";

    /// <summary>An output binds more scalar fields than the effective limit.</summary>
    public const string TooManyFields = "too-many-fields";

    /// <summary>An output binds more images than the effective limit.</summary>
    public const string TooManyImages = "too-many-images";

    /// <summary>An output would create more repeating rows than the effective limit.</summary>
    public const string TooManyRows = "too-many-rows";

    /// <summary>The batch carries more image bytes across all outputs than the effective limit.</summary>
    public const string TooManyTotalImageBytes = "too-many-total-image-bytes";

    /// <summary>The batch would create more repeating rows across all outputs than the effective limit.</summary>
    public const string TooManyTotalRows = "too-many-total-rows";

    /// <summary>A value names a slot the template does not have.</summary>
    public const string UnknownTemplateValue = "unknown-template-value";

    /// <summary>The binding needs a template feature this engine does not support for that format.</summary>
    public const string UnsupportedTemplateFeature = "unsupported-template-feature";

    /// <summary>A bound value is longer than the effective limit.</summary>
    public const string ValueTooLong = "value-too-long";

    /// <summary>A typed value is bound to a slot that cannot hold it.</summary>
    public const string WrongSlotKind = "wrong-slot-kind";
}

/// <summary>
/// Diagnostics from Word document comparison. A refusal names the area it applies to; findings in
/// covered areas are still reported.
/// </summary>
public static class ComparisonDiagnosticCodes
{
    /// <summary>The original has no free body paragraph that can anchor an inserted redline.</summary>
    public const string ComparisonAnchorUnavailable = "comparison-anchor-unavailable";

    /// <summary>The comparison reached its difference limit.</summary>
    public const string ComparisonDifferenceLimitExceeded = "comparison-difference-limit-exceeded";

    /// <summary>An input exceeds the comparison's size limit.</summary>
    public const string ComparisonInputLimitExceeded = "comparison-input-limit-exceeded";

    /// <summary>A document has more body paragraphs than the comparison limit.</summary>
    public const string ComparisonParagraphLimitExceeded = "comparison-paragraph-limit-exceeded";

    /// <summary>Comparison requires two Word .docx packages.</summary>
    public const string UnsupportedComparisonFormat = "unsupported-comparison-format";

    /// <summary>An input still has pending tracked revisions, which must be resolved first.</summary>
    public const string UnsupportedExistingRevisions = "unsupported-existing-revisions";

    /// <summary>Headers or footers changed, which is outside comparison coverage.</summary>
    public const string UnsupportedHeaderFooterChange = "unsupported-header-footer-change";

    /// <summary>Images changed, which is outside comparison coverage.</summary>
    public const string UnsupportedImageChange = "unsupported-image-change";

    /// <summary>Nodes outside free body text changed, such as properties, fields, comments or sections.</summary>
    public const string UnsupportedNodeChange = "unsupported-node-change";

    /// <summary>Content outside the free body changed.</summary>
    public const string UnsupportedNonBodyChange = "unsupported-non-body-change";

    /// <summary>Footnotes or endnotes changed, which is outside comparison coverage.</summary>
    public const string UnsupportedNoteChange = "unsupported-note-change";

    /// <summary>Numbering definitions changed, which is outside comparison coverage.</summary>
    public const string UnsupportedNumberingChange = "unsupported-numbering-change";

    /// <summary>Another package part changed, which is outside comparison coverage.</summary>
    public const string UnsupportedPackageChange = "unsupported-package-change";

    /// <summary>
    /// A paragraph's run structure or direct formatting changed in a way the plan cannot reproduce.
    /// </summary>
    public const string UnsupportedParagraphMarkupChange = "unsupported-paragraph-markup-change";

    /// <summary>A paragraph's style changed.</summary>
    public const string UnsupportedStyleChange = "unsupported-style-change";

    /// <summary>Style definitions changed, which is outside comparison coverage.</summary>
    public const string UnsupportedStyleDefinitionChange = "unsupported-style-definition-change";

    /// <summary>Table geometry changed, so cells cannot be aligned safely.</summary>
    public const string UnsupportedTableChange = "unsupported-table-change";

    /// <summary>
    /// A cell's formatting changed together with its text, so the change cannot be reproduced
    /// without losing formatting.
    /// </summary>
    public const string UnsupportedTableMarkupChange = "unsupported-table-markup-change";
}

/// <summary>Diagnostics from Word document assembly.</summary>
public static class AssemblyDiagnosticCodes
{
    /// <summary>A source or the output is not a readable package.</summary>
    public const string InvalidMergePackage = "invalid-merge-package";

    /// <summary>Source themes, font tables or layout-affecting settings disagree.</summary>
    public const string MergeIncompatibleFormatting = "merge-incompatible-formatting";

    /// <summary>The expanded package size or part count exceeds the host limit.</summary>
    public const string MergeResourceLimit = "merge-resource-limit";

    /// <summary>A source references a definition that does not exist.</summary>
    public const string MergeUnresolvedReference = "merge-unresolved-reference";

    /// <summary>An input changed since the assembly was previewed.</summary>
    public const string StaleMergeSource = "stale-merge-source";

    /// <summary>A source contains a content element assembly does not support.</summary>
    public const string UnsupportedMergeContent = "unsupported-merge-content";

    /// <summary>A source contains a content control other than a plain-text control.</summary>
    public const string UnsupportedMergeContentControl = "unsupported-merge-content-control";

    /// <summary>
    /// A source links external content other than http, https or mailto hyperlinks. External
    /// content is never fetched.
    /// </summary>
    public const string UnsupportedMergeExternalContent = "unsupported-merge-external-content";

    /// <summary>A source contains a field other than switch-free PAGE, NUMPAGES or SECTIONPAGES.</summary>
    public const string UnsupportedMergeField = "unsupported-merge-field";

    /// <summary>A source is not a standard .docx package.</summary>
    public const string UnsupportedMergeFormat = "unsupported-merge-format";

    /// <summary>A source contains a part content type assembly does not support.</summary>
    public const string UnsupportedMergePart = "unsupported-merge-part";

    /// <summary>A shared part does not use the standard Word package layout.</summary>
    public const string UnsupportedMergePartLayout = "unsupported-merge-part-layout";

    /// <summary>A source contains an unsupported or duplicate relationship.</summary>
    public const string UnsupportedMergeRelationship = "unsupported-merge-relationship";

    /// <summary>A source uses review, protection or external processing settings.</summary>
    public const string UnsupportedMergeSettings = "unsupported-merge-settings";
}

/// <summary>Failure codes carried by <see cref="RenderResult.FailureCode"/>.</summary>
public static class RenderFailureCodes
{
    /// <summary>The document exceeded the configured input-size limit.</summary>
    public const string InputLimitExceeded = "input-limit-exceeded";

    /// <summary>The render options are invalid.</summary>
    public const string InvalidRenderOptions = "invalid-render-options";

    /// <summary>A renderer process exceeded the configured memory limit.</summary>
    public const string MemoryLimitExceeded = "memory-limit-exceeded";

    /// <summary>Rendering exceeded the configured output-size limit.</summary>
    public const string OutputLimitExceeded = "output-limit-exceeded";

    /// <summary>Rendering exceeded the configured page-count limit.</summary>
    public const string PageLimitExceeded = "page-limit-exceeded";

    /// <summary>Rendering exceeded the configured time limit.</summary>
    public const string RenderTimeout = "render-timeout";

    /// <summary>A renderer process exited unsuccessfully.</summary>
    public const string RendererFailed = "renderer-failed";

    /// <summary>The renderer completed without producing its expected output.</summary>
    public const string RendererOutputMissing = "renderer-output-missing";

    /// <summary>A renderer executable was not found or could not be started.</summary>
    public const string RendererUnavailable = "renderer-unavailable";
}
