using System.Text.Json.Serialization;

namespace OfficeAgent.Abstractions;

/// <summary>
/// Defines a versioned set of document operations to validate and apply as a unit.
/// </summary>
public sealed class DocumentPlan
{
    /// <summary>
    /// Gets the contract version supported by this engine build.
    /// </summary>
    public const string CurrentContractVersion = "0.2";

    /// <summary>
    /// Gets the contract version used by the plan.
    /// </summary>
    public string ContractVersion { get; init; } = CurrentContractVersion;

    /// <summary>
    /// Gets the document format expected by the plan.
    /// </summary>
    public DocumentFormat Format { get; init; } = DocumentFormat.Word;

    /// <summary>
    /// Gets the snapshot the plan was authored against. When set, the engine rejects
    /// the plan with <see cref="ValidationErrorCodes.StaleSnapshot"/> if the live
    /// document has drifted. Leave <see langword="null"/> to opt out of drift detection.
    /// </summary>
    public SnapshotToken? Snapshot { get; init; }

    /// <summary>
    /// Gets the operations to validate and apply in order.
    /// </summary>
    public IReadOnlyList<PlanOperation> Operations { get; init; } = Array.Empty<PlanOperation>();
}

/// <summary>
/// Specifies whether a text edit is written directly or as a tracked revision.
/// </summary>
public enum ChangeMode
{
    /// <summary>The edit is written as a tracked revision where the format supports it.</summary>
    Tracked,

    /// <summary>The edit is applied directly to the document content.</summary>
    Direct
}

/// <summary>
/// Represents the base type for all plan operations. Only operations implemented by
/// a registered module are part of the wire contract; reserved/future verbs are
/// intentionally absent so an agent never sees a verb that always fails.
/// </summary>
[JsonPolymorphic(TypeDiscriminatorPropertyName = "op")]
[JsonDerivedType(typeof(FillOp), "fill")]
[JsonDerivedType(typeof(ChangeTextOp), "changeText")]
[JsonDerivedType(typeof(InsertOp), "insert")]
[JsonDerivedType(typeof(CommentOp), "comment")]
[JsonDerivedType(typeof(FormatOp), "format")]
[JsonDerivedType(typeof(SetPropertyOp), "setProperty")]
[JsonDerivedType(typeof(RevisionOp), "revision")]
[JsonDerivedType(typeof(InsertTableRowsOp), "insertTableRows")]
[JsonDerivedType(typeof(RemoveTableRowsOp), "removeTableRows")]
[JsonDerivedType(typeof(InsertTableColumnsOp), "insertTableColumns")]
[JsonDerivedType(typeof(RemoveTableColumnsOp), "removeTableColumns")]
[JsonDerivedType(typeof(CopyStylesOp), "copyStyles")]
[JsonDerivedType(typeof(ClearStylesOp), "clearStyles")]
[JsonDerivedType(typeof(InsertImageOp), "insertImage")]
[JsonDerivedType(typeof(RemoveImageOp), "removeImage")]
public abstract class PlanOperation
{
    /// <summary>
    /// Gets the anchor targeted by the operation.
    /// </summary>
    public Anchor Target { get; init; } = null!;
}

/// <summary>
/// Populates a structural slot such as a content control or bookmark.
/// </summary>
public sealed class FillOp : PlanOperation
{
    /// <summary>
    /// Gets the value to place in the target slot.
    /// </summary>
    public string Value { get; init; } = string.Empty;
}

/// <summary>
/// Replaces content-verified text.
/// </summary>
public sealed class ChangeTextOp : PlanOperation
{
    /// <summary>
    /// Gets the replacement text.
    /// </summary>
    public string With { get; init; } = string.Empty;

    /// <summary>
    /// Gets how the replacement is represented in the document.
    /// </summary>
    public ChangeMode Mode { get; init; } = ChangeMode.Tracked;
}

/// <summary>
/// Specifies where inserted content is placed relative to the target anchor.
/// </summary>
public enum InsertPosition
{
    /// <summary>Insert before the target.</summary>
    Before,

    /// <summary>Insert after the target.</summary>
    After
}

/// <summary>
/// Inserts a paragraph or table relative to an anchor.
/// </summary>
public sealed class InsertOp : PlanOperation
{
    /// <summary>
    /// Gets where the new content is inserted relative to the target.
    /// </summary>
    public InsertPosition Position { get; init; } = InsertPosition.After;

    /// <summary>
    /// Gets the paragraph text to insert when <see cref="Table"/> is not set.
    /// </summary>
    public string? Text { get; init; }

    /// <summary>
    /// Gets the style id to apply to the inserted paragraph.
    /// </summary>
    public string? StyleId { get; init; }

    /// <summary>
    /// Gets table data to insert instead of a paragraph.
    /// </summary>
    public TableData? Table { get; init; }
}

/// <summary>
/// Contains tabular data for an inserted Word table.
/// </summary>
public sealed class TableData
{
    /// <summary>
    /// Gets the table header labels.
    /// </summary>
    public IReadOnlyList<string> Headers { get; init; } = Array.Empty<string>();

    /// <summary>
    /// Gets the table body rows.
    /// </summary>
    public IReadOnlyList<IReadOnlyList<string>> Rows { get; init; } = Array.Empty<IReadOnlyList<string>>();

    /// <summary>
    /// Gets the optional table style id.
    /// </summary>
    public string? StyleId { get; init; }
}

/// <summary>
/// Specifies the lifecycle action for a comment operation. Only <see cref="Add"/>
/// is currently implemented.
/// </summary>
public enum CommentAction
{
    /// <summary>Add a new comment.</summary>
    Add
}

/// <summary>
/// Performs a review comment action.
/// </summary>
public sealed class CommentOp : PlanOperation
{
    /// <summary>
    /// Gets the comment body.
    /// </summary>
    public string Text { get; init; } = string.Empty;

    /// <summary>
    /// Gets the display author for a new comment.
    /// </summary>
    public string Author { get; init; } = "OfficeAgent";

    /// <summary>
    /// Gets the author initials for a new comment.
    /// </summary>
    public string Initials { get; init; } = "OA";

    /// <summary>
    /// Gets the comment lifecycle action.
    /// </summary>
    public CommentAction Action { get; init; } = CommentAction.Add;
}

/// <summary>
/// Unified formatting verb. Applies any combination of a named style (<see cref="StyleId"/>)
/// and direct character / paragraph / border properties to the target element. Every
/// property is optional; properties left <see langword="null"/> are not changed. The handler
/// dispatches by target type:
/// <list type="bullet">
/// <item><see cref="TextSpanAnchor"/> - paragraph + runs (empty <c>Expect</c> = whole paragraph).</item>
/// <item><see cref="NodeAnchor"/> <c>kind=table</c> - table style and border.</item>
/// <item><see cref="NodeAnchor"/> <c>kind=tableRow</c> with path <c>table#N/row#M</c> - row height + every paragraph and run inside.</item>
/// <item><see cref="NodeAnchor"/> <c>kind=tableCell</c> with path <c>table#N/cell#R/C</c> - cell border + every paragraph and run inside.</item>
/// <item><see cref="NodeAnchor"/> <c>kind=image</c> - resize to <see cref="WidthPx"/> × <see cref="HeightPx"/>.</item>
/// </list>
/// </summary>
public sealed class FormatOp : PlanOperation
{
    /// <summary>A named style to apply (paragraph style for paragraphs/rows/cells; table style for tables).</summary>
    public string? StyleId { get; init; }

    /// <summary>Font family name applied to runs (e.g. "Calibri", "Arial").</summary>
    public string? FontFamily { get; init; }

    /// <summary>Font size in half-points: 24 = 12pt, 22 = 11pt, 20 = 10pt.</summary>
    public int? SizeHalfPoints { get; init; }

    /// <summary>Bold runs.</summary>
    public bool? Bold { get; init; }

    /// <summary>Italic runs.</summary>
    public bool? Italic { get; init; }

    /// <summary>Single underline runs.</summary>
    public bool? Underline { get; init; }

    /// <summary>Highlight color: yellow, green, cyan, magenta, blue, red, darkBlue, darkCyan, darkGreen, darkMagenta, darkRed, darkYellow, darkGray, lightGray, black, white, none.</summary>
    public string? Highlight { get; init; }

    /// <summary>Hex RGB font color, e.g. "FF0000".</summary>
    public string? Color { get; init; }

    /// <summary>Paragraph horizontal alignment: left, center, right, justify (alias both).</summary>
    public string? Alignment { get; init; }

    /// <summary>Left indent in twips (1/20 of a point; 1440 = 1 inch).</summary>
    public int? IndentLeftTwips { get; init; }

    /// <summary>Right indent in twips.</summary>
    public int? IndentRightTwips { get; init; }

    /// <summary>First-line indent in twips.</summary>
    public int? IndentFirstLineTwips { get; init; }

    /// <summary>Spacing before the paragraph in twips.</summary>
    public int? SpacingBeforeTwips { get; init; }

    /// <summary>Spacing after the paragraph in twips.</summary>
    public int? SpacingAfterTwips { get; init; }

    /// <summary>Border style: single, double, dotted, dashed, thick, none.</summary>
    public string? BorderStyle { get; init; }

    /// <summary>Border width in eighths of a point (8 = 1pt).</summary>
    public int? BorderSizeEighths { get; init; }

    /// <summary>Border hex RGB color, e.g. "000000".</summary>
    public string? BorderColor { get; init; }

    /// <summary>Width in pixels at 96 DPI (images and table rows).</summary>
    public int? WidthPx { get; init; }

    /// <summary>Height in pixels at 96 DPI (images and table rows).</summary>
    public int? HeightPx { get; init; }
}

/// <summary>
/// Sets a property on an addressed node.
/// </summary>
public sealed class SetPropertyOp : PlanOperation
{
    /// <summary>
    /// Gets the property selector understood by the target node provider.
    /// </summary>
    public string Name { get; init; } = "value";

    /// <summary>
    /// Gets the property value to write.
    /// </summary>
    public string? Value { get; init; }
}

/// <summary>
/// Specifies whether tracked revisions are accepted or rejected.
/// </summary>
public enum RevisionAction
{
    /// <summary>Accept the addressed revision.</summary>
    Accept,

    /// <summary>Reject the addressed revision.</summary>
    Reject
}

/// <summary>
/// Accepts or rejects tracked revisions.
/// </summary>
public sealed class RevisionOp : PlanOperation
{
    /// <summary>Gets the revision action.</summary>
    public RevisionAction Action { get; init; } = RevisionAction.Accept;
}

/// <summary>
/// Placement of inserted rows or columns relative to <see cref="InsertTableRowsOp.RowIndex"/>
/// (or <see cref="InsertTableColumnsOp.ColumnIndex"/>). <see cref="End"/> appends to the
/// end of the table; <see cref="Start"/> prepends (after the header conceptually);
/// <see cref="Before"/> and <see cref="After"/> position relative to the supplied index
/// (negative indices count from the end).
/// </summary>
public enum TablePosition
{
    /// <summary>Append after the last existing row/column.</summary>
    End,

    /// <summary>Prepend before the first existing row/column.</summary>
    Start,

    /// <summary>Insert before the row/column at the supplied index.</summary>
    Before,

    /// <summary>Insert after the row/column at the supplied index.</summary>
    After
}

/// <summary>
/// Inserts rows into an existing table identified by a table <see cref="NodeAnchor"/>.
/// Each row is a list of cell texts; missing trailing cells are left empty.
/// Use <see cref="Position"/> and <see cref="RowIndex"/> to control placement.
/// </summary>
public sealed class InsertTableRowsOp : PlanOperation
{
    /// <summary>The rows to insert.</summary>
    public IReadOnlyList<IReadOnlyList<string>> Rows { get; init; } =
        Array.Empty<IReadOnlyList<string>>();

    /// <summary>Where to insert the rows. Default is <see cref="TablePosition.End"/> (append).</summary>
    public TablePosition Position { get; init; } = TablePosition.End;

    /// <summary>
    /// Zero-based row index used when <see cref="Position"/> is <see cref="TablePosition.Before"/>
    /// or <see cref="TablePosition.After"/>. Negative values count from the end (-1 = last row).
    /// </summary>
    public int RowIndex { get; init; }
}

/// <summary>
/// Removes rows from an existing table addressed by a table <see cref="NodeAnchor"/>.
/// Indices are zero-based; negative values count from the end (-1 = last row).
/// When <see cref="OnlyIfEmpty"/> is true, only rows whose every cell is whitespace
/// are removed, which is the safe choice when the LLM wants to "clean up" blank rows.
/// </summary>
public sealed class RemoveTableRowsOp : PlanOperation
{
    /// <summary>The row indices to remove. If empty and <see cref="OnlyIfEmpty"/> is true, every empty row is removed.</summary>
    public IReadOnlyList<int> RowIndices { get; init; } = Array.Empty<int>();

    /// <summary>When true, only rows whose cells are all whitespace are actually removed.</summary>
    public bool OnlyIfEmpty { get; init; }
}

/// <summary>
/// Inserts one or more columns into an existing table. Each entry in
/// <see cref="Columns"/> is a column-major list of cell texts (one per row,
/// header first). Shorter columns are padded with empty cells.
/// </summary>
public sealed class InsertTableColumnsOp : PlanOperation
{
    /// <summary>Column-major data; one inner list per new column, one entry per row.</summary>
    public IReadOnlyList<IReadOnlyList<string>> Columns { get; init; } =
        Array.Empty<IReadOnlyList<string>>();

    /// <summary>Where to insert the columns. Default is <see cref="TablePosition.End"/> (rightmost).</summary>
    public TablePosition Position { get; init; } = TablePosition.End;

    /// <summary>Zero-based column index used when <see cref="Position"/> is Before/After. Negative counts from the right.</summary>
    public int ColumnIndex { get; init; }
}

/// <summary>
/// Removes one or more columns from an existing table by zero-based index.
/// Negative indices count from the right (-1 = last column).
/// </summary>
public sealed class RemoveTableColumnsOp : PlanOperation
{
    /// <summary>The column indices to remove.</summary>
    public IReadOnlyList<int> ColumnIndices { get; init; } = Array.Empty<int>();
}

/// <summary>
/// Copies formatting from a source element to a destination element. Both anchors
/// are <see cref="TextSpanAnchor"/>s; when <c>Expect</c> is empty on either side
/// the entire paragraph is used. Direct run properties (bold, italic, colour, etc.)
/// and paragraph-level properties (including the assigned style id) are copied
/// according to <see cref="Scope"/>.
/// </summary>
public sealed class CopyStylesOp : PlanOperation
{
    /// <summary>The source element whose formatting is copied. Use a <see cref="TextSpanAnchor"/>.</summary>
    public Anchor Source { get; init; } = null!;

    /// <summary>Which scope to copy: <c>run</c>, <c>paragraph</c>, or <c>all</c> (default).</summary>
    public string Scope { get; init; } = "all";
}

/// <summary>
/// Removes formatting from the target element. Target is a <see cref="TextSpanAnchor"/>;
/// when <c>Expect</c> is empty the entire paragraph is affected. Use <see cref="Scope"/>
/// to limit clearing to direct run properties, paragraph-level properties, or both.
/// </summary>
public sealed class ClearStylesOp : PlanOperation
{
    /// <summary>Which scope to clear: <c>run</c>, <c>paragraph</c>, or <c>all</c> (default).</summary>
    public string Scope { get; init; } = "all";
}

/// <summary>
/// Inserts an image into the document, anchored to a <see cref="TextSpanAnchor"/>.
/// Provide image bytes one of two ways: inline as <see cref="Base64Bytes"/>, or
/// indirectly by the opaque <see cref="ImageDocumentId"/> previously returned by
/// adding the image to a provider connection (<see cref="ImageConnectionId"/>).
/// Exactly one of the two routes must be set. The image is placed inline in a new
/// paragraph before or after the anchor paragraph per <see cref="Position"/>.
/// </summary>
public sealed class InsertImageOp : PlanOperation
{
    /// <summary>Base64-encoded image bytes. Mutually exclusive with <see cref="ImageDocumentId"/>.</summary>
    public string? Base64Bytes { get; init; }

    /// <summary>
    /// Connection id of the provider the image was added to.
    /// Required when <see cref="ImageDocumentId"/> is set.
    /// </summary>
    public string? ImageConnectionId { get; init; }

    /// <summary>
    /// Opaque, provider-assigned document id for an image previously added to a
    /// provider connection. Mutually exclusive with <see cref="Base64Bytes"/>.
    /// </summary>
    public string? ImageDocumentId { get; init; }

    /// <summary>Image format: <c>png</c> (default), <c>jpeg</c>, <c>gif</c>, <c>bmp</c>, or <c>tiff</c>.</summary>
    public string ImageType { get; init; } = "png";

    /// <summary>Display width in pixels at 96 DPI. Default 200.</summary>
    public int WidthPx { get; init; } = 200;

    /// <summary>Display height in pixels at 96 DPI. Default 200.</summary>
    public int HeightPx { get; init; } = 200;

    /// <summary>Whether to insert before or after the anchor paragraph.</summary>
    public InsertPosition Position { get; init; } = InsertPosition.After;

    /// <summary>Optional alt text describing the image for accessibility.</summary>
    public string? AltText { get; init; }
}

/// <summary>
/// Removes a specific image addressed by a <see cref="NodeAnchor"/> with
/// <c>Kind="image"</c> and <c>Path="image#N"</c>. Image paths are surfaced
/// by <c>inspect_document.nodes</c>. The underlying image resource is released once no
/// other drawing still references it, so removal leaves no orphaned image bytes behind.
/// </summary>
public sealed class RemoveImageOp : PlanOperation
{
}
