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
    /// Gets the document format expected by the plan. Defaults to
    /// <see cref="DocumentFormat.Unspecified"/>, meaning the plan applies to whatever
    /// format the document is; set it only to assert that the document must be a
    /// particular format, and a mismatch then fails the plan.
    /// </summary>
    /// <remarks>
    /// The agent tools document a plan as <c>{ "operations": [ … ] }</c> with no format
    /// field, so defaulting to a concrete format would silently bind every such plan to
    /// it and reject the others.
    /// </remarks>
    public DocumentFormat Format { get; init; } = DocumentFormat.Unspecified;

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
// The verb map lives in the converter rather than in [JsonDerivedType] attributes:
// System.Text.Json will not combine a custom converter with its own polymorphism, and its
// built-in reader requires the discriminator to be the object's first property - which a
// model writing a plan has no reason to do. PlanOperationJsonConverterTests keeps the map
// complete as verbs are added.
[JsonConverter(typeof(PlanOperationJsonConverter))]
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
/// Inserts a new paragraph relative to an anchor. To insert a table, use
/// <see cref="InsertTableOp"/>.
/// </summary>
public sealed class InsertOp : PlanOperation
{
    /// <summary>
    /// Gets where the new content is inserted relative to the target.
    /// </summary>
    public InsertPosition Position { get; init; } = InsertPosition.After;

    /// <summary>
    /// Gets the paragraph text to insert.
    /// </summary>
    public string? Text { get; init; }

    /// <summary>
    /// Gets the style id to apply to the inserted paragraph. Word only - a deck has no
    /// paragraph style table.
    /// </summary>
    public string? StyleId { get; init; }

    /// <summary>
    /// Gets the outline level of the inserted paragraph, zero-based - the bullet depth on
    /// a slide. PresentationML only: the Word module refuses it rather than dropping it,
    /// because a list level silently lost renders as a document that looks wrong.
    /// </summary>
    public int? Level { get; init; }
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
/// Specifies the lifecycle action for a comment operation. Which actions a module
/// implements depends on the module, not on the format: both formats can record a
/// resolved comment - PresentationML in the modern comment's status, WordprocessingML in
/// the <c>commentsExtended</c> part's <c>w15:done</c> - but only the PowerPoint module
/// implements <see cref="Resolve"/> today. The Word module reports it as
/// <c>unsupported-operation</c>.
/// </summary>
public enum CommentAction
{
    /// <summary>Add a new comment.</summary>
    Add,

    /// <summary>
    /// Mark an existing comment resolved, addressed by a comment
    /// <see cref="NodeAnchor"/>. The comment and its replies are kept; only its status
    /// changes, so the review history survives.
    /// </summary>
    Resolve
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

    /// <summary>Width in pixels at 96 DPI (images, table rows, and any slide shape).</summary>
    public int? WidthPx { get; init; }

    /// <summary>Height in pixels at 96 DPI (images, table rows, and any slide shape).</summary>
    public int? HeightPx { get; init; }

    /// <summary>
    /// Distance in pixels at 96 DPI from the slide's left edge. PresentationML only:
    /// a slide positions shapes absolutely, whereas a Word document lays them out in flow.
    /// </summary>
    public int? XPx { get; init; }

    /// <summary>Distance in pixels at 96 DPI from the slide's top edge. PresentationML only.</summary>
    public int? YPx { get; init; }
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
/// Inserts a new table relative to an anchored paragraph (a <see cref="TextSpanAnchor"/>).
/// A first-class table verb so an agent can create a table directly; the generic
/// <see cref="InsertOp"/> inserts paragraphs only.
/// </summary>
public sealed class InsertTableOp : PlanOperation
{
    /// <summary>Gets where the new table is inserted relative to the target paragraph.</summary>
    public InsertPosition Position { get; init; } = InsertPosition.After;

    /// <summary>Gets the table content to insert.</summary>
    public TableData Table { get; init; } = new();
}

/// <summary>
/// Removes an entire table addressed by a table <see cref="NodeAnchor"/> with
/// <c>Kind="table"</c> and <c>Path="table#N"</c>. The table and all of its rows are
/// deleted; to drop only some rows or columns, use the row/column verbs instead.
/// </summary>
public sealed class RemoveTableOp : PlanOperation
{
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

/// <summary>
/// Specifies where a slide lands relative to the deck or to a reference slide.
/// </summary>
public enum SlidePosition
{
    /// <summary>Before every existing slide.</summary>
    Start,

    /// <summary>After every existing slide. The default for a new slide.</summary>
    End,

    /// <summary>Immediately before the reference slide.</summary>
    Before,

    /// <summary>Immediately after the reference slide.</summary>
    After
}

/// <summary>
/// Describes one slide to author. The layout supplies the geometry and styling; this
/// carries only the content that goes into it, so a generated slide looks like one the
/// deck's own template would produce rather than a hand-placed text box.
/// </summary>
public sealed class SlideData
{
    /// <summary>
    /// Gets the layout the slide uses: <c>title</c>, <c>titleAndContent</c> (the default
    /// when a body is supplied), <c>sectionHeader</c>, <c>titleOnly</c>, or <c>blank</c>.
    /// Resolved against the layouts the deck's own slide master defines, so a deck built
    /// from a corporate template keeps that template's appearance.
    /// </summary>
    public string? Layout { get; init; }

    /// <summary>Gets the title placeholder's text. Omitted leaves the placeholder empty.</summary>
    public string? Title { get; init; }

    /// <summary>Gets the body placeholder's paragraphs - one entry per bullet.</summary>
    public IReadOnlyList<string> Body { get; init; } = Array.Empty<string>();

    /// <summary>Gets the speaker notes. Omitted creates no notes slide.</summary>
    public string? Notes { get; init; }
}

/// <summary>
/// Adds a slide to a presentation. <see cref="SlidePosition.Start"/> and
/// <see cref="SlidePosition.End"/> need no target; <see cref="SlidePosition.Before"/> and
/// <see cref="SlidePosition.After"/> take a slide <see cref="NodeAnchor"/> as the
/// reference point.
/// </summary>
public sealed class InsertSlideOp : PlanOperation
{
    /// <summary>Gets where the slide lands. The default appends to the deck.</summary>
    public SlidePosition Position { get; init; } = SlidePosition.End;

    /// <summary>Gets the slide's layout and content.</summary>
    public SlideData Slide { get; init; } = new();
}

/// <summary>
/// Removes the slide addressed by a <see cref="NodeAnchor"/> with <c>Kind="slide"</c>.
/// The slide, its notes, and its relationships go with it; the layout and master it used
/// stay, because other slides share them.
/// </summary>
public sealed class RemoveSlideOp : PlanOperation
{
}

/// <summary>
/// Reorders the deck by moving the slide addressed by a slide <see cref="NodeAnchor"/>.
/// Slide ids are durable across reordering, so anchors into any slide keep working.
/// </summary>
public sealed class MoveSlideOp : PlanOperation
{
    /// <summary>Gets where the slide lands.</summary>
    public SlidePosition Position { get; init; } = SlidePosition.End;

    /// <summary>
    /// Gets the reference slide's path (for example <c>slide#257</c>), required for
    /// <see cref="SlidePosition.Before"/> and <see cref="SlidePosition.After"/>.
    /// </summary>
    public string? RelativeTo { get; init; }
}

/// <summary>
/// Copies the slide addressed by a slide <see cref="NodeAnchor"/>, content and all. The
/// copy receives its own slide id and its own shape ids, and by default lands immediately
/// after the original - what duplicating a slide in PowerPoint does.
/// </summary>
public sealed class DuplicateSlideOp : PlanOperation
{
    /// <summary>Gets where the copy lands. The default places it after the original.</summary>
    public SlidePosition Position { get; init; } = SlidePosition.After;

    /// <summary>
    /// Gets the reference slide's path for <see cref="SlidePosition.Before"/> and
    /// <see cref="SlidePosition.After"/>. Omitted, the original is the reference.
    /// </summary>
    public string? RelativeTo { get; init; }
}

/// <summary>
/// Adds a free-standing text box to a slide. Unlike a placeholder, it belongs to the
/// slide rather than to the layout, so it carries its own position and size.
/// </summary>
/// <remarks>
/// The target is a slide <see cref="NodeAnchor"/>. Text that belongs in the slide's title
/// or body should go through the layout's placeholders instead - a text box laid over a
/// placeholder looks right until the template changes underneath it.
/// </remarks>
public sealed class InsertShapeOp : PlanOperation
{
    /// <summary>Gets the paragraphs the box contains - one entry per line.</summary>
    public IReadOnlyList<string> Text { get; init; } = Array.Empty<string>();

    /// <summary>Gets the distance in pixels at 96 DPI from the slide's left edge.</summary>
    public int? XPx { get; init; }

    /// <summary>Gets the distance in pixels at 96 DPI from the slide's top edge.</summary>
    public int? YPx { get; init; }

    /// <summary>Gets the width in pixels at 96 DPI. Default 400.</summary>
    public int WidthPx { get; init; } = 400;

    /// <summary>Gets the height in pixels at 96 DPI. Default 100.</summary>
    public int HeightPx { get; init; } = 100;
}

/// <summary>
/// Removes a shape addressed by a <see cref="NodeAnchor"/> with <c>Kind="shape"</c> and
/// <c>Path="shape#{slideId}/{shapeId}"</c>. Works for any shape a slide holds - a text
/// box, a table frame, or a picture - so it is the general counterpart to the
/// type-specific <see cref="RemoveTableOp"/> and <see cref="RemoveImageOp"/>.
/// </summary>
/// <remarks>
/// Removing a placeholder is refused: the layout would immediately re-offer it as an empty
/// prompt, so the slide would look unchanged while its content was gone. Clear the text
/// instead.
/// </remarks>
public sealed class RemoveShapeOp : PlanOperation
{
}

/// <summary>Specifies the lifecycle action for a <see cref="SectionOp"/>.</summary>
public enum SectionAction
{
    /// <summary>Start a new section at the target slide.</summary>
    Add,

    /// <summary>Rename the target section.</summary>
    Rename,

    /// <summary>
    /// Remove the target section. Its slides are kept and join the section before it, or
    /// become unsectioned when it was the first - deleting a grouping must not delete the
    /// things being grouped.
    /// </summary>
    Remove
}

/// <summary>
/// Manages the named slide groups PowerPoint shows in the thumbnail pane.
/// </summary>
/// <remarks>
/// Sections partition the deck in presentation order: a section owns a contiguous run of
/// slides, and once a deck has any section every slide belongs to one. The module keeps
/// that invariant as slides are added, moved, copied and removed, so the grouping cannot
/// drift out of step with the deck.
/// <para>
/// <see cref="SectionAction.Add"/> targets the slide the section starts at;
/// <see cref="SectionAction.Rename"/> and <see cref="SectionAction.Remove"/> target a
/// section <see cref="NodeAnchor"/> with <c>Kind="section"</c>.
/// </para>
/// </remarks>
public sealed class SectionOp : PlanOperation
{
    /// <summary>Gets the action to perform.</summary>
    public SectionAction Action { get; init; } = SectionAction.Add;

    /// <summary>Gets the section name, for <see cref="SectionAction.Add"/> and <see cref="SectionAction.Rename"/>.</summary>
    public string Name { get; init; } = string.Empty;
}

/// <summary>
/// Configures the running items along a slide's edge: the footer text, the slide number,
/// and the date. Targeting a slide node configures that slide; omitting the target
/// configures every slide, which is what PowerPoint's "Apply to All" does.
/// </summary>
/// <remarks>
/// A slide has no header. PresentationML carries a header flag on <c>p:hf</c>, but it
/// applies to notes and handout pages only - PowerPoint's own dialog greys it out on the
/// Slide tab - so the module refuses one rather than writing a field nothing renders.
/// </remarks>
public sealed class HeaderFooterOp : PlanOperation
{
    /// <summary>
    /// Gets the footer text. An empty string clears it. Leaving this
    /// <see langword="null"/> keeps whatever the slide already shows.
    /// </summary>
    public string? Footer { get; init; }

    /// <summary>Gets whether the footer is displayed.</summary>
    public bool? ShowFooter { get; init; }

    /// <summary>Gets whether the slide number is displayed.</summary>
    public bool? ShowSlideNumber { get; init; }

    /// <summary>Gets whether the date is displayed.</summary>
    public bool? ShowDateTime { get; init; }

    /// <summary>
    /// Gets fixed date text. When <see langword="null"/> and the date is shown, the slide
    /// carries a field PowerPoint refreshes on open instead - the "update automatically"
    /// option, which is what a deck presented more than once wants.
    /// </summary>
    public string? DateTime { get; init; }
}

/// <summary>Distinguishes the two kinds of timeline media a slide can carry.</summary>
public enum MediaKind
{
    /// <summary>A movie, shown in a frame on the slide.</summary>
    Video,

    /// <summary>A sound, shown as a speaker icon.</summary>
    Audio
}

/// <summary>
/// Embeds video or audio in a slide. The bytes travel inside the package, so the deck
/// still plays when it is mailed on - a linked file would not.
/// </summary>
/// <remarks>
/// Supply the media inline as <see cref="Base64Bytes"/>, or indirectly by the opaque
/// <see cref="MediaDocumentId"/> of a document already registered with a provider
/// connection. Exactly one of the two routes must be set, matching how
/// <see cref="InsertImageOp"/> takes an image.
/// </remarks>
public sealed class InsertMediaOp : PlanOperation
{
    /// <summary>Gets whether this is video or audio.</summary>
    public MediaKind Kind { get; init; } = MediaKind.Video;

    /// <summary>Gets the media bytes, base64-encoded.</summary>
    public string? Base64Bytes { get; init; }

    /// <summary>Gets the connection holding the media document, used with <see cref="MediaDocumentId"/>.</summary>
    public string? MediaConnectionId { get; init; }

    /// <summary>Gets the opaque id of a registered document holding the media bytes.</summary>
    public string? MediaDocumentId { get; init; }

    /// <summary>
    /// Gets the file extension the bytes are in - <c>mp4</c>, <c>m4a</c>, <c>mp3</c>,
    /// <c>wav</c>. It selects the media type the package declares, which is how PowerPoint
    /// decides whether it can play the stream at all.
    /// </summary>
    public string MediaType { get; init; } = "mp4";

    /// <summary>
    /// Gets the poster image shown before playback, base64-encoded PNG. Omitted, the frame
    /// is a plain placeholder; PowerPoint does not generate one from the media itself.
    /// </summary>
    public string? PosterBase64 { get; init; }

    /// <summary>Gets the distance in pixels at 96 DPI from the slide's left edge.</summary>
    public int? XPx { get; init; }

    /// <summary>Gets the distance in pixels at 96 DPI from the slide's top edge.</summary>
    public int? YPx { get; init; }

    /// <summary>Gets the frame width in pixels at 96 DPI. Default 480.</summary>
    public int WidthPx { get; init; } = 480;

    /// <summary>Gets the frame height in pixels at 96 DPI. Default 270.</summary>
    public int HeightPx { get; init; } = 270;

    /// <summary>Gets optional alt text describing the media for accessibility.</summary>
    public string? AltText { get; init; }
}
