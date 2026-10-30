namespace OfficeAgent.Abstractions;

/// <summary>
/// Identifies a specific inspected state of a document.
/// </summary>
public sealed class SnapshotToken
{
    /// <summary>Gets the snapshot entity tag.</summary>
    public string ETag { get; init; } = string.Empty;

    /// <summary>Initializes a new instance of the <see cref="SnapshotToken"/> class.</summary>
    public SnapshotToken()
    {
    }

    /// <summary>Initializes a new instance of the <see cref="SnapshotToken"/> class.</summary>
    /// <param name="etag">The snapshot entity tag.</param>
    public SnapshotToken(string etag) => ETag = etag;
}

/// <summary>
/// Contains the structured, anchored model returned by inspection.
/// </summary>
public sealed class InspectResult
{
    /// <summary>Gets the document format.</summary>
    public DocumentFormat Format { get; init; }

    /// <summary>Gets the snapshot token for drift detection.</summary>
    public SnapshotToken Snapshot { get; init; } = new();

    /// <summary>Gets heading-derived outline nodes.</summary>
    public IReadOnlyList<OutlineNode> Outline { get; init; } = Array.Empty<OutlineNode>();

    /// <summary>Gets the document style catalog.</summary>
    public StyleCatalog Styles { get; init; } = new();

    /// <summary>Gets all anchors surfaced by the inspection pass.</summary>
    public IReadOnlyList<Anchor> Anchors { get; init; } = Array.Empty<Anchor>();

    /// <summary>Gets paragraph content discovered during inspection.</summary>
    public IReadOnlyList<ParagraphInfo> Paragraphs { get; init; } = Array.Empty<ParagraphInfo>();

    /// <summary>Gets structural anchors such as content controls and bookmarks.</summary>
    public IReadOnlyList<StructuralAnchor> StructuralAnchors { get; init; } = Array.Empty<StructuralAnchor>();

    /// <summary>Gets format-specific nodes surfaced by node providers.</summary>
    public IReadOnlyList<NodeInfo> Nodes { get; init; } = Array.Empty<NodeInfo>();

    /// <summary>Gets worksheets surfaced by Excel inspection.</summary>
    public IReadOnlyList<WorksheetInfo> Worksheets { get; init; } = Array.Empty<WorksheetInfo>();

    /// <summary>Gets populated cells from the bounded Excel inspection range.</summary>
    public IReadOnlyList<CellInfo> Cells { get; init; } = Array.Empty<CellInfo>();

    /// <summary>
    /// Gets a workbook's defined names that refer to one range, such as <c>Revenue</c> for
    /// <c>Summary!$B$2</c>. Excel only; hidden and built-in names are left out.
    /// </summary>
    public IReadOnlyList<DefinedNameInfo> DefinedNames { get; init; } = Array.Empty<DefinedNameInfo>();

    /// <summary>
    /// Gets the slide layouts a deck defines, by the names PowerPoint shows, with the
    /// placeholders each one offers. Any of these names can be given as an inserted slide's
    /// <see cref="SlideData.Layout"/>. PowerPoint only.
    /// </summary>
    public IReadOnlyList<SlideLayoutInfo> Layouts { get; init; } = Array.Empty<SlideLayoutInfo>();
}

/// <summary>One slide layout a deck defines.</summary>
public sealed class SlideLayoutInfo
{
    /// <summary>Gets the layout's name as PowerPoint shows it, such as <c>Two Content</c>.</summary>
    public string Name { get; init; } = string.Empty;

    /// <summary>
    /// Gets the name of the slide master the layout belongs to. When several masters define
    /// a layout of the same name, <c>master/name</c> picks one.
    /// </summary>
    public string Master { get; init; } = string.Empty;

    /// <summary>
    /// Gets the placeholders a slide on this layout offers, in layout order. Date, footer
    /// and slide-number placeholders are left out.
    /// </summary>
    public IReadOnlyList<LayoutPlaceholderInfo> Placeholders { get; init; } = Array.Empty<LayoutPlaceholderInfo>();
}

/// <summary>One placeholder on a slide layout.</summary>
public sealed class LayoutPlaceholderInfo
{
    /// <summary>Gets the placeholder's shape name, which an inserted slide's shape keeps.</summary>
    public string Name { get; init; } = string.Empty;

    /// <summary>
    /// Gets what the placeholder holds: <c>title</c>, <c>subtitle</c>, <c>body</c>,
    /// <c>content</c>, <c>picture</c>, <c>chart</c>, <c>table</c>, <c>diagram</c> or <c>media</c>.
    /// </summary>
    public string Kind { get; init; } = string.Empty;
}

/// <summary>A workbook defined name and the range it refers to.</summary>
public sealed class DefinedNameInfo
{
    /// <summary>Gets the name as Excel shows it.</summary>
    public string Name { get; init; } = string.Empty;

    /// <summary>Gets the durable sheet id of the range.</summary>
    public uint SheetId { get; init; }

    /// <summary>Gets the A1 range, without the sheet name or <c>$</c> signs, such as <c>B2</c> or <c>A1:D6</c>.</summary>
    public string Range { get; init; } = string.Empty;
}

/// <summary>Describes one worksheet and its table catalog.</summary>
public sealed class WorksheetInfo
{
    /// <summary>Gets the durable workbook sheet id.</summary>
    public uint SheetId { get; init; }
    /// <summary>Gets the worksheet display name.</summary>
    public string Name { get; init; } = string.Empty;
    /// <summary>Gets the used A1 dimension when present.</summary>
    public string? Dimension { get; init; }
    /// <summary>Gets tables defined on the sheet.</summary>
    public IReadOnlyList<SpreadsheetTableInfo> Tables { get; init; } = Array.Empty<SpreadsheetTableInfo>();
}

/// <summary>Describes an Excel table.</summary>
public sealed class SpreadsheetTableInfo
{
    /// <summary>Gets the table name.</summary>
    public string Name { get; init; } = string.Empty;
    /// <summary>Gets the table display name.</summary>
    public string DisplayName { get; init; } = string.Empty;
    /// <summary>Gets the A1 range occupied by the table.</summary>
    public string Reference { get; init; } = string.Empty;
}

/// <summary>Describes one populated Excel cell without evaluating formulas.</summary>
public sealed class CellInfo
{
    /// <summary>Gets the cell anchor.</summary>
    public CellAnchor Anchor { get; init; } = new();
    /// <summary>Gets the cell's stored XML value.</summary>
    public string? RawValue { get; init; }
    /// <summary>Gets text resolved from inline or shared strings and cached formula values.</summary>
    public string? DisplayValue { get; init; }
    /// <summary>Gets the formula text when the cell contains a formula.</summary>
    public string? Formula { get; init; }

    /// <summary>
    /// Gets the value as Excel displays it, with the cell's number format applied - <c>$4,200,000</c>
    /// or <c>12.5%</c> rather than <c>4200000</c> or <c>0.125</c>. Numbers are written with the
    /// invariant culture's separators. A formula's cached result is formatted; a formula with no
    /// cached result has no formatted value.
    /// </summary>
    public string? FormattedValue { get; init; }
}

/// <summary>
/// Describes an addressable format-specific object.
/// </summary>
public sealed class NodeInfo
{
    /// <summary>Gets the provider-defined node kind.</summary>
    public string Kind { get; init; } = string.Empty;

    /// <summary>Gets the provider-defined node path.</summary>
    public string Path { get; init; } = string.Empty;

    /// <summary>Gets a display summary for the node.</summary>
    public string Summary { get; init; } = string.Empty;

    /// <summary>Gets the anchor for the node.</summary>
    public NodeAnchor? Anchor { get; init; }

    /// <summary>
    /// Gets the name an author gave the node, or <see langword="null"/> when it has none: a
    /// PowerPoint shape, table or chart's Selection Pane name, or a Word table's alt-text title.
    /// A name starting with <see cref="DataBindingMarker.Prefix"/> links the node to a workbook range.
    /// </summary>
    public string? Name { get; init; }
}

/// <summary>
/// Represents one node in the document outline.
/// </summary>
public sealed class OutlineNode
{
    /// <summary>Gets the outline heading text.</summary>
    public string Text { get; init; } = string.Empty;

    /// <summary>Gets the heading level.</summary>
    public int Level { get; init; }

    /// <summary>Gets the anchor associated with the outline node.</summary>
    public Anchor? Anchor { get; init; }

    /// <summary>Gets child outline nodes.</summary>
    public IReadOnlyList<OutlineNode> Children { get; init; } = Array.Empty<OutlineNode>();
}

/// <summary>
/// Describes one paragraph in the inspected content model.
/// </summary>
public sealed class ParagraphInfo
{
    /// <summary>Gets the stable paragraph identifier.</summary>
    public string ParaId { get; init; } = string.Empty;

    /// <summary>Gets the logical paragraph text.</summary>
    public string Text { get; init; } = string.Empty;

    /// <summary>Gets the paragraph style id, when one is assigned.</summary>
    public string? StyleId { get; init; }

    /// <summary>
    /// Gets the containing node path when the paragraph is inside a structural
    /// container (e.g. <c>table#0</c> when the paragraph lives in a table cell),
    /// or <see langword="null"/> when it is a free-flowing body paragraph.
    /// </summary>
    public string? In { get; init; }

    /// <summary>
    /// Gets where the paragraph lives: <c>body</c>, <c>header</c>, <c>footer</c>,
    /// <c>footnote</c>, or <c>endnote</c>. Body content — including tables — belongs in
    /// <c>body</c>; anchoring body edits to a header/footnote/endnote paragraph places the
    /// change outside the document flow.
    /// </summary>
    public string Location { get; init; } = "body";

    /// <summary>
    /// Gets the paragraph text with its character formatting marked as numbered tags, such as
    /// <c>Payment is due within &lt;1&gt;30 days&lt;/1&gt;.</c>, or <see langword="null"/> when
    /// the whole paragraph has one formatting. Rewrite it keeping the tags around the words
    /// that should carry each formatting, and send it as <see cref="ChangeTextOp.WithMarked"/>.
    /// </summary>
    /// <remarks>
    /// Untagged text has the formatting most of the paragraph has. Also <see langword="null"/>
    /// for a Word paragraph whose text crosses a hyperlink, content control or tracked change,
    /// which a whole-paragraph rewrite cannot keep.
    /// </remarks>
    public string? MarkedText { get; init; }
}

/// <summary>
/// Contains the styles discovered in the document.
/// </summary>
public sealed class StyleCatalog
{
    /// <summary>Gets style metadata entries.</summary>
    public IReadOnlyList<StyleInfo> Styles { get; init; } = Array.Empty<StyleInfo>();
}

/// <summary>
/// Describes one style in the document style catalog.
/// </summary>
public sealed class StyleInfo
{
    /// <summary>Gets the style id.</summary>
    public string Id { get; init; } = string.Empty;

    /// <summary>Gets the display style name.</summary>
    public string Name { get; init; } = string.Empty;

    /// <summary>Gets the base style id, when the style inherits from another style.</summary>
    public string? BasedOn { get; init; }

    /// <summary>Gets a value indicating whether the style is custom.</summary>
    public bool IsCustom { get; init; }

    /// <summary>Gets how many inspected paragraphs use the style.</summary>
    public int InUseCount { get; init; }
}
