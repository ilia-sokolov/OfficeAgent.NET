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
