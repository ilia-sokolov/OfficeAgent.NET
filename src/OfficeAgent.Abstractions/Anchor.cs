using System.Text.Json.Serialization;

namespace OfficeAgent.Abstractions;

/// <summary>
/// Represents an engine-issued address that can be used as the target of a document operation.
/// </summary>
/// <remarks>
/// Anchors are returned by inspection and search APIs. Callers should preserve anchor values exactly and should not construct ids from document offsets or Open XML paths unless they own a format module that defines those values.
/// Polymorphism is handled by <see cref="AnchorJsonConverter"/>, which both writes the
/// <c>$anchor</c> discriminator and accepts JSON that omits it (inferring the concrete
/// type from property names) so LLM-generated plans remain robust.
/// </remarks>
[JsonConverter(typeof(AnchorJsonConverter))]
public abstract class Anchor
{
    /// <summary>
    /// Gets the opaque identifier assigned to the address within an inspection result.
    /// </summary>
    public string Id { get; init; } = string.Empty;
}

/// <summary>
/// Identifies a purpose-built document slot such as a Word content control or bookmark.
/// </summary>
public sealed class StructuralAnchor : Anchor
{
    /// <summary>
    /// Gets the slot tag or bookmark name.
    /// </summary>
    public string Tag { get; init; } = string.Empty;

    /// <summary>
    /// Gets the structural anchor kind, for example <c>contentControl</c> or <c>bookmark</c>.
    /// </summary>
    public string Kind { get; init; } = "contentControl";
}

/// <summary>
/// Identifies expected text within a paragraph.
/// </summary>
/// <remarks>
/// The expected text and occurrence are verified when a plan is applied so stale documents fail safely instead of editing an unintended span.
/// </remarks>
public sealed class TextSpanAnchor : Anchor
{
    /// <summary>
    /// Gets the stable paragraph identifier that contains the target span.
    /// </summary>
    public string ParaId { get; init; } = string.Empty;

    /// <summary>
    /// Gets the text expected at the target span.
    /// </summary>
    public string Expect { get; init; } = string.Empty;

    /// <summary>
    /// Gets the zero-based occurrence of <see cref="Expect"/> within the paragraph.
    /// </summary>
    public int Occurrence { get; init; }
}

/// <summary>
/// Identifies a style table entry by style id.
/// </summary>
public sealed class StyleAnchor : Anchor
{
    /// <summary>
    /// Gets the style identifier.
    /// </summary>
    public string StyleId { get; init; } = string.Empty;
}

/// <summary>
/// Identifies an Excel cell or range for future spreadsheet modules.
/// </summary>
public sealed class CellAnchor : Anchor
{
    /// <summary>
    /// Gets the worksheet name.
    /// </summary>
    public string Sheet { get; init; } = string.Empty;

    /// <summary>
    /// Gets the A1-style cell or range reference.
    /// </summary>
    public string Ref { get; init; } = string.Empty;
}

/// <summary>
/// Identifies a PowerPoint shape for future presentation modules.
/// </summary>
public sealed class ShapeAnchor : Anchor
{
    /// <summary>
    /// Gets the slide identifier.
    /// </summary>
    public string SlideId { get; init; } = string.Empty;

    /// <summary>
    /// Gets the shape identifier.
    /// </summary>
    public string ShapeId { get; init; } = string.Empty;
}

/// <summary>
/// Identifies a format-specific object exposed through a node provider.
/// </summary>
/// <remarks>
/// Word uses node anchors for objects that are not simple paragraph text spans, such as document properties and tracked revisions. Each format module owns the vocabulary for <see cref="Kind"/>, <see cref="Path"/>, and <see cref="Props"/>.
/// </remarks>
public sealed class NodeAnchor : Anchor
{
    /// <summary>
    /// Gets the node kind understood by the format module.
    /// </summary>
    public string Kind { get; init; } = string.Empty;

    /// <summary>
    /// Gets the provider-defined locator for the node.
    /// </summary>
    public string Path { get; init; } = string.Empty;

    /// <summary>
    /// Gets the optional content fingerprint verified before applying an operation.
    /// </summary>
    public string? Expect { get; init; }

    /// <summary>
    /// Gets the zero-based occurrence used when the provider path is not unique by itself.
    /// </summary>
    public int Occurrence { get; init; }

    /// <summary>
    /// Gets optional provider-defined qualifiers for the target.
    /// </summary>
    public IReadOnlyDictionary<string, string>? Props { get; init; }
}
