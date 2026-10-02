using System.Diagnostics.CodeAnalysis;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;

namespace OfficeAgent.Word;

/// <summary>
/// Surfaces and resolves one kind of Word object behind a uniform seam. Adding a
/// Word primitive = implement a provider (enumerate for inspect + content-verified
/// resolve for apply) and contribute it to <see cref="WordModule"/> directly or via
/// dependency injection - no new anchor class and, where an existing verb fits, no
/// new verb.
/// </summary>
[Experimental(EngineExtensibility.DiagnosticId, UrlFormat = EngineExtensibility.UrlFormat)]
public interface IWordNodeProvider
{
    /// <summary>Gets the node kind this provider owns, as it appears in <c>inspect.nodes</c>.</summary>
    string Kind { get; }

    /// <summary>Lists the provider's nodes for inspection.</summary>
    IEnumerable<NodeInfo> Enumerate(WordObjectMap map);

    /// <summary>Re-locates a node from its anchor at apply time, or returns null when it is gone.</summary>
    ResolvedNode? Resolve(NodeAnchor anchor, WordObjectMap map);
}

/// <summary>A lightweight view over an open Word package for providers and handlers.</summary>
[Experimental(EngineExtensibility.DiagnosticId, UrlFormat = EngineExtensibility.UrlFormat)]
public sealed class WordObjectMap
{
    /// <summary>Gets the open package the map views.</summary>
    public IOpenXmlPackage Package { get; }

    /// <summary>Gets the package as a typed Word document.</summary>
    public WordprocessingDocument Doc => (WordprocessingDocument)Package.Package;

    /// <summary>Gets the main document part; throws when the package has none.</summary>
    public MainDocumentPart Main => Doc.MainDocumentPart
        ?? throw new InvalidOperationException("Word document has no main part.");

    /// <summary>Initializes a map over an open Word package.</summary>
    public WordObjectMap(IOpenXmlPackage package) => Package = package;
}

/// <summary>A node re-located from its anchor: the live element(s) plus a current value.</summary>
[Experimental(EngineExtensibility.DiagnosticId, UrlFormat = EngineExtensibility.UrlFormat)]
public sealed class ResolvedNode
{
    /// <summary>Gets the node kind.</summary>
    public string Kind { get; init; } = string.Empty;

    /// <summary>Gets the live elements that make up the node.</summary>
    public IReadOnlyList<OpenXmlElement> Elements { get; init; } = Array.Empty<OpenXmlElement>();

    /// <summary>Gets the node's current value, when it has one.</summary>
    public string? Value { get; init; }
}
