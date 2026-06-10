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
public interface IWordNodeProvider
{
    string Kind { get; }


    IEnumerable<NodeInfo> Enumerate(WordObjectMap map);


    ResolvedNode? Resolve(NodeAnchor anchor, WordObjectMap map);
}

/// <summary>A lightweight view over an open Word package for providers and handlers.</summary>
public sealed class WordObjectMap
{
    public IOpenXmlPackage Package { get; }

    public WordprocessingDocument Doc => (WordprocessingDocument)Package.Package;

    public MainDocumentPart Main => Doc.MainDocumentPart
        ?? throw new InvalidOperationException("Word document has no main part.");

    public WordObjectMap(IOpenXmlPackage package) => Package = package;
}

/// <summary>A node re-located from its anchor: the live element(s) plus a current value.</summary>
public sealed class ResolvedNode
{
    public string Kind { get; init; } = string.Empty;
    public IReadOnlyList<OpenXmlElement> Elements { get; init; } = Array.Empty<OpenXmlElement>();
    public string? Value { get; init; }
}
