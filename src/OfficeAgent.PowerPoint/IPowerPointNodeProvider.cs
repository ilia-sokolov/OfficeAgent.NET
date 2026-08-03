using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;

namespace OfficeAgent.PowerPoint;

/// <summary>
/// Surfaces and resolves one kind of PowerPoint object behind a uniform seam. Adding a
/// presentation primitive = implement a provider (enumerate for inspect + content-verified
/// resolve for apply) and contribute it to <see cref="PowerPointModule"/> directly or via
/// dependency injection - no new anchor class and, where an existing verb fits, no new verb.
/// </summary>
/// <summary>
/// Marks an operation handler as belonging to the PowerPoint module.
/// </summary>
/// <remarks>
/// Handlers are contributed through dependency injection, and a bare
/// <see cref="IOperationHandler"/> registration cannot say which format it understands:
/// every module resolving <c>IEnumerable&lt;IOperationHandler&gt;</c> would receive it,
/// so a Word-specific handler for a verb PowerPoint does not implement would be applied
/// to decks. Contributing through this interface keeps a handler with its own format.
/// </remarks>
public interface IPowerPointOperationHandler : IOperationHandler
{
}

public interface IPowerPointNodeProvider
{
    /// <summary>The node kind this provider owns, as it appears in <c>inspect.nodes</c>.</summary>
    string Kind { get; }

    /// <summary>Lists the provider's nodes for inspection.</summary>
    IEnumerable<NodeInfo> Enumerate(PowerPointObjectMap map);

    /// <summary>Re-locates a node from its anchor at apply time, or returns null when it is gone.</summary>
    ResolvedNode? Resolve(NodeAnchor anchor, PowerPointObjectMap map);
}

/// <summary>A lightweight view over an open presentation for providers and handlers.</summary>
public sealed class PowerPointObjectMap
{
    /// <summary>The open package.</summary>
    public IOpenXmlPackage Package { get; }

    /// <summary>The typed presentation document.</summary>
    public PresentationDocument Doc => (PresentationDocument)Package.Package;

    /// <summary>The presentation part.</summary>
    public PresentationPart Main => Doc.PresentationPart
        ?? throw new InvalidOperationException("Presentation has no presentation part.");

    /// <summary>Initializes the map over an open package.</summary>
    public PowerPointObjectMap(IOpenXmlPackage package) => Package = package;
}

/// <summary>A node re-located from its anchor: the live element(s) plus a current value.</summary>
/// <remarks>
/// Deliberately separate from the Word module's namesake. The two formats agree on the
/// shape of the answer but not on what a node is, and coupling them would make either
/// module's node vocabulary a breaking change for the other.
/// </remarks>
public sealed class ResolvedNode
{
    /// <summary>The node kind that resolved.</summary>
    public string Kind { get; init; } = string.Empty;

    /// <summary>The live elements the anchor pointed at.</summary>
    public IReadOnlyList<OpenXmlElement> Elements { get; init; } = Array.Empty<OpenXmlElement>();

    /// <summary>The node's current value, when it has one.</summary>
    public string? Value { get; init; }
}
