using OfficeAgent.Abstractions;
using OfficeAgent.Core;

namespace OfficeAgent.PowerPoint;

/// <summary>
/// The deck's core properties (title, creator, subject, …) as <c>docProperty</c> nodes with
/// <c>core/*</c> paths, the same paths a Word document answers.
/// </summary>
internal sealed class CorePropertyNodeProvider : IPowerPointNodeProvider
{
    public string Kind => PackageCoreProperties.Kind;

    public IEnumerable<NodeInfo> Enumerate(PowerPointObjectMap map) => PackageCoreProperties.Nodes(map.Doc);

    public ResolvedNode? Resolve(NodeAnchor anchor, PowerPointObjectMap map)
    {
        var name = PackageCoreProperties.NameOf(anchor.Path);
        return PackageCoreProperties.IsCore(name)
            ? new ResolvedNode { Kind = Kind, Value = PackageCoreProperties.Read(map.Doc, name) }
            : null;
    }
}
