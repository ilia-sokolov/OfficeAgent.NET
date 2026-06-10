using DocumentFormat.OpenXml;
using OfficeAgent.Core;
using OfficeAgent.Abstractions;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.Word;

/// <summary>
/// Surfaces and resolves tracked revisions (<c>w:ins</c> / <c>w:del</c>) across every
/// text host (body, headers, footers, footnotes, endnotes). Accepting/rejecting them
/// is deterministic XML surgery.
/// </summary>
internal sealed class RevisionNodeProvider : IWordNodeProvider
{
    public string Kind => "revision";

    public IEnumerable<NodeInfo> Enumerate(WordObjectMap map)
    {
        foreach (var (root, _) in WordModel.TextHosts(map.Package))
        {
            foreach (var ins in root.Descendants<InsertedRun>())
                yield return Node("ins", ins.Id?.Value, string.Concat(ins.Descendants<Text>().Select(t => t.Text)));

            foreach (var del in root.Descendants<DeletedRun>())
                yield return Node("del", del.Id?.Value, string.Concat(del.Descendants<DeletedText>().Select(t => t.Text)));
        }
    }

    public ResolvedNode? Resolve(NodeAnchor anchor, WordObjectMap map)
    {
        var inserted = new List<InsertedRun>();
        var deleted = new List<DeletedRun>();
        foreach (var (root, _) in WordModel.TextHosts(map.Package))
        {
            inserted.AddRange(root.Descendants<InsertedRun>());
            deleted.AddRange(root.Descendants<DeletedRun>());
        }

        IEnumerable<OpenXmlElement> elements = anchor.Path switch
        {
            "all" => inserted.Cast<OpenXmlElement>().Concat(deleted),
            var p when p.StartsWith("ins#", StringComparison.Ordinal) =>
                inserted.Where(e => e.Id?.Value == p.Substring(4)),
            var p when p.StartsWith("del#", StringComparison.Ordinal) =>
                deleted.Where(e => e.Id?.Value == p.Substring(4)),
            _ => Enumerable.Empty<OpenXmlElement>()
        };

        var list = elements.ToList();
        if (list.Count == 0 && anchor.Path != "all") return null;
        return new ResolvedNode { Kind = Kind, Elements = list };
    }

    private NodeInfo Node(string tag, string? id, string text) => new()
    {
        Kind = Kind,
        Path = $"{tag}#{id}",
        Summary = $"{tag}: {text}",
        Anchor = new NodeAnchor { Id = $"rev:{tag}:{id}", Kind = Kind, Path = $"{tag}#{id}", Expect = text }
    };
}
