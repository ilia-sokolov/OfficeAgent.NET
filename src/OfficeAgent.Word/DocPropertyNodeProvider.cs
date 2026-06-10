using OfficeAgent.Abstractions;

namespace OfficeAgent.Word;

/// <summary>
/// Surfaces and resolves OPC core document properties (creator, title, subject, …).
/// Fully deterministic: these are plain metadata, no layout involved.
/// </summary>
internal sealed class DocPropertyNodeProvider : IWordNodeProvider
{
    public string Kind => "docProperty";

    private static readonly string[] CoreNames =
        { "creator", "title", "subject", "keywords", "lastModifiedBy", "revision", "category", "contentStatus" };

    public static bool IsCore(string name) => Array.IndexOf(CoreNames, name) >= 0;

    public IEnumerable<NodeInfo> Enumerate(WordObjectMap map)
    {
        foreach (var name in CoreNames)
        {
            var value = Read(map, name);
            yield return new NodeInfo
            {
                Kind = Kind,
                Path = $"core/{name}",
                Summary = value ?? string.Empty,
                Anchor = new NodeAnchor { Id = $"docprop:{name}", Kind = Kind, Path = $"core/{name}", Expect = value }
            };
        }
    }

    public ResolvedNode? Resolve(NodeAnchor anchor, WordObjectMap map)
    {
        var name = NameOf(anchor.Path);
        return IsCore(name) ? new ResolvedNode { Kind = Kind, Value = Read(map, name) } : null;
    }

    public static string NameOf(string path) =>
        path.StartsWith("core/", StringComparison.Ordinal) ? path.Substring("core/".Length) : path;

    public static string? Read(WordObjectMap map, string name)
    {
        var p = map.Doc.PackageProperties;
        return name switch
        {
            "creator" => p.Creator,
            "title" => p.Title,
            "subject" => p.Subject,
            "keywords" => p.Keywords,
            "lastModifiedBy" => p.LastModifiedBy,
            "revision" => p.Revision,
            "category" => p.Category,
            "contentStatus" => p.ContentStatus,
            _ => null
        };
    }

    public static void Write(WordObjectMap map, string name, string? value)
    {
        var p = map.Doc.PackageProperties;
        switch (name)
        {
            case "creator": p.Creator = value; break;
            case "title": p.Title = value; break;
            case "subject": p.Subject = value; break;
            case "keywords": p.Keywords = value; break;
            case "lastModifiedBy": p.LastModifiedBy = value; break;
            case "revision": p.Revision = value; break;
            case "category": p.Category = value; break;
            case "contentStatus": p.ContentStatus = value; break;
        }
    }
}
