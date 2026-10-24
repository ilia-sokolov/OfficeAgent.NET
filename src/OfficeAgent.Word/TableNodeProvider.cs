using DocumentFormat.OpenXml;
using OfficeAgent.Abstractions;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.Word;

/// <summary>
/// Surfaces and resolves Word tables across every text host. Each table is addressed
/// by its zero-based document order, e.g. <c>table#0</c>, <c>table#1</c>.
/// </summary>
internal sealed class TableNodeProvider : IWordNodeProvider
{
    public string Kind => "table";

    public IEnumerable<NodeInfo> Enumerate(WordObjectMap map)
    {
        int index = 0;
        foreach (var (root, _) in WordModel.TextHosts(map.Package))
        {
            foreach (var table in root.Descendants<Table>())
            {
                var rows = table.Elements<TableRow>().ToList();
                var headerCells = rows.FirstOrDefault()
                    ?.Descendants<TableCell>()
                    .Select(c => string.Concat(c.Descendants<Text>().Select(t => t.Text)))
                    .ToList() ?? new List<string>();
                var summary = $"Table {index}: {rows.Count} row(s) x {headerCells.Count} column(s). Headers: {string.Join(" | ", headerCells)}";

                yield return new NodeInfo
                {
                    Kind = Kind,
                    Path = $"table#{index}",
                    Summary = summary,
                    Name = CaptionOf(table),
                    Anchor = new NodeAnchor
                    {
                        Id = $"tbl:{index}",
                        Kind = Kind,
                        Path = $"table#{index}"
                    }
                };
                index++;
            }
        }
    }

    /// <summary>The table's alt-text title (<c>w:tblCaption</c>), which is where a data binding marker lives.</summary>
    internal static string? CaptionOf(Table table) =>
        table.GetFirstChild<TableProperties>()?.GetFirstChild<TableCaption>()?.Val?.Value is { Length: > 0 } caption
            ? caption
            : null;

    public ResolvedNode? Resolve(NodeAnchor anchor, WordObjectMap map)
    {
        if (!anchor.Path.StartsWith("table#", StringComparison.Ordinal)) return null;
        if (!int.TryParse(anchor.Path.Substring("table#".Length), out var target)) return null;

        int index = 0;
        foreach (var (root, _) in WordModel.TextHosts(map.Package))
        {
            foreach (var table in root.Descendants<Table>())
            {
                if (index == target)
                    return new ResolvedNode
                    {
                        Kind = Kind,
                        Elements = new OpenXmlElement[] { table }
                    };
                index++;
            }
        }
        return null;
    }
}
