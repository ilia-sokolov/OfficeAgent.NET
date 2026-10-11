using System.Globalization;
using DocumentFormat.OpenXml;
using OfficeAgent.Abstractions;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.Word;

/// <summary>
/// Surfaces the document's numbering instances (<c>w:num</c>) as <c>list</c> nodes, so a
/// caller can continue a list the document already defines with <c>format</c>'s
/// <c>numberingId</c>. The path is <c>list#{numId}</c>; the summary describes the levels by
/// what they look like - number format, the text Word prints, and the paragraph style a
/// level names - which is how a caller tells a contract's clause numbering from a bullet list.
/// </summary>
internal sealed class ListNodeProvider : IWordNodeProvider
{
    private const int LevelsShown = 3;

    public string Kind => "list";

    public IEnumerable<NodeInfo> Enumerate(WordObjectMap map)
    {
        var numbering = map.Main.NumberingDefinitionsPart?.Numbering;
        if (numbering is null) yield break;

        foreach (var instance in numbering.Elements<NumberingInstance>())
        {
            if (instance.NumberID?.Value is not { } id) continue;
            var path = $"list#{id.ToString(CultureInfo.InvariantCulture)}";
            yield return new NodeInfo
            {
                Kind = Kind,
                Path = path,
                Summary = Describe(map, instance, id),
                Anchor = new NodeAnchor { Id = $"list:{id}", Kind = Kind, Path = path }
            };
        }
    }

    public ResolvedNode? Resolve(NodeAnchor anchor, WordObjectMap map)
    {
        if (!anchor.Path.StartsWith("list#", StringComparison.Ordinal) ||
            !int.TryParse(anchor.Path.Substring("list#".Length), NumberStyles.None, CultureInfo.InvariantCulture, out var id))
            return null;

        var instance = map.Main.NumberingDefinitionsPart?.Numbering?.Elements<NumberingInstance>()
            .FirstOrDefault(n => n.NumberID?.Value == id);
        return instance is null
            ? null
            : new ResolvedNode { Kind = Kind, Elements = new OpenXmlElement[] { instance }, Value = Describe(map, instance, id) };
    }

    private static string Describe(WordObjectMap map, NumberingInstance instance, int id)
    {
        var definition = WordNumbering.AbstractFor(map.Main, instance);
        var name = definition?.AbstractNumDefinitionName?.Val?.Value
                   ?? definition?.StyleLink?.Val?.Value
                   ?? definition?.NumberingStyleLink?.Val?.Value;
        var levels = new List<string>();
        for (var level = 0; level <= WordNumbering.MaxLevel && levels.Count < LevelsShown; level++)
        {
            var definitionLevel = WordNumbering.LevelFor(map.Main, instance, level);
            if (definitionLevel is null) continue;
            var format = definitionLevel.NumberingFormat?.Val?.InnerText ?? "decimal";
            var text = definitionLevel.LevelText?.Val?.Value ?? string.Empty;
            var style = definitionLevel.ParagraphStyleIdInLevel?.Val?.Value;
            levels.Add($"level {level} {format} \"{text}\"{(style is null ? string.Empty : $" style {style}")}");
        }

        var head = name is null ? $"List {id}" : $"List {id} ({name})";
        return levels.Count == 0 ? head : $"{head}: {string.Join("; ", levels)}";
    }

}
