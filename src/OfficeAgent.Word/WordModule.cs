using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using DocumentFormat.OpenXml.Wordprocessing;
using DocFormat = OfficeAgent.Abstractions.DocumentFormat;

namespace OfficeAgent.Word;

/// <summary>
/// Provides Word inspection, search, and supported plan operation handling over
/// WordprocessingML across the body, headers, footers, footnotes, and endnotes.
/// </summary>
public sealed class WordModule : IFormatModule
{
    public DocFormat Format => DocFormat.Word;

    private readonly IReadOnlyList<IWordNodeProvider> _providers;

    public IReadOnlyList<IOperationHandler> Handlers { get; }

    public WordModule() : this(TimeProvider.System) { }

    /// <summary>
    /// Initializes the module with a clock plus optional externally contributed handlers
    /// and node providers (e.g. registered through dependency injection). Built-in
    /// handlers take precedence for the verbs they support; contributed handlers extend
    /// the module to new operations.
    /// </summary>
    public WordModule(
        TimeProvider clock,
        IEnumerable<IOperationHandler>? extraHandlers = null,
        IEnumerable<IWordNodeProvider>? extraProviders = null)
    {
        Handlers = new IOperationHandler[]
        {
            new ChangeTextHandler(clock),
            new FillHandler(),
            new CommentHandler(clock),
            new InsertHandler(),
            new FormatHandler(),
            new SetPropertyHandler(),
            new RevisionHandler(),
            new InsertTableRowsHandler(),
            new RemoveTableRowsHandler(),
            new InsertTableColumnsHandler(),
            new RemoveTableColumnsHandler(),
            new CopyStylesHandler(),
            new ClearStylesHandler(),
            new InsertImageHandler(),
            new RemoveImageHandler()
        }
        .Concat(extraHandlers ?? Enumerable.Empty<IOperationHandler>())
        .ToList();

        _providers = new IWordNodeProvider[]
        {
            new DocPropertyNodeProvider(),
            new RevisionNodeProvider(),
            new TableNodeProvider(),
            new ImageNodeProvider()
        }
        .Concat(extraProviders ?? Enumerable.Empty<IWordNodeProvider>())
        .ToList();
    }

    public bool CanHandle(IOpenXmlPackage package) => package.Format == DocFormat.Word;

    public IReadOnlyDictionary<string, string> Stabilize(IOpenXmlPackage package) =>
        WordModel.Stabilize(package);

    public InspectResult Inspect(IOpenXmlPackage package, InspectOptions options)
    {
        var styleNames = StyleNameMap(package);

        var paragraphs = new List<ParagraphInfo>();
        var anchors = new List<Anchor>();
        var headings = new List<(int Level, string Text, Anchor Anchor)>();

        var wantContent = options.Fidelity == Fidelity.Content;
        var wantStructure = options.Fidelity != Fidelity.Outline;

        // Map each <w:tbl> element to its table#N path so paragraphs can advertise containment.
        var tablePaths = BuildTablePathIndex(package);

        foreach (var (paragraph, paraId, _) in WordModel.Paragraphs(package))
        {
            var text = WordModel.Text.GetLogicalText(paragraph);
            var styleId = WordModel.StyleOf(paragraph);
            string? inPath = null;
            var ancestorTable = paragraph.Ancestors<Table>().FirstOrDefault();
            if (ancestorTable is not null && tablePaths.TryGetValue(ancestorTable, out var path))
                inPath = path;

            if (wantContent)
                paragraphs.Add(new ParagraphInfo { ParaId = paraId, Text = text, StyleId = styleId, In = inPath });

            var anchor = new TextSpanAnchor { Id = paraId, ParaId = paraId, Expect = text, Occurrence = 0 };
            if (wantContent) anchors.Add(anchor);

            if (HeadingLevel(styleId, styleId is null ? null : TryGet(styleNames, styleId)) is int level)
                headings.Add((level, text, anchor));
        }

        var structural = wantStructure ? StructuralAnchors(package).ToList() : new List<StructuralAnchor>();
        if (wantContent) anchors.AddRange(structural);

        var nodes = new List<NodeInfo>();
        if (wantContent)
        {
            var map = new WordObjectMap(package);
            foreach (var provider in _providers)
                foreach (var node in provider.Enumerate(map))
                {
                    nodes.Add(node);
                    if (node.Anchor is not null) anchors.Add(node.Anchor);
                }
        }

        return new InspectResult
        {
            Format = DocFormat.Word,
            Snapshot = WordModel.Snapshot(package),
            Outline = Nest(headings),
            Styles = wantStructure ? BuildCatalog(package, paragraphs) : new StyleCatalog(),
            Anchors = anchors,
            Paragraphs = paragraphs,
            StructuralAnchors = structural,
            Nodes = nodes
        };
    }

    public IReadOnlyList<FindHit> Find(IOpenXmlPackage package, FindQuery query)
    {
        var hits = new List<FindHit>();
        var comparison = WordModel.Comparison(query.Options.CaseSensitive);

        Regex? regex = BuildRegex(query);

        foreach (var (paragraph, paraId, _) in WordModel.Paragraphs(package))
        {
            var text = WordModel.Text.GetLogicalText(paragraph);
            if (text.Length == 0) continue;

            var seen = new Dictionary<string, int>(StringComparer.Ordinal);

            if (regex is not null)
            {
                foreach (Match m in regex.Matches(text))
                {
                    if (m.Length == 0) continue;
                    AddHit(hits, seen, paraId, text, m.Value, m.Index);
                }
            }
            else
            {
                int from = 0;
                while (true)
                {
                    int idx = text.IndexOf(query.Pattern, from, comparison);
                    if (idx < 0) break;
                    AddHit(hits, seen, paraId, text, text.Substring(idx, query.Pattern.Length), idx);
                    from = idx + query.Pattern.Length;
                }
            }
        }

        return hits;
    }

    private static void AddHit(List<FindHit> hits, Dictionary<string, int> seen, string paraId, string paragraphText, string matched, int index)
    {
        int occurrence = seen.TryGetValue(matched, out var n) ? n : 0;
        seen[matched] = occurrence + 1;

        hits.Add(new FindHit
        {
            Anchor = new TextSpanAnchor { Id = paraId, ParaId = paraId, Expect = matched, Occurrence = occurrence },
            Text = matched,
            Context = WordModel.Snippet(paragraphText, index, matched.Length)
        });
    }

    private static Regex? BuildRegex(FindQuery query)
    {
        var opts = query.Options;
        if (!opts.Regex && !opts.WholeWord)
            return null;

        var pattern = opts.Regex ? query.Pattern : Regex.Escape(query.Pattern);
        if (opts.WholeWord)
            pattern = $@"\b(?:{pattern})\b";

        var regexOptions = RegexOptions.CultureInvariant;
        if (!opts.CaseSensitive)
            regexOptions |= RegexOptions.IgnoreCase;

        return new Regex(pattern, regexOptions);
    }

    private static IEnumerable<StructuralAnchor> StructuralAnchors(IOpenXmlPackage package)
    {
        foreach (var (root, _) in WordModel.TextHosts(package))
        {
            foreach (var sdt in root.Descendants<SdtElement>())
            {
                var tag = sdt.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value;
                if (!string.IsNullOrEmpty(tag))
                    yield return new StructuralAnchor { Id = $"cc:{tag}", Tag = tag, Kind = "contentControl" };
            }

            foreach (var bookmark in root.Descendants<BookmarkStart>())
            {
                var name = bookmark.Name?.Value;
                if (!string.IsNullOrEmpty(name) && name != "_GoBack")
                    yield return new StructuralAnchor { Id = $"bm:{name}", Tag = name, Kind = "bookmark" };
            }
        }
    }

    private static Dictionary<string, string> StyleNameMap(IOpenXmlPackage package)
    {
        var styles = WordModel.Doc(package).MainDocumentPart?.StyleDefinitionsPart?.Styles;
        var map = new Dictionary<string, string>(StringComparer.Ordinal);
        if (styles is null) return map;

        foreach (var style in styles.Elements<Style>())
        {
            var id = style.StyleId?.Value;
            var name = style.StyleName?.Val?.Value;
            if (id is not null && name is not null)
                map[id] = name;
        }
        return map;
    }

    private static StyleCatalog BuildCatalog(IOpenXmlPackage package, IReadOnlyList<ParagraphInfo> paragraphs)
    {
        var styles = WordModel.Doc(package).MainDocumentPart?.StyleDefinitionsPart?.Styles;
        if (styles is null)
            return new StyleCatalog();

        var usage = paragraphs
            .Where(p => p.StyleId is not null)
            .GroupBy(p => p.StyleId!)
            .ToDictionary(g => g.Key, g => g.Count(), StringComparer.Ordinal);

        var infos = new List<StyleInfo>();
        foreach (var style in styles.Elements<Style>())
        {
            var id = style.StyleId?.Value;
            if (id is null) continue;

            infos.Add(new StyleInfo
            {
                Id = id,
                Name = style.StyleName?.Val?.Value ?? id,
                BasedOn = style.BasedOn?.Val?.Value,
                IsCustom = style.CustomStyle?.Value ?? false,
                InUseCount = usage.TryGetValue(id, out var count) ? count : 0
            });
        }

        return new StyleCatalog { Styles = infos };
    }

    private static string? TryGet(Dictionary<string, string> map, string key) =>
        map.TryGetValue(key, out var value) ? value : null;

    private static Dictionary<Table, string> BuildTablePathIndex(IOpenXmlPackage package)
    {
        var map = new Dictionary<Table, string>();
        int index = 0;
        foreach (var (root, _) in WordModel.TextHosts(package))
            foreach (var table in root.Descendants<Table>())
                map[table] = $"table#{index++}";
        return map;
    }

    private static int? HeadingLevel(string? styleId, string? styleName)
    {
        if (styleId is { Length: > 0 } &&
            styleId.StartsWith("Heading", StringComparison.OrdinalIgnoreCase) &&
            int.TryParse(styleId.Substring("Heading".Length), out var byId))
            return byId;

        if (styleName is { Length: > 0 } &&
            styleName.StartsWith("heading ", StringComparison.OrdinalIgnoreCase) &&
            int.TryParse(styleName.Substring("heading ".Length), out var byName))
            return byName;

        return null;
    }

    private static IReadOnlyList<OutlineNode> Nest(List<(int Level, string Text, Anchor Anchor)> headings)
    {
        var roots = new List<MutableOutline>();
        var stack = new Stack<MutableOutline>();

        foreach (var (level, text, anchor) in headings)
        {
            var node = new MutableOutline { Level = level, Text = text, Anchor = anchor };

            while (stack.Count > 0 && stack.Peek().Level >= level)
                stack.Pop();

            if (stack.Count == 0)
                roots.Add(node);
            else
                stack.Peek().Children.Add(node);

            stack.Push(node);
        }

        return roots.Select(r => r.ToNode()).ToList();
    }

    private sealed class MutableOutline
    {
        public int Level;
        public string Text = string.Empty;
        public Anchor? Anchor;
        public List<MutableOutline> Children = new();

        public OutlineNode ToNode() => new()
        {
            Level = Level,
            Text = Text,
            Anchor = Anchor,
            Children = Children.Select(c => c.ToNode()).ToList()
        };
    }
}
