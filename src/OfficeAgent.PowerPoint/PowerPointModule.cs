using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Presentation;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using A = DocumentFormat.OpenXml.Drawing;
using DocFormat = OfficeAgent.Abstractions.DocumentFormat;

namespace OfficeAgent.PowerPoint;

/// <summary>
/// Provides PowerPoint inspection, search, and supported plan operation handling over
/// PresentationML across slides, their tables, and their notes.
/// </summary>
public sealed class PowerPointModule : IFormatModule, IBlankDocumentFactory
{
    /// <inheritdoc />
    public DocFormat Format => DocFormat.PowerPoint;

    /// <inheritdoc />
    public string Extension => ".pptx";

    /// <summary>
    /// Returns a minimal valid .pptx: one slide carrying an empty title placeholder,
    /// addressable as <c>slide256/shape2/p0</c> by an initial plan, over the slide
    /// master, layout, and theme a presentation needs in order to open at all.
    /// </summary>
    public byte[] CreateBlank() => PowerPointBlankDocument.Create();

    /// <inheritdoc />
    public IReadOnlyList<IOperationHandler> Handlers { get; }

    private readonly IReadOnlyList<IPowerPointNodeProvider> _providers;

    /// <summary>Initializes the module with its built-in handlers.</summary>
    public PowerPointModule() : this(TimeProvider.System) { }

    /// <summary>
    /// Initializes the module with a clock plus optional externally contributed handlers
    /// and node providers. Built-in handlers take precedence for the verbs they support;
    /// contributed handlers extend the module to new operations.
    /// </summary>
    public PowerPointModule(
        TimeProvider clock,
        IEnumerable<IOperationHandler>? extraHandlers = null,
        IEnumerable<IPowerPointNodeProvider>? extraProviders = null)
    {
        Handlers = new IOperationHandler[]
        {
            new SlideChangeTextHandler(),
            new SlideInsertTableHandler(),
            new SlideRemoveTableHandler(),
            new SlideInsertTableRowsHandler(),
            new SlideRemoveTableRowsHandler(),
            new SlideInsertTableColumnsHandler(),
            new SlideRemoveTableColumnsHandler(),
            new SlideInsertImageHandler(),
            new SlideRemoveImageHandler(),
            new SlideCommentHandler(clock),
            new SlideFormatHandler(),
            new SlideInsertHandler(),
            new SlideRemoveHandler(),
            new SlideMoveHandler(),
            new SlideDuplicateHandler()
        }
        .Concat(extraHandlers ?? Enumerable.Empty<IOperationHandler>())
        .ToList();

        _providers = new IPowerPointNodeProvider[]
        {
            new SlideNodeProvider(),
            new SlideTableNodeProvider(),
            new SlideImageNodeProvider(),
            new SlideCommentNodeProvider()
        }
        .Concat(extraProviders ?? Enumerable.Empty<IPowerPointNodeProvider>())
        .ToList();
    }

    /// <inheritdoc />
    public bool CanHandle(IOpenXmlPackage package) => package.Format == DocFormat.PowerPoint;

    /// <summary>
    /// No-op: DrawingML has no per-paragraph identifier to mint, so there is nothing to
    /// stabilise and no aliases to translate.
    /// </summary>
    /// <remarks>
    /// Word assigns <c>w14:paraId</c> here so that an operation which shifts paragraph
    /// offsets cannot redirect a later operation's target. PowerPoint paragraph ids are
    /// scoped to a shape and positional only within it, and none of this module's
    /// operations add or remove paragraphs inside a body, so offsets cannot shift
    /// mid-plan. Should a paragraph-inserting verb ever be supported here, it will need
    /// a durable id scheme rather than an empty alias map.
    /// </remarks>
    public IReadOnlyDictionary<string, string> Stabilize(IOpenXmlPackage package) =>
        new Dictionary<string, string>(0, StringComparer.Ordinal);

    /// <inheritdoc />
    public InspectResult Inspect(IOpenXmlPackage package, InspectOptions options)
    {
        var wantContent = options.Fidelity == Fidelity.Content;

        var paragraphs = new List<ParagraphInfo>();
        var anchors = new List<Anchor>();
        var outline = new List<OutlineNode>();

        foreach (var slide in PowerPointModel.Slides(package))
        {
            // Each slide is an outline entry titled by its title placeholder, so the
            // agent can navigate a deck the way it navigates a Word document's headings.
            var title = TitleOf(slide);
            var slideAnchor = new ShapeAnchor
            {
                Id = $"slide{slide.SlideId}",
                SlideId = slide.SlideId.ToString(),
                ShapeId = string.Empty
            };
            outline.Add(new OutlineNode
            {
                Level = 1,
                Text = title.Length > 0 ? title : $"Slide {slide.Number}",
                Anchor = slideAnchor
            });
            if (wantContent) anchors.Add(slideAnchor);

            if (!wantContent) continue;

            foreach (var paragraph in PowerPointModel.Paragraphs(slide))
            {
                var text = PowerPointModel.TextOf(paragraph.Paragraph);
                paragraphs.Add(new ParagraphInfo
                {
                    ParaId = paragraph.ParaId,
                    Text = text,
                    StyleId = null,
                    In = paragraph.Host.Key,
                    Location = paragraph.Location
                });
                anchors.Add(new TextSpanAnchor
                {
                    Id = paragraph.ParaId,
                    ParaId = paragraph.ParaId,
                    Expect = text,
                    Occurrence = 0
                });
            }
        }

        var nodes = new List<NodeInfo>();
        if (wantContent)
        {
            var map = new PowerPointObjectMap(package);
            foreach (var provider in _providers)
                foreach (var node in provider.Enumerate(map))
                {
                    nodes.Add(node);
                    if (node.Anchor is not null) anchors.Add(node.Anchor);
                }
        }

        return new InspectResult
        {
            Format = DocFormat.PowerPoint,
            Snapshot = PowerPointModel.Snapshot(package),
            Outline = outline,
            // A deck has no style table an agent can edit the way Word's is edited;
            // slide layouts govern appearance and are not addressable as styles.
            Styles = new StyleCatalog(),
            Anchors = anchors,
            Paragraphs = paragraphs,
            // PresentationML has no counterpart to Word's content controls or bookmarks:
            // a placeholder is a layout role, not an addressable slot a plan can fill.
            StructuralAnchors = Array.Empty<StructuralAnchor>(),
            Nodes = nodes
        };
    }

    /// <inheritdoc />
    public IReadOnlyList<FindHit> Find(IOpenXmlPackage package, FindQuery query)
    {
        var hits = new List<FindHit>();
        var comparison = PowerPointModel.Comparison(query.Options.CaseSensitive);
        var regex = BuildRegex(query);

        foreach (var paragraph in PowerPointModel.Paragraphs(package))
        {
            var text = PowerPointModel.TextOf(paragraph.Paragraph);
            if (text.Length == 0) continue;

            var seen = new Dictionary<string, int>(StringComparer.Ordinal);

            if (regex is not null)
            {
                foreach (Match match in regex.Matches(text))
                {
                    if (match.Length == 0) continue;
                    AddHit(hits, seen, paragraph.ParaId, text, match.Value, match.Index);
                }
            }
            else
            {
                var from = 0;
                while (true)
                {
                    var index = text.IndexOf(query.Pattern, from, comparison);
                    if (index < 0) break;
                    AddHit(hits, seen, paragraph.ParaId, text, text.Substring(index, query.Pattern.Length), index);
                    from = index + query.Pattern.Length;
                }
            }
        }

        return hits;
    }

    /// <summary>The text of the slide's title placeholder, for the outline entry.</summary>
    private static string TitleOf(SlideRef slide)
    {
        foreach (var shape in slide.Part.Slide.Descendants<Shape>())
        {
            var placeholder = shape.NonVisualShapeProperties?
                .ApplicationNonVisualDrawingProperties?.PlaceholderShape;
            if (placeholder is null) continue;

            var type = placeholder.Type?.Value;
            var isTitle = type is null
                || type == PlaceholderValues.Title
                || type == PlaceholderValues.CenteredTitle;
            if (!isTitle || shape.TextBody is null) continue;

            var text = string.Join(" ", shape.TextBody.Elements<A.Paragraph>()
                .Select(PowerPointModel.TextOf)
                .Where(t => t.Length > 0));
            if (text.Length > 0) return text;
        }
        return string.Empty;
    }

    private static void AddHit(
        List<FindHit> hits,
        Dictionary<string, int> seen,
        string paraId,
        string paragraphText,
        string matched,
        int index)
    {
        var occurrence = seen.TryGetValue(matched, out var n) ? n : 0;
        seen[matched] = occurrence + 1;

        hits.Add(new FindHit
        {
            Anchor = new TextSpanAnchor { Id = paraId, ParaId = paraId, Expect = matched, Occurrence = occurrence },
            Text = matched,
            Context = PowerPointModel.Snippet(paragraphText, index, matched.Length)
        });
    }

    private static Regex? BuildRegex(FindQuery query)
    {
        var options = query.Options;
        if (!options.Regex && !options.WholeWord) return null;

        var pattern = options.Regex ? query.Pattern : Regex.Escape(query.Pattern);
        if (options.WholeWord) pattern = $@"\b(?:{pattern})\b";

        var regexOptions = RegexOptions.CultureInvariant;
        if (!options.CaseSensitive) regexOptions |= RegexOptions.IgnoreCase;

        return new Regex(pattern, regexOptions);
    }
}
