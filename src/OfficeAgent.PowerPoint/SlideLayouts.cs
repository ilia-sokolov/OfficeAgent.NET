using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeAgent.PowerPoint;

/// <summary>
/// The layout vocabulary a plan names, and the layouts a newly created deck ships with.
/// </summary>
/// <remarks>
/// A layout is what makes a generated slide look like the template rather than like a
/// hand-placed text box: the slide states its text, and position, size, font, and bullet
/// styling all come from the layout's placeholders. Decks created here get the five
/// layouts that cover ordinary authoring; a deck built from a corporate template is
/// matched against whatever layouts <em>it</em> defines, so its own styling wins.
/// </remarks>
internal static class SlideLayouts
{
    public const string Title = "title";
    public const string TitleAndContent = "titleAndContent";
    public const string SectionHeader = "sectionHeader";
    public const string TitleOnly = "titleOnly";
    public const string Blank = "blank";

    /// <summary>Geometry for a 16:9 slide, matching what PowerPoint's own layouts use.</summary>
    public static readonly Box TitleBox = new(838200, 365125, 10515600, 1325563);
    public static readonly Box BodyBox = new(838200, 1825625, 10515600, 4351338);
    private static readonly Box CenteredTitleBox = new(1524000, 1122363, 9144000, 2387600);
    private static readonly Box SubtitleBox = new(1524000, 3602038, 9144000, 1655762);
    private static readonly Box SectionTitleBox = new(831850, 1709738, 10515600, 2852737);
    private static readonly Box SectionBodyBox = new(831850, 4589463, 10515600, 1500187);

    /// <summary>Every layout a created deck ships with, in the order PowerPoint lists them.</summary>
    public static readonly IReadOnlyList<LayoutDefinition> All = new[]
    {
        new LayoutDefinition(Title, SlideLayoutValues.Title, "Title Slide",
            new PlaceholderSpec(PlaceholderValues.CenteredTitle, null, "Title 1", CenteredTitleBox, true),
            new PlaceholderSpec(PlaceholderValues.SubTitle, 1U, "Subtitle 2", SubtitleBox, false)),

        new LayoutDefinition(TitleAndContent, SlideLayoutValues.Object, "Title and Content",
            new PlaceholderSpec(PlaceholderValues.Title, null, "Title 1", TitleBox, true),
            new PlaceholderSpec(null, 1U, "Content Placeholder 2", BodyBox, false)),

        new LayoutDefinition(SectionHeader, SlideLayoutValues.SectionHeader, "Section Header",
            new PlaceholderSpec(PlaceholderValues.Title, null, "Title 1", SectionTitleBox, true),
            new PlaceholderSpec(PlaceholderValues.Body, 1U, "Text Placeholder 2", SectionBodyBox, false)),

        new LayoutDefinition(TitleOnly, SlideLayoutValues.TitleOnly, "Title Only",
            new PlaceholderSpec(PlaceholderValues.Title, null, "Title 1", TitleBox, true)),

        new LayoutDefinition(Blank, SlideLayoutValues.Blank, "Blank")
    };

    /// <summary>The layout a plan gets when it names none, inferred from what it supplies.</summary>
    public static string DefaultFor(OfficeAgent.Abstractions.SlideData slide) =>
        slide.Body.Count > 0 ? TitleAndContent
        : !string.IsNullOrEmpty(slide.Title) ? TitleOnly
        : Blank;

    /// <summary>
    /// Finds the layout part a named layout should use in <em>this</em> deck. Matching is
    /// by the layout's own <c>type</c> attribute rather than by our part ordering, because
    /// the deck may come from a template whose layouts we did not create. When nothing
    /// matches, the caller falls back rather than failing: a slide in the wrong layout is
    /// recoverable, a refused edit on someone's corporate template is just an obstacle.
    /// </summary>
    public static SlideLayoutPart? Find(SlideMasterPart master, string name)
    {
        var wanted = All.FirstOrDefault(l =>
            string.Equals(l.Name, name, StringComparison.OrdinalIgnoreCase));
        if (wanted is null) return null;

        foreach (var part in master.SlideLayoutParts)
            if (part.SlideLayout?.Type?.Value == wanted.Type)
                return part;

        return null;
    }

    /// <summary>The layout names a plan may use, for the error message when one is wrong.</summary>
    public static string Names => string.Join(", ", All.Select(l => l.Name));

    /// <summary>A rectangle in EMUs.</summary>
    internal readonly struct Box
    {
        public Box(long x, long y, long cx, long cy)
        {
            X = x; Y = y; Cx = cx; Cy = cy;
        }

        public long X { get; }
        public long Y { get; }
        public long Cx { get; }
        public long Cy { get; }
    }

    /// <summary>One placeholder on a layout or master.</summary>
    internal sealed class PlaceholderSpec
    {
        public PlaceholderSpec(
            PlaceholderValues? type, uint? index, string name, Box box, bool majorFont)
        {
            Type = type;
            Index = index;
            Name = name;
            Box = box;
            MajorFont = majorFont;
        }

        /// <summary>Null means the generic content placeholder, which has no type attribute.</summary>
        public PlaceholderValues? Type { get; }

        public uint? Index { get; }
        public string Name { get; }
        public Box Box { get; }

        /// <summary>Whether the text uses the theme's major (heading) font.</summary>
        public bool MajorFont { get; }

        /// <summary>True when this placeholder holds the slide's title.</summary>
        public bool IsTitle =>
            Type == PlaceholderValues.Title || Type == PlaceholderValues.CenteredTitle;
    }

    /// <summary>One layout: its plan-facing name, its PresentationML type, and its placeholders.</summary>
    internal sealed class LayoutDefinition
    {
        private readonly PlaceholderSpec[] _placeholders;

        public LayoutDefinition(
            string name, SlideLayoutValues type, string displayName, params PlaceholderSpec[] placeholders)
        {
            Name = name;
            Type = type;
            DisplayName = displayName;
            _placeholders = placeholders;
        }

        public string Name { get; }
        public SlideLayoutValues Type { get; }
        public string DisplayName { get; }

        public SlideLayout Build()
        {
            uint shapeId = 2U;
            var shapes = _placeholders
                .Select(p => (OpenXmlElement)BuildPlaceholder(shapeId++, p))
                .ToArray();

            return new SlideLayout(
                new CommonSlideData(BuildTree(shapes)) { Name = DisplayName },
                new ColorMapOverride(new A.MasterColorMapping()))
            { Type = Type };
        }
    }

    /// <summary>Builds a placeholder shape for a layout or master.</summary>
    internal static Shape Placeholder(
        uint shapeId, string name, PlaceholderValues? type, uint? index, Box box, bool majorFont)
    {
        var placeholder = new PlaceholderShape();
        if (type is { } t) placeholder.Type = t;
        if (index is { } i) placeholder.Index = i;

        // The font choice is what makes a title read as a title: pointing level 1 at the
        // theme's major face is how PowerPoint's own layouts do it.
        var listStyle = new A.ListStyle(
            new A.Level1ParagraphProperties(
                new A.DefaultRunProperties(
                    new A.LatinFont { Typeface = majorFont ? "+mj-lt" : "+mn-lt" })));

        return new Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = shapeId, Name = name },
                new NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(placeholder)),
            new ShapeProperties(new A.Transform2D(
                new A.Offset { X = box.X, Y = box.Y },
                new A.Extents { Cx = box.Cx, Cy = box.Cy })),
            new TextBody(
                new A.BodyProperties(),
                listStyle,
                new A.Paragraph(new A.EndParagraphRunProperties { Language = "en-US" })));
    }

    private static Shape BuildPlaceholder(uint shapeId, PlaceholderSpec spec) =>
        Placeholder(shapeId, spec.Name, spec.Type, spec.Index, spec.Box, spec.MajorFont);

    /// <summary>
    /// A shape tree. p:spTree needs both p:nvGrpSpPr and p:grpSpPr, in that order, before
    /// any shape; without the second the part does not validate.
    /// </summary>
    internal static ShapeTree BuildTree(params OpenXmlElement[] shapes)
    {
        var tree = new ShapeTree(
            new NonVisualGroupShapeProperties(
                new NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                new NonVisualGroupShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new GroupShapeProperties());

        foreach (var shape in shapes) tree.Append(shape);
        return tree;
    }
}