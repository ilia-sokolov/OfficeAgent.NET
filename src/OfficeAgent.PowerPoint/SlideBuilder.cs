using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeAgent.Abstractions;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeAgent.PowerPoint;

/// <summary>
/// Builds a slide from plan content and the layout it will use.
/// </summary>
/// <remarks>
/// The slide's shapes name the layout's placeholders and carry no geometry of their own,
/// which is how PowerPoint itself authors a slide: the layout owns position, size, font,
/// and bullet styling, so changing the layout later restyles the slide instead of leaving
/// it pinned to coordinates baked in at creation time.
/// </remarks>
internal static class SlideBuilder
{
    /// <summary>Builds the slide, filling whichever of the layout's placeholders it can.</summary>
    public static Slide Build(SlideData data, SlideLayoutPart layoutPart)
    {
        var shapes = new List<OpenXmlElement>();
        uint shapeId = 2U;

        var placeholders = PlaceholdersOf(layoutPart);

        if (data.Title is { Length: > 0 } title)
        {
            var target = placeholders.FirstOrDefault(IsTitle);
            if (target is not null)
                shapes.Add(Filled(shapeId++, "Title 1", target, new[] { title }));
        }

        if (data.Body.Count > 0)
        {
            // Anything that is not the title is where body text belongs - a content
            // placeholder, a body placeholder, or a subtitle, depending on the layout.
            var target = placeholders.FirstOrDefault(p => !IsTitle(p));
            if (target is not null)
                shapes.Add(Filled(shapeId, "Content Placeholder 2", target, data.Body));
        }

        return new Slide(new CommonSlideData(SlideLayouts.BuildTree(shapes.ToArray())));
    }

    /// <summary>
    /// Attaches speaker notes. The notes body is a placeholder of its own, so the text
    /// lands where PowerPoint's notes pane reads it rather than as a floating shape.
    /// </summary>
    public static void AddNotes(SlidePart slidePart, string notes)
    {
        var notesPart = slidePart.AddNewPart<NotesSlidePart>();
        notesPart.NotesSlide = new NotesSlide(
            new CommonSlideData(SlideLayouts.BuildTree(
                new Shape(
                    new NonVisualShapeProperties(
                        new NonVisualDrawingProperties { Id = 2U, Name = "Notes Placeholder 1" },
                        new NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                        new ApplicationNonVisualDrawingProperties(
                            new PlaceholderShape { Type = PlaceholderValues.Body, Index = 1U })),
                    new ShapeProperties(),
                    Body(new[] { notes })))));
    }

    /// <summary>The placeholder shapes a layout defines, in the order it defines them.</summary>
    private static IReadOnlyList<Shape> PlaceholdersOf(SlideLayoutPart layoutPart) =>
        layoutPart.SlideLayout?.CommonSlideData?.ShapeTree?
            .Elements<Shape>()
            .Where(s => s.NonVisualShapeProperties?
                .ApplicationNonVisualDrawingProperties?.PlaceholderShape is not null)
            .ToList()
        ?? (IReadOnlyList<Shape>)Array.Empty<Shape>();

    private static bool IsTitle(Shape shape)
    {
        var type = shape.NonVisualShapeProperties?
            .ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value;

        // A placeholder with no type attribute is the generic content one, not a title.
        return type == PlaceholderValues.Title || type == PlaceholderValues.CenteredTitle;
    }

    /// <summary>
    /// A slide shape naming the layout placeholder it inherits from, carrying only text.
    /// </summary>
    private static Shape Filled(uint shapeId, string name, Shape layoutPlaceholder, IReadOnlyList<string> lines)
    {
        var source = layoutPlaceholder.NonVisualShapeProperties?
            .ApplicationNonVisualDrawingProperties?.PlaceholderShape;

        var placeholder = new PlaceholderShape();
        if (source?.Type?.Value is { } type) placeholder.Type = type;
        if (source?.Index?.Value is { } index) placeholder.Index = index;

        return new Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = shapeId, Name = name },
                new NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(placeholder)),
            // Deliberately empty: an empty p:spPr is what makes the shape inherit the
            // layout's geometry. Writing a transform here would pin it instead.
            new ShapeProperties(),
            Body(lines));
    }

    /// <summary>One paragraph per line, so a body renders as the bullets the layout styles.</summary>
    private static TextBody Body(IReadOnlyList<string> lines)
    {
        var body = new TextBody(new A.BodyProperties(), new A.ListStyle());

        foreach (var line in lines)
            body.Append(new A.Paragraph(
                new A.Run(
                    new A.RunProperties { Language = "en-US" },
                    new A.Text(line))));

        return body;
    }
}
