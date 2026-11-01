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
    /// <summary>
    /// Builds the slide, filling whichever of the layout's placeholders it can. With
    /// <paramref name="everyPlaceholder"/>, the layout's other placeholders are added empty
    /// and keep the layout's own shape names, so they can be found and filled afterwards.
    /// </summary>
    public static Slide Build(SlideData data, SlideLayoutPart layoutPart, bool everyPlaceholder = false)
    {
        var placeholders = PlaceholdersOf(layoutPart);

        var titleTarget = data.Title is { Length: > 0 } ? placeholders.FirstOrDefault(IsTitle) : null;
        // Body text belongs in a placeholder that holds text - content, body, or a subtitle,
        // depending on the layout - never a picture, chart or table frame.
        var bodyTarget = data.Body.Count > 0 ? placeholders.FirstOrDefault(TakesBody) : null;

        var shapes = new List<OpenXmlElement>();
        uint shapeId = 2U;

        if (everyPlaceholder)
        {
            foreach (var placeholder in placeholders)
            {
                var kind = KindOf(placeholder);
                if (kind is null) continue;

                var name = placeholder.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value
                    ?? $"Placeholder {shapeId}";
                var lines = placeholder == titleTarget ? new[] { data.Title! }
                    : placeholder == bodyTarget ? data.Body
                    : Array.Empty<string>();
                var holdsText = kind is "title" or "subtitle" or "body" or "content";
                shapes.Add(Filled(shapeId++, name, placeholder, lines, holdsText));
            }

            return new Slide(new CommonSlideData(SlideLayouts.BuildTree(shapes.ToArray())));
        }

        if (titleTarget is not null)
            shapes.Add(Filled(shapeId++, "Title 1", titleTarget, new[] { data.Title! }));

        if (bodyTarget is not null)
            shapes.Add(Filled(shapeId, "Content Placeholder 2", bodyTarget, data.Body));

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

    /// <summary>
    /// Why the slide's own text would not fit the layout, or null when it fits. A title or
    /// body with no placeholder to go into would otherwise vanish without a word.
    /// </summary>
    public static string? Misfit(SlideData data, SlideLayoutPart layoutPart, string layoutName)
    {
        var placeholders = PlaceholdersOf(layoutPart);
        if (data.Title is { Length: > 0 } && !placeholders.Any(IsTitle))
            return $"Layout '{layoutName}' has no title placeholder, so the slide's title would be lost. " +
                "Choose a layout with a title, or leave the title out.";
        if (data.Body.Count > 0 && !placeholders.Any(TakesBody))
            return $"Layout '{layoutName}' has no text placeholder for the body, so it would be lost. " +
                "Choose a layout with a content or body placeholder, or leave the body out.";
        return null;
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

    private static string? KindOf(Shape shape) =>
        shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape is { } placeholder
            ? SlideLayouts.KindOf(placeholder)
            : null;

    private static bool TakesBody(Shape shape) =>
        KindOf(shape) is "content" or "body" or "subtitle";

    /// <summary>
    /// A slide shape naming the layout placeholder it inherits from, carrying only text.
    /// A picture, chart or table placeholder carries no text body, as PowerPoint writes it.
    /// </summary>
    private static Shape Filled(
        uint shapeId, string name, Shape layoutPlaceholder, IReadOnlyList<string> lines, bool holdsText = true)
    {
        var source = layoutPlaceholder.NonVisualShapeProperties?
            .ApplicationNonVisualDrawingProperties?.PlaceholderShape;

        var placeholder = new PlaceholderShape();
        if (source?.Type?.Value is { } type) placeholder.Type = type;
        if (source?.Index?.Value is { } index) placeholder.Index = index;

        var shape = new Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = shapeId, Name = name },
                new NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(placeholder)),
            // Deliberately empty: an empty p:spPr is what makes the shape inherit the
            // layout's geometry. Writing a transform here would pin it instead.
            new ShapeProperties());
        if (holdsText) shape.Append(Body(lines));
        return shape;
    }

    /// <summary>One paragraph per line, so a body renders as the bullets the layout styles.</summary>
    private static TextBody Body(IReadOnlyList<string> lines)
    {
        var body = new TextBody(new A.BodyProperties(), new A.ListStyle());

        // An empty placeholder still needs one paragraph, or there is nothing to address
        // when it is filled later.
        if (lines.Count == 0)
            body.Append(new A.Paragraph(new A.EndParagraphRunProperties { Language = "en-US" }));

        foreach (var line in lines)
            body.Append(new A.Paragraph(
                new A.Run(
                    new A.RunProperties { Language = "en-US" },
                    new A.Text(line))));

        return body;
    }
}
