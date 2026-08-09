using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeAgent.PowerPoint;

/// <summary>
/// Configures the running items along a slide's edge: footer text, slide number, and date.
/// </summary>
/// <remarks>
/// Each is a placeholder shape on the slide, inheriting its position from the layout the
/// way a title does. The slide number and an auto-updating date are <c>a:fld</c> fields
/// rather than literal text, so PowerPoint recomputes them - a slide number written as
/// text would be wrong the moment a slide was inserted ahead of it.
/// <para>
/// A slide has no header: PresentationML puts a header flag on <c>p:hf</c>, but it governs
/// notes and handout pages, which is why PowerPoint greys the box out on the Slide tab.
/// </para>
/// </remarks>
internal sealed class SlideHeaderFooterHandler : IOperationHandler
{
    /// <inheritdoc />
    public bool CanHandle(PlanOperation operation) => operation is HeaderFooterOp;

    /// <inheritdoc />
    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (HeaderFooterOp)operation;

        if (op.Footer is null && op.ShowFooter is null &&
            op.ShowSlideNumber is null && op.ShowDateTime is null && op.DateTime is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "headerFooter needs at least one of footer, showFooter, showSlideNumber, " +
                "showDateTime, or dateTime.", op.Target));

        var slides = Scope(context, op, out var error);
        if (error is not null) return OperationPreview.Fail(error);

        var parts = new List<string>();
        if (op.Footer is not null)
            parts.Add(op.Footer.Length == 0 ? "footer cleared" : $"footer \"{op.Footer}\"");
        if (op.ShowFooter is { } f) parts.Add($"footer {(f ? "shown" : "hidden")}");
        if (op.ShowSlideNumber is { } n) parts.Add($"slide number {(n ? "shown" : "hidden")}");
        if (op.ShowDateTime is { } d) parts.Add($"date {(d ? "shown" : "hidden")}");
        if (op.DateTime is { Length: > 0 } text) parts.Add($"date \"{text}\"");

        return OperationPreview.Ok(new ProposedChange
        {
            Target = op.Target,
            Verb = "headerFooter",
            Before = string.Empty,
            After = string.Join(", ", parts),
            Context = slides.Count == 1 ? $"slide {slides[0].Number}" : $"{slides.Count} slides",
            BlastRadius = slides.Count
        });
    }

    /// <inheritdoc />
    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (HeaderFooterOp)operation;
        var slides = Scope(context, op, out _);

        // The master must declare the placeholders, or a slide carrying one has nothing to
        // inherit its position from and PowerPoint drops it into the top-left corner.
        HeaderFooterLayout.EnsureOnMasterAndLayouts(context.Package);

        foreach (var slide in slides)
        {
            if (op.Footer is not null || op.ShowFooter is not null)
                Set(slide, PlaceholderValues.Footer,
                    show: op.ShowFooter ?? op.Footer is { Length: > 0 },
                    text: op.Footer,
                    field: null);

            if (op.ShowSlideNumber is { } showNumber)
                Set(slide, PlaceholderValues.SlideNumber,
                    show: showNumber,
                    text: null,
                    // The number is a field so PowerPoint renumbers it when slides move.
                    field: "slidenum");

            if (op.ShowDateTime is not null || op.DateTime is not null)
                Set(slide, PlaceholderValues.DateAndTime,
                    show: op.ShowDateTime ?? op.DateTime is { Length: > 0 },
                    text: op.DateTime,
                    // A null DateTime means "update automatically", PowerPoint's own default.
                    field: op.DateTime is null ? "datetime1" : null);
        }
    }

    /// <summary>The slides the operation covers: one, or the whole deck when untargeted.</summary>
    private static IReadOnlyList<SlideRef> Scope(
        ApplyContext context, HeaderFooterOp op, out ValidationError? error)
    {
        error = null;

        if (op.Target is null)
            return PowerPointModel.Slides(context.Package).ToList();

        if (op.Target is not NodeAnchor { Kind: "slide" } anchor)
        {
            error = new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "headerFooter targets a slide - { \"kind\": \"slide\", \"path\": \"slide#256\" } - " +
                "or no target at all to apply to every slide.",
                op.Target);
            return Array.Empty<SlideRef>();
        }

        var slide = SlideList.Target(context, anchor);
        if (slide is null)
        {
            error = SlideList.NoSuchSlide(anchor);
            return Array.Empty<SlideRef>();
        }

        return new[] { slide };
    }

    /// <summary>
    /// Adds, updates, or removes one running item on a slide. Hiding removes the shape
    /// rather than blanking it: an empty placeholder still shows PowerPoint's editing
    /// prompt, so a "hidden" footer would remain visible to whoever opened the deck.
    /// </summary>
    private static void Set(SlideRef slide, PlaceholderValues type, bool show, string? text, string? field)
    {
        var tree = slide.Part.Slide.CommonSlideData?.ShapeTree;
        if (tree is null) return;

        var existing = tree.Elements<Shape>().FirstOrDefault(s =>
            s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?
                .PlaceholderShape?.Type?.Value == type);

        if (!show)
        {
            existing?.Remove();
            return;
        }

        var shape = existing;
        if (shape is null)
        {
            shape = HeaderFooterLayout.SlideShape(PowerPointModel.NextShapeId(slide.Part), type);
            tree.Append(shape);
        }

        var body = shape.TextBody ??= new TextBody(new A.BodyProperties(), new A.ListStyle());
        body.RemoveAllChildren<A.Paragraph>();

        var paragraph = new A.Paragraph();
        if (field is not null)
        {
            paragraph.Append(new A.Field(
                new A.RunProperties { Language = "en-US" },
                new A.Text(text ?? DefaultFieldText(field)))
            {
                Id = "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}",
                Type = field
            });
        }
        else
        {
            paragraph.Append(new A.Run(
                new A.RunProperties { Language = "en-US" },
                new A.Text(text ?? string.Empty)));
        }

        body.Append(paragraph);
    }

    /// <summary>
    /// The literal a field carries as its last known value. PowerPoint recomputes it on
    /// open, but a reader that does not - or a thumbnail - shows this instead of nothing.
    /// </summary>
    private static string DefaultFieldText(string field) =>
        field == "slidenum" ? "1" : System.DateTime.Now.ToString("d");
}
