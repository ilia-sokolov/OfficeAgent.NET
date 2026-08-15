using DocumentFormat.OpenXml;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeAgent.PowerPoint;

/// <summary>
/// Applies character and paragraph formatting on a slide, in a table cell, or in the
/// slide notes, and resizes a picture.
/// </summary>
/// <remarks>
/// <para>
/// A text target formats exactly the anchored span: the run engine splits runs so the
/// covered set is the span and nothing either side of it changes. Formatting a whole
/// paragraph is expressed by anchoring the paragraph's full text.
/// </para>
/// <para>
/// The properties that carry over from the shared verb are the ones DrawingML actually
/// has: bold, italic, underline, size, font family, colour, highlight, and alignment.
/// The Word-only measures - twip indents and spacing, table borders - are refused rather
/// than silently ignored, because an agent that sees an operation succeed will believe
/// the deck now looks the way it asked for.
/// </para>
/// </remarks>
internal sealed class SlideFormatHandler : IOperationHandler
{
    private const long EmuPerPixel = 9525L;

    /// <inheritdoc />
    public bool CanHandle(PlanOperation operation) =>
        operation is FormatOp { Target: TextSpanAnchor } ||
        operation is FormatOp { Target: NodeAnchor { Kind: "image" } } ||
        operation is FormatOp { Target: NodeAnchor { Kind: "shape" } } ||
        operation is FormatOp { Target: NodeAnchor { Kind: "slide" } };

    /// <inheritdoc />
    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (FormatOp)operation;

        if (Unsupported(op) is { } unsupported)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation, unsupported, op.Target));

        return op.Target switch
        {
            NodeAnchor { Kind: "slide" } slide => PreviewBackground(op, slide, context),
            NodeAnchor { Kind: "shape" } shape => PreviewShape(op, shape, context),
            NodeAnchor node => PreviewImage(op, node, context),
            _ => PreviewText(op, (TextSpanAnchor)op.Target!, context)
        };
    }

    /// <inheritdoc />
    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (FormatOp)operation;

        if (op.Target is NodeAnchor { Kind: "slide" } slideAnchor)
        {
            var slide = SlideList.Target(context, slideAnchor)
                ?? throw new InvalidOperationException($"Slide '{slideAnchor.Path}' vanished before apply.");
            SlidePaint.SetBackground(slide, op.FillColor!);
            return;
        }

        if (op.Target is NodeAnchor { Kind: "shape" } shapeAnchor)
        {
            var located = ShapeNodeProvider.Locate(shapeAnchor.Path, context.Package)
                ?? throw new InvalidOperationException($"Shape '{shapeAnchor.Path}' vanished before apply.");
            located.Arrange(
                op.XPx is { } x ? Emu.FromPixels(x) : null,
                op.YPx is { } y ? Emu.FromPixels(y) : null,
                op.WidthPx is { } w ? Emu.FromPixels(w) : null,
                op.HeightPx is { } h ? Emu.FromPixels(h) : null);
            SlidePaint.SetShapeFill(located.Element, op.FillColor, op.LineColor, op.LineWidthPx);
            if (op.VerticalAlignment is { Length: > 0 } vertical)
                SlidePaint.SetVerticalAlignment(located.Element, vertical);
            return;
        }

        if (op.Target is NodeAnchor node)
        {
            var picture = SlideImageNodeProvider.Locate(node.Path, new PowerPointObjectMap(context.Package))
                ?? throw new InvalidOperationException($"Image '{node.Path}' vanished before apply.");
            Resize(picture.Picture, op);
            return;
        }

        var anchor = (TextSpanAnchor)op.Target!;
        var paragraph = PowerPointModel.ResolveParagraph(context, anchor.ParaId)
            ?? throw new InvalidOperationException($"Paragraph '{anchor.ParaId}' vanished before apply.");

        foreach (var run in Covered(paragraph, anchor))
            ApplyRunProperties(run, op);

        ApplyParagraphProperties(paragraph.Paragraph, op);
    }

    // ── text ──────────────────────────────────────────────────────────────────

    private static OperationPreview PreviewText(FormatOp op, TextSpanAnchor anchor, ApplyContext context)
    {
        var paragraph = PowerPointModel.ResolveParagraph(context, anchor.ParaId);
        if (paragraph is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No paragraph with id '{anchor.ParaId}'.", anchor));

        var text = PowerPointModel.TextOf(paragraph.Paragraph);

        // An empty expect means the whole paragraph, which is how an agent formats a
        // heading or a bullet without having to restate its text.
        if (!string.IsNullOrEmpty(anchor.Expect))
        {
            var comparison = PowerPointModel.Comparison(caseSensitive: true);
            var occurrences = PowerPointModel.Text.CountOccurrences(text, anchor.Expect, comparison);
            if (occurrences == 0)
                return OperationPreview.Fail(new ValidationError(
                    ValidationErrorCodes.ExpectMismatch,
                    $"Expected text '{anchor.Expect}' not found in paragraph '{anchor.ParaId}' (the deck drifted).",
                    anchor));
            if (PowerPointModel.Text.IndexOfOccurrence(text, anchor.Expect, anchor.Occurrence, comparison) < 0)
                return OperationPreview.Fail(new ValidationError(
                    ValidationErrorCodes.AmbiguousAnchor,
                    $"Occurrence {anchor.Occurrence} of '{anchor.Expect}' does not exist ({occurrences} found).",
                    anchor));
        }

        var described = Describe(op);
        if (described.Length == 0)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "format needs at least one property to change.", anchor));

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "format",
            Before = string.IsNullOrEmpty(anchor.Expect) ? text : anchor.Expect,
            After = described,
            Context = $"slide {paragraph.Slide.Number} ({paragraph.Location})",
            BlastRadius = 1
        });
    }

    /// <summary>The runs the anchor covers: the whole paragraph, or just the named span.</summary>
    private static IReadOnlyList<OpenXmlElement> Covered(ParagraphRef paragraph, TextSpanAnchor anchor)
    {
        if (string.IsNullOrEmpty(anchor.Expect))
            return paragraph.Paragraph.Elements<A.Run>().Cast<OpenXmlElement>().ToList();

        var text = PowerPointModel.TextOf(paragraph.Paragraph);
        var start = PowerPointModel.Text.IndexOfOccurrence(
            text, anchor.Expect, anchor.Occurrence, PowerPointModel.Comparison(caseSensitive: true));
        if (start < 0)
            throw new InvalidOperationException($"Expected text '{anchor.Expect}' not found at apply time.");

        return PowerPointModel.Text.IsolateSpan(paragraph.Paragraph, start, anchor.Expect.Length);
    }

    /// <summary>
    /// Writes the run properties. <c>a:rPr</c> is a sequence, so each child is placed in
    /// schema order rather than appended: fill, then highlight, then the font list.
    /// </summary>
    private static void ApplyRunProperties(OpenXmlElement element, FormatOp op)
    {
        if (element is not A.Run run) return;

        var properties = run.RunProperties;
        if (properties is null)
        {
            properties = new A.RunProperties { Language = "en-US" };
            run.InsertAt(properties, 0);
        }

        if (op.Bold is { } bold) properties.Bold = bold;
        if (op.Italic is { } italic) properties.Italic = italic;
        if (op.Underline is { } underline)
            properties.Underline = underline ? A.TextUnderlineValues.Single : A.TextUnderlineValues.None;
        if (op.SizeHalfPoints is { } half) properties.FontSize = half * 50;   // half-points → hundredths of a point

        if (op.Color is { Length: > 0 } color)
        {
            foreach (var existing in properties.Elements<A.SolidFill>().ToList()) existing.Remove();
            properties.InsertAt(new A.SolidFill(new A.RgbColorModelHex { Val = Hex(color) }), 0);
        }

        if (op.Highlight is { Length: > 0 } highlight)
        {
            foreach (var existing in properties.Elements<A.Highlight>().ToList()) existing.Remove();
            if (!string.Equals(highlight, "none", StringComparison.OrdinalIgnoreCase))
            {
                var element_ = new A.Highlight(new A.RgbColorModelHex { Val = HighlightHex(highlight) });
                var fill = properties.Elements<A.SolidFill>().FirstOrDefault();
                if (fill is null) properties.InsertAt(element_, 0);
                else fill.InsertAfterSelf(element_);
            }
        }

        if (op.FontFamily is { Length: > 0 } font)
        {
            foreach (var existing in properties.Elements<A.LatinFont>().ToList()) existing.Remove();
            properties.AppendChild(new A.LatinFont { Typeface = font });
        }
    }

    private static void ApplyParagraphProperties(A.Paragraph paragraph, FormatOp op)
    {
        if (op.Alignment is not { Length: > 0 } alignment) return;

        var properties = paragraph.ParagraphProperties;
        if (properties is null)
        {
            properties = new A.ParagraphProperties();
            paragraph.InsertAt(properties, 0);
        }
        properties.Alignment = alignment.ToLowerInvariant() switch
        {
            "center" or "centre" => A.TextAlignmentTypeValues.Center,
            "right" => A.TextAlignmentTypeValues.Right,
            "justify" or "both" => A.TextAlignmentTypeValues.Justified,
            _ => A.TextAlignmentTypeValues.Left
        };
    }

    // ── image ─────────────────────────────────────────────────────────────────

    private static OperationPreview PreviewImage(FormatOp op, NodeAnchor anchor, ApplyContext context)
    {
        var picture = SlideImageNodeProvider.Locate(anchor.Path, new PowerPointObjectMap(context.Package));
        if (picture is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No image with path '{anchor.Path}'. Image paths come from inspect_document.nodes.",
                anchor));

        if (op.WidthPx is null && op.HeightPx is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "Formatting an image changes its size, so widthPx or heightPx is required.", anchor));
        if (op.WidthPx is <= 0 || op.HeightPx is <= 0)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "widthPx and heightPx must be positive.", anchor));

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "format",
            Before = string.Empty,
            After = $"[image {op.WidthPx?.ToString() ?? "auto"}×{op.HeightPx?.ToString() ?? "auto"}]",
            Context = $"slide {picture.Slide.Number}",
            BlastRadius = 1
        });
    }

    private static void Resize(P.Picture picture, FormatOp op)
    {
        var transform = picture.ShapeProperties?.Transform2D;
        if (transform is null) return;

        transform.Extents ??= new A.Extents();
        if (op.WidthPx is { } width) transform.Extents.Cx = width * EmuPerPixel;
        if (op.HeightPx is { } height) transform.Extents.Cy = height * EmuPerPixel;
    }

    // ── slide background ──────────────────────────────────────────────────────

    /// <summary>
    /// A slide target paints the background and nothing else. The rest of the vocabulary
    /// needs something to apply to - a run, a paragraph, a box - and a slide is none of
    /// them, so accepting bold here would change nothing while appearing to succeed.
    /// </summary>
    private static OperationPreview PreviewBackground(FormatOp op, NodeAnchor anchor, ApplyContext context)
    {
        var slide = SlideList.Target(context, anchor);
        if (slide is null) return OperationPreview.Fail(SlideList.NoSuchSlide(anchor));

        if (op.FillColor is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "Formatting a slide sets its background, so fillColor is required - a hex " +
                "colour such as \"1F3A5F\", or \"none\" to clear it. To style the slide's " +
                "text, target a paragraph; to style a box on it, target a shape.",
                anchor));

        if (!SlidePaint.IsColour(op.FillColor))
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"'{op.FillColor}' is not a colour. Use six hex digits such as \"1F3A5F\", or \"none\".",
                anchor));

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "format",
            Before = slide.Part.Slide.CommonSlideData?.Background is null ? "inherited" : "painted",
            After = SlidePaint.IsNone(op.FillColor) ? "background cleared" : $"background #{SlidePaint.Hex(op.FillColor)}",
            Context = $"slide {slide.Number}",
            BlastRadius = 1
        });
    }

    // ── shape geometry ────────────────────────────────────────────────────────

    /// <summary>
    /// Moving and resizing any shape - a text box, a table frame, a picture. Geometry is
    /// the only thing a shape-targeted format changes; run and paragraph properties need a
    /// text anchor, because a shape may hold many paragraphs and styling all of them is
    /// rarely what was meant.
    /// </summary>
    private static OperationPreview PreviewShape(FormatOp op, NodeAnchor anchor, ApplyContext context)
    {
        var located = ShapeNodeProvider.Locate(anchor.Path, context.Package);
        if (located is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No shape with path '{anchor.Path}'. Shape paths come from inspect_document.nodes.",
                anchor));

        if (op.WidthPx is null && op.HeightPx is null && op.XPx is null && op.YPx is null &&
            op.FillColor is null && op.LineColor is null && op.LineWidthPx is null &&
            op.VerticalAlignment is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "Formatting a shape moves, resizes or paints it, so one of widthPx, heightPx, " +
                "xPx, yPx, fillColor, lineColor, lineWidthPx or verticalAlignment is required. " +
                "To style its text, target a paragraph instead.",
                anchor));

        foreach (var (name, value) in new[] { ("fillColor", op.FillColor), ("lineColor", op.LineColor) })
            if (value is not null && !SlidePaint.IsColour(value))
                return OperationPreview.Fail(new ValidationError(
                    ValidationErrorCodes.InvalidOperation,
                    $"'{value}' is not a colour for {name}. Use six hex digits such as \"1F3A5F\", or \"none\".",
                    anchor));

        if (op.LineWidthPx is < 0)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation, "lineWidthPx cannot be negative.", anchor));

        if (op.VerticalAlignment is not null && !SlidePaint.IsAnchor(op.VerticalAlignment))
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"'{op.VerticalAlignment}' is not a vertical alignment. Expected top, middle, or bottom.",
                anchor));

        if (op.WidthPx is <= 0 || op.HeightPx is <= 0)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "widthPx and heightPx must be positive.", anchor));

        if (op.XPx is < 0 || op.YPx is < 0)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "xPx and yPx are measured from the slide's top-left corner and cannot be negative.",
                anchor));

        var box = located.Box();
        var before = box is { } b
            ? $"{b.X / Emu.PerPixel},{b.Y / Emu.PerPixel} {b.Cx / Emu.PerPixel}×{b.Cy / Emu.PerPixel}"
            : "inherited from layout";

        var parts = new List<string>();
        if (op.XPx is not null || op.YPx is not null) parts.Add($"at {op.XPx?.ToString() ?? "="},{op.YPx?.ToString() ?? "="}");
        if (op.WidthPx is not null || op.HeightPx is not null)
            parts.Add($"{op.WidthPx?.ToString() ?? "auto"}×{op.HeightPx?.ToString() ?? "auto"}");
        if (op.FillColor is not null) parts.Add(SlidePaint.IsNone(op.FillColor) ? "no fill" : $"fill #{SlidePaint.Hex(op.FillColor)}");
        if (op.LineColor is not null) parts.Add(SlidePaint.IsNone(op.LineColor) ? "no outline" : $"outline #{SlidePaint.Hex(op.LineColor)}");
        if (op.LineWidthPx is { } lw) parts.Add($"outline {lw}px");
        if (op.VerticalAlignment is { Length: > 0 } va) parts.Add($"text {va}");

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "format",
            Before = before,
            After = string.Join(", ", parts),
            Context = $"slide {located.Slide.Number}",
            BlastRadius = 1
        });
    }

    // ── vocabulary ────────────────────────────────────────────────────────────

    /// <summary>
    /// Names the properties this format cannot honour on a slide, so the operation fails
    /// instead of appearing to succeed while changing nothing.
    /// </summary>
    private static string? Unsupported(FormatOp op)
    {
        var rejected = new List<string>();
        if (op.StyleId is { Length: > 0 }) rejected.Add("styleId (a deck has no paragraph style table)");
        if (op.IndentLeftTwips is not null || op.IndentRightTwips is not null ||
            op.IndentFirstLineTwips is not null) rejected.Add("indents");
        if (op.SpacingBeforeTwips is not null || op.SpacingAfterTwips is not null) rejected.Add("spacing");
        if (op.BorderStyle is { Length: > 0 } || op.BorderSizeEighths is not null ||
            op.BorderColor is { Length: > 0 } || op.BorderEdges is { Length: > 0 })
            rejected.Add("borders");
        if (op.PageBreakBefore is not null)
            rejected.Add("pageBreakBefore (a deck has slides, not pages - use insertSlide)");
        if (op.ColumnWidthsPx is not null)
            rejected.Add("columnWidthsPx (a deck's table follows its frame - resize that with widthPx)");
        if (op.ListStyle is { Length: > 0 } || op.ListLevel is not null || op.ListId is not null)
            rejected.Add("listStyle/listLevel/listId (a slide's bullets come from its layout - " +
                         "use the 'level' on insert for depth)");

        return rejected.Count == 0
            ? null
            : $"The PowerPoint module cannot apply {string.Join(", ", rejected)}. " +
              "Supported here: bold, italic, underline, sizeHalfPoints, fontFamily, color, highlight, " +
              "alignment, widthPx/heightPx/xPx/yPx on a shape or image.";
    }

    private static string Describe(FormatOp op)
    {
        var parts = new List<string>();
        if (op.Bold is { } b) parts.Add(b ? "bold" : "not bold");
        if (op.Italic is { } i) parts.Add(i ? "italic" : "not italic");
        if (op.Underline is { } u) parts.Add(u ? "underline" : "no underline");
        if (op.SizeHalfPoints is { } s) parts.Add($"{s / 2.0:0.#}pt");
        if (op.FontFamily is { Length: > 0 } f) parts.Add(f);
        if (op.Color is { Length: > 0 } c) parts.Add($"color {c}");
        if (op.Highlight is { Length: > 0 } h) parts.Add($"highlight {h}");
        if (op.Alignment is { Length: > 0 } a) parts.Add($"align {a}");
        return string.Join(", ", parts);
    }

    private static string Hex(string value) =>
        value.TrimStart('#').ToUpperInvariant();

    /// <summary>
    /// DrawingML has no named highlight vocabulary the way WordprocessingML does, so the
    /// shared names are mapped to the colours Office uses for them.
    /// </summary>
    private static string HighlightHex(string name) => name.ToLowerInvariant() switch
    {
        "yellow" => "FFFF00",
        "green" => "00FF00",
        "cyan" => "00FFFF",
        "magenta" => "FF00FF",
        "blue" => "0000FF",
        "red" => "FF0000",
        "darkblue" => "000080",
        "darkcyan" => "008080",
        "darkgreen" => "008000",
        "darkmagenta" => "800080",
        "darkred" => "800000",
        "darkyellow" => "808000",
        "darkgray" or "darkgrey" => "808080",
        "lightgray" or "lightgrey" => "C0C0C0",
        "black" => "000000",
        "white" => "FFFFFF",
        _ => Hex(name)
    };
}
