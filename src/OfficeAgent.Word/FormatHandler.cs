using DocumentFormat.OpenXml;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using WColor = DocumentFormat.OpenXml.Wordprocessing.Color;

namespace OfficeAgent.Word;

/// <summary>
/// Unified format handler. Dispatches on the target type and applies whichever
/// properties of the <see cref="FormatOp"/> are set: a named style id, character
/// properties (font/size/bold/italic/underline/highlight/colour), paragraph
/// properties (alignment/indent/spacing), borders, and dimensions.
/// </summary>
internal sealed class FormatHandler : IOperationHandler
{
    // 9525 EMU per pixel at 96 DPI.
    private const long EmuPerPixel = 9525;

    public bool CanHandle(PlanOperation operation) =>
        operation is FormatOp { Target: TextSpanAnchor }
        || operation is FormatOp { Target: NodeAnchor { Kind: "table" or "tableRow" or "tableCell" or "image" } };

    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (FormatOp)operation;
        if (!HasAnyProperty(op))
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "format requires at least one of: styleId, fontFamily, sizeHalfPoints, bold, italic, underline, highlight, color, alignment, indent*, spacing*, border*, widthPx, heightPx.",
                op.Target));

        switch (op.Target)
        {
            case TextSpanAnchor ts:
                if (WordModel.ResolveParagraph(context, ts.ParaId) is null)
                    return OperationPreview.Fail(new ValidationError(
                        ValidationErrorCodes.AnchorNotFound,
                        $"No paragraph with id '{ts.ParaId}'.", ts));
                return Ok(op, ts, "paragraph/span");

            case NodeAnchor n when n.Kind == "table":
                if (TableLocator.FindTable(context.Package, n.Path) is null)
                    return OperationPreview.Fail(new ValidationError(
                        ValidationErrorCodes.AnchorNotFound, $"No table at '{n.Path}'.", n));
                return Ok(op, n, "table");

            case NodeAnchor n when n.Kind == "tableRow":
                if (TableLocator.FindRow(context.Package, n.Path) is null)
                    return OperationPreview.Fail(new ValidationError(
                        ValidationErrorCodes.AnchorNotFound, $"No table row at '{n.Path}'.", n));
                return Ok(op, n, "tableRow");

            case NodeAnchor n when n.Kind == "tableCell":
                if (TableLocator.FindCell(context.Package, n.Path) is null)
                    return OperationPreview.Fail(new ValidationError(
                        ValidationErrorCodes.AnchorNotFound, $"No table cell at '{n.Path}'.", n));
                return Ok(op, n, "tableCell");

            case NodeAnchor n when n.Kind == "image":
                if (FindDrawing(context.Package, n.Path) is null)
                    return OperationPreview.Fail(new ValidationError(
                        ValidationErrorCodes.AnchorNotFound, $"No image at '{n.Path}'.", n));
                return Ok(op, n, "image");
        }

        return OperationPreview.Fail(new ValidationError(
            ValidationErrorCodes.UnsupportedOperation,
            "format only supports text-span, table, tableRow, tableCell, and image targets.",
            op.Target));
    }

    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (FormatOp)operation;
        switch (op.Target)
        {
            case TextSpanAnchor ts: ApplyToParagraphSpan(context, op, ts); break;
            case NodeAnchor n when n.Kind == "table": ApplyToTable(context, op, n); break;
            case NodeAnchor n when n.Kind == "tableRow": ApplyToRow(context, op, n); break;
            case NodeAnchor n when n.Kind == "tableCell": ApplyToCell(context, op, n); break;
            case NodeAnchor n when n.Kind == "image": ApplyToImage(context, op, n); break;
        }
    }

    // ── Target: paragraph + span runs ─────────────────────────────────────

    private static void ApplyToParagraphSpan(ApplyContext context, FormatOp op, TextSpanAnchor anchor)
    {
        var paragraph = WordModel.ResolveParagraph(context, anchor.ParaId)
            ?? throw new InvalidOperationException($"Paragraph '{anchor.ParaId}' vanished before apply.");

        ApplyParagraphProperties(paragraph, op);

        IReadOnlyList<Run> runs;
        if (string.IsNullOrEmpty(anchor.Expect))
        {
            runs = paragraph.Elements<Run>().ToList();
        }
        else
        {
            var text = WordModel.Text.GetLogicalText(paragraph);
            int start = WordModel.Text.IndexOfOccurrence(text, anchor.Expect, anchor.Occurrence, WordModel.Comparison(true));
            if (start < 0)
                throw new InvalidOperationException($"Expected text '{anchor.Expect}' not found at apply time.");
            runs = WordModel.Text.IsolateSpan(paragraph, start, anchor.Expect.Length).OfType<Run>().ToList();
        }

        foreach (var run in runs)
            ApplyRunProperties(run, op);
    }

    // ── Target: whole table ───────────────────────────────────────────────

    private static void ApplyToTable(ApplyContext context, FormatOp op, NodeAnchor anchor)
    {
        var table = TableLocator.FindTable(context.Package, anchor.Path)!;
        var properties = table.GetFirstChild<TableProperties>() ?? table.InsertAt(new TableProperties(), 0)!;

        if (op.StyleId is not null)
        {
            ReplaceChild(properties, new TableStyle { Val = op.StyleId });
        }

        if (HasBorder(op))
        {
            var borders = BuildTableBorders(op);
            ReplaceChild(properties, borders);
        }
    }

    // ── Target: table row ─────────────────────────────────────────────────

    private static void ApplyToRow(ApplyContext context, FormatOp op, NodeAnchor anchor)
    {
        var row = TableLocator.FindRow(context.Package, anchor.Path)!;

        if (op.HeightPx is int h)
        {
            var rowProperties = row.GetFirstChild<TableRowProperties>() ?? row.InsertAt(new TableRowProperties(), 0)!;
            ReplaceChild(rowProperties, new TableRowHeight { Val = (uint)(h * 15) }); // 1px ≈ 15 twips at 96dpi
        }

        ApplyCharacterAndParagraphPropertiesToContainer(row, op);
    }

    // ── Target: table cell ────────────────────────────────────────────────

    private static void ApplyToCell(ApplyContext context, FormatOp op, NodeAnchor anchor)
    {
        var cell = TableLocator.FindCell(context.Package, anchor.Path)!;
        var properties = cell.GetFirstChild<TableCellProperties>() ?? cell.InsertAt(new TableCellProperties(), 0)!;

        if (HasBorder(op))
            ReplaceChild(properties, BuildTableCellBorders(op));

        ApplyCharacterAndParagraphPropertiesToContainer(cell, op);
    }

    // ── Target: image ─────────────────────────────────────────────────────

    private static void ApplyToImage(ApplyContext context, FormatOp op, NodeAnchor anchor)
    {
        var drawing = FindDrawing(context.Package, anchor.Path)!;
        if (op.WidthPx is null && op.HeightPx is null) return;

        var inline = drawing.GetFirstChild<DW.Inline>();
        if (inline?.Extent is { } extent)
        {
            if (op.WidthPx is int w) extent.Cx = w * EmuPerPixel;
            if (op.HeightPx is int h) extent.Cy = h * EmuPerPixel;
        }

        var picExtents = drawing.Descendants<A.Extents>().FirstOrDefault();
        if (picExtents is not null)
        {
            if (op.WidthPx is int w2) picExtents.Cx = w2 * EmuPerPixel;
            if (op.HeightPx is int h2) picExtents.Cy = h2 * EmuPerPixel;
        }
    }

    // ── Helpers ──────────────────────────────────────────────────────────

    private static void ApplyCharacterAndParagraphPropertiesToContainer(OpenXmlElement container, FormatOp op)
    {
        foreach (var paragraph in container.Descendants<Paragraph>())
        {
            ApplyParagraphProperties(paragraph, op);
            foreach (var run in paragraph.Elements<Run>())
                ApplyRunProperties(run, op);
        }
    }

    private static void ApplyRunProperties(Run run, FormatOp op)
    {
        var rPr = run.RunProperties ??= new RunProperties();

        if (op.FontFamily is { Length: > 0 } font)
            ReplaceChild(rPr, new RunFonts { Ascii = font, HighAnsi = font, ComplexScript = font });
        if (op.SizeHalfPoints is int sz)
        {
            ReplaceChild(rPr, new FontSize { Val = sz.ToString() });
            ReplaceChild(rPr, new FontSizeComplexScript { Val = sz.ToString() });
        }
        if (op.Bold == true) ReplaceChild(rPr, new Bold());
        if (op.Italic == true) ReplaceChild(rPr, new Italic());
        if (op.Underline == true) ReplaceChild(rPr, new Underline { Val = UnderlineValues.Single });
        if (op.Highlight is not null) ReplaceChild(rPr, new Highlight { Val = ParseHighlight(op.Highlight) });
        if (op.Color is not null) ReplaceChild(rPr, new WColor { Val = op.Color });
    }

    private static void ApplyParagraphProperties(Paragraph paragraph, FormatOp op)
    {
        if (op.StyleId is null && op.Alignment is null
            && op.IndentLeftTwips is null && op.IndentRightTwips is null
            && op.IndentFirstLineTwips is null
            && op.SpacingBeforeTwips is null && op.SpacingAfterTwips is null
            && !HasBorder(op))
            return;

        var pPr = paragraph.ParagraphProperties ??= new ParagraphProperties();

        if (op.StyleId is not null)
            ReplaceChild(pPr, new ParagraphStyleId { Val = op.StyleId });

        if (op.Alignment is not null)
            ReplaceChild(pPr, new Justification { Val = ParseAlignment(op.Alignment) });

        if (op.IndentLeftTwips is not null || op.IndentRightTwips is not null || op.IndentFirstLineTwips is not null)
        {
            var existing = pPr.GetFirstChild<Indentation>() ?? new Indentation();
            existing.Remove();
            var ind = (Indentation)existing.CloneNode(deep: true);
            if (op.IndentLeftTwips is int left) ind.Left = left.ToString();
            if (op.IndentRightTwips is int right) ind.Right = right.ToString();
            if (op.IndentFirstLineTwips is int firstLine) ind.FirstLine = firstLine.ToString();
            pPr.AppendChild(ind);
        }

        if (op.SpacingBeforeTwips is not null || op.SpacingAfterTwips is not null)
        {
            var existing = pPr.GetFirstChild<SpacingBetweenLines>() ?? new SpacingBetweenLines();
            existing.Remove();
            var sp = (SpacingBetweenLines)existing.CloneNode(deep: true);
            if (op.SpacingBeforeTwips is int b) sp.Before = b.ToString();
            if (op.SpacingAfterTwips is int a) sp.After = a.ToString();
            pPr.AppendChild(sp);
        }

        if (HasBorder(op))
            ReplaceChild(pPr, BuildParagraphBorders(op));
    }

    // ── Borders ──────────────────────────────────────────────────────────

    private static bool HasBorder(FormatOp op) =>
        op.BorderStyle is not null || op.BorderSizeEighths is not null || op.BorderColor is not null;

    private static ParagraphBorders BuildParagraphBorders(FormatOp op)
    {
        var style = ParseBorderStyle(op.BorderStyle);
        var size = (uint)(op.BorderSizeEighths ?? 4);
        var color = op.BorderColor ?? "auto";
        return new ParagraphBorders(
            new TopBorder        { Val = style, Size = size, Color = color },
            new BottomBorder     { Val = style, Size = size, Color = color },
            new LeftBorder       { Val = style, Size = size, Color = color },
            new RightBorder      { Val = style, Size = size, Color = color });
    }

    private static TableBorders BuildTableBorders(FormatOp op)
    {
        var style = ParseBorderStyle(op.BorderStyle);
        var size = (uint)(op.BorderSizeEighths ?? 4);
        var color = op.BorderColor ?? "auto";
        return new TableBorders(
            new TopBorder { Val = style, Size = size, Color = color },
            new BottomBorder { Val = style, Size = size, Color = color },
            new LeftBorder { Val = style, Size = size, Color = color },
            new RightBorder { Val = style, Size = size, Color = color },
            new InsideHorizontalBorder { Val = style, Size = size, Color = color },
            new InsideVerticalBorder { Val = style, Size = size, Color = color });
    }

    private static TableCellBorders BuildTableCellBorders(FormatOp op)
    {
        var style = ParseBorderStyle(op.BorderStyle);
        var size = (uint)(op.BorderSizeEighths ?? 4);
        var color = op.BorderColor ?? "auto";
        return new TableCellBorders(
            new TopBorder { Val = style, Size = size, Color = color },
            new BottomBorder { Val = style, Size = size, Color = color },
            new LeftBorder { Val = style, Size = size, Color = color },
            new RightBorder { Val = style, Size = size, Color = color });
    }

    // ── Parsing ──────────────────────────────────────────────────────────

    private static JustificationValues ParseAlignment(string a) => a.Trim().ToLowerInvariant() switch
    {
        "left" or "start" => JustificationValues.Left,
        "center" or "centre" => JustificationValues.Center,
        "right" or "end" => JustificationValues.Right,
        "justify" or "both" => JustificationValues.Both,
        _ => throw new ArgumentException($"Unknown alignment '{a}'. Use left, center, right, or justify.")
    };

    private static BorderValues ParseBorderStyle(string? s) => (s ?? "single").Trim().ToLowerInvariant() switch
    {
        "single" => BorderValues.Single,
        "double" => BorderValues.Double,
        "dotted" => BorderValues.Dotted,
        "dashed" => BorderValues.Dashed,
        "thick" => BorderValues.Thick,
        "none" => BorderValues.None,
        _ => BorderValues.Single
    };

    private static HighlightColorValues ParseHighlight(string name)
    {
        var n = name.Trim().ToLowerInvariant();
        return n switch
        {
            "yellow" => HighlightColorValues.Yellow,
            "green" => HighlightColorValues.Green,
            "cyan" => HighlightColorValues.Cyan,
            "magenta" => HighlightColorValues.Magenta,
            "blue" => HighlightColorValues.Blue,
            "red" => HighlightColorValues.Red,
            "darkblue" => HighlightColorValues.DarkBlue,
            "darkcyan" => HighlightColorValues.DarkCyan,
            "darkgreen" => HighlightColorValues.DarkGreen,
            "darkmagenta" => HighlightColorValues.DarkMagenta,
            "darkred" => HighlightColorValues.DarkRed,
            "darkyellow" => HighlightColorValues.DarkYellow,
            "darkgray" or "darkgrey" => HighlightColorValues.DarkGray,
            "lightgray" or "lightgrey" => HighlightColorValues.LightGray,
            "black" => HighlightColorValues.Black,
            "white" => HighlightColorValues.White,
            "none" => HighlightColorValues.None,
            _ => throw new ArgumentException($"Unknown highlight '{name}'.")
        };
    }

    private static bool HasAnyProperty(FormatOp op) =>
        op.StyleId is not null ||
        op.FontFamily is not null || op.SizeHalfPoints is not null ||
        op.Bold is not null || op.Italic is not null || op.Underline is not null ||
        op.Highlight is not null || op.Color is not null ||
        op.Alignment is not null ||
        op.IndentLeftTwips is not null || op.IndentRightTwips is not null || op.IndentFirstLineTwips is not null ||
        op.SpacingBeforeTwips is not null || op.SpacingAfterTwips is not null ||
        HasBorder(op) ||
        op.WidthPx is not null || op.HeightPx is not null;

    private static OperationPreview Ok(FormatOp op, Anchor target, string scope)
    {
        var props = new List<string>();
        if (op.StyleId is not null) props.Add($"styleId={op.StyleId}");
        if (op.FontFamily is not null) props.Add($"font={op.FontFamily}");
        if (op.SizeHalfPoints is not null) props.Add($"size={op.SizeHalfPoints / 2.0}pt");
        if (op.Bold == true) props.Add("bold");
        if (op.Italic == true) props.Add("italic");
        if (op.Underline == true) props.Add("underline");
        if (op.Highlight is not null) props.Add($"highlight={op.Highlight}");
        if (op.Color is not null) props.Add($"color=#{op.Color}");
        if (op.Alignment is not null) props.Add($"align={op.Alignment}");
        if (HasBorder(op)) props.Add($"border={op.BorderStyle ?? "single"}");
        if (op.WidthPx is not null || op.HeightPx is not null) props.Add($"{op.WidthPx}x{op.HeightPx}px");

        return OperationPreview.Ok(new ProposedChange
        {
            Target = target,
            Verb = "format",
            Before = scope,
            After = string.Join(",", props),
            Context = target is TextSpanAnchor ts ? ts.ParaId : ((NodeAnchor)target).Path,
            BlastRadius = 1
        });
    }

    private static void ReplaceChild<T>(OpenXmlElement parent, T newChild) where T : OpenXmlElement
    {
        var existing = parent.GetFirstChild<T>();
        if (existing is not null) existing.Remove();
        parent.AppendChild(newChild);
    }

    private static Drawing? FindDrawing(IOpenXmlPackage package, string path)
    {
        if (!path.StartsWith("image#", System.StringComparison.Ordinal)) return null;
        if (!int.TryParse(path.Substring("image#".Length), out var target)) return null;
        int i = 0;
        foreach (var d in ImageNodeProvider.EnumerateDrawings(package))
        {
            if (i == target) return d;
            i++;
        }
        return null;
    }
}
