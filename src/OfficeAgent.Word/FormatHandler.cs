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

        if (!AreEdges(op.BorderEdges))
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"'{op.BorderEdges}' is not a border edge list. Expected a comma-separated subset of: {BorderEdgeNames}.",
                op.Target));

        if (op.ListStyle is not null && !WordNumbering.IsStyle(op.ListStyle))
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"'{op.ListStyle}' is not a list style. Expected one of: {WordNumbering.Names}.",
                op.Target));

        if (op.ListLevel is { } requested && (requested < 0 || requested > WordNumbering.MaxLevel))
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"listLevel must be between 0 and {WordNumbering.MaxLevel}; got {requested}.",
                op.Target));

        if (op.ListStyle is null && (op.ListLevel is not null || op.ListId is not null))
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "listLevel and listId need a listStyle to belong to.", op.Target));

        if (op.ColumnWidthsPx is { } columns)
        {
            if (op.Target is not NodeAnchor { Kind: "table" })
                return OperationPreview.Fail(new ValidationError(
                    ValidationErrorCodes.InvalidOperation,
                    "columnWidthsPx belongs on a table target.", op.Target));

            if (columns.Count == 0 || columns.Any(w => w <= 0))
                return OperationPreview.Fail(new ValidationError(
                    ValidationErrorCodes.InvalidOperation,
                    "columnWidthsPx needs a positive width for every column.", op.Target));
        }

        if (!HasAnyProperty(op))
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "format requires at least one of: styleId, fontFamily, sizeHalfPoints, bold, italic, underline, highlight, color, alignment, indent*, spacing*, border*, pageBreakBefore, listStyle, columnWidthsPx, widthPx, heightPx.",
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

        ApplyParagraphProperties(context, paragraph, op);

        IReadOnlyList<Run> runs;
        if (string.IsNullOrEmpty(anchor.Expect))
        {
            // Through the dialect, so a whole-paragraph format reaches runs nested in a
            // tracked insertion, hyperlink or content control exactly as a span format does.
            runs = WordModel.Dialect.GetRuns(paragraph).OfType<Run>().ToList();
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

        if (op.ColumnWidthsPx is { Count: > 0 } widths)
            SetColumnWidths(table, properties, widths);
    }

    /// <summary>
    /// Sets the column widths, in the three places Word reads them from: the grid, every
    /// cell, and the table's own width.
    /// </summary>
    /// <remarks>
    /// The grid alone is a hint. Word honours it until a cell's content is wider, then
    /// quietly reflows the whole table - so the cell widths have to agree with the grid, and
    /// the layout has to be fixed, or a long description drags its column open and squeezes
    /// the figures beside it.
    /// </remarks>
    private static void SetColumnWidths(
        Table table, TableProperties properties, IReadOnlyList<int> widthsPx)
    {
        // 1440 twips to the inch, 96 pixels to the inch.
        const int TwipsPerPixel = 15;

        var widths = widthsPx.Select(px => px * TwipsPerPixel).ToList();
        var total = widths.Sum();

        var grid = table.GetFirstChild<TableGrid>();
        if (grid is null)
        {
            grid = new TableGrid();
            table.InsertAfter(grid, properties);
        }

        grid.RemoveAllChildren<GridColumn>();
        foreach (var width in widths)
            grid.AppendChild(new GridColumn { Width = width.ToString() });

        foreach (var row in table.Elements<TableRow>())
        {
            var cells = row.Elements<TableCell>().ToList();
            for (var i = 0; i < cells.Count && i < widths.Count; i++)
            {
                var cellProperties = cells[i].TableCellProperties ??= new TableCellProperties();
                cellProperties.GetFirstChild<TableCellWidth>()?.Remove();
                cellProperties.InsertAt(
                    new TableCellWidth { Width = widths[i].ToString(), Type = TableWidthUnitValues.Dxa },
                    0);
            }
        }

        ReplaceChild(properties, new TableWidth
        {
            Width = total.ToString(),
            Type = TableWidthUnitValues.Dxa
        });
        ReplaceChild(properties, new TableLayout { Type = TableLayoutValues.Fixed });
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

        // A row has no borders of its own in WordprocessingML - the cells carry them. A
        // caller asking to rule a row means every cell in it, which is how a header row
        // gets its underline; ignoring the request instead would look like nothing happened.
        if (HasBorder(op))
            foreach (var cell in row.Elements<TableCell>())
            {
                var cellProperties = cell.GetFirstChild<TableCellProperties>()
                    ?? cell.InsertAt(new TableCellProperties(), 0)!;
                ReplaceChild(cellProperties, BuildTableCellBorders(op));
            }

        ApplyCharacterAndParagraphPropertiesToContainer(context, row, op);
    }

    // ── Target: table cell ────────────────────────────────────────────────

    private static void ApplyToCell(ApplyContext context, FormatOp op, NodeAnchor anchor)
    {
        var cell = TableLocator.FindCell(context.Package, anchor.Path)!;
        var properties = cell.GetFirstChild<TableCellProperties>() ?? cell.InsertAt(new TableCellProperties(), 0)!;

        if (HasBorder(op))
            ReplaceChild(properties, BuildTableCellBorders(op));

        ApplyCharacterAndParagraphPropertiesToContainer(context, cell, op);
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

    private static void ApplyCharacterAndParagraphPropertiesToContainer(ApplyContext context, OpenXmlElement container, FormatOp op)
    {
        foreach (var paragraph in container.Descendants<Paragraph>())
        {
            ApplyParagraphProperties(context, paragraph, op);
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

    private static void ApplyParagraphProperties(ApplyContext context, Paragraph paragraph, FormatOp op)
    {
        // Numbering is written first: it creates the w:pPr the rest of this then fills in,
        // and it owns the one child that has to sit ahead of the indent and the spacing.
        if (op.ListStyle is { Length: > 0 } list)
            WordNumbering.Apply(
                WordModel.Doc(context.Package).MainDocumentPart!,
                paragraph, list, op.ListLevel ?? 0, op.ListId ?? 0);

        if (op.StyleId is null && op.Alignment is null
            && op.IndentLeftTwips is null && op.IndentRightTwips is null
            && op.IndentFirstLineTwips is null
            && op.SpacingBeforeTwips is null && op.SpacingAfterTwips is null
            && op.PageBreakBefore is null
            && !HasBorder(op))
            return;

        var pPr = paragraph.ParagraphProperties ??= new ParagraphProperties();

        if (op.StyleId is not null)
            ReplaceChild(pPr, new ParagraphStyleId { Val = op.StyleId });

        // w:pageBreakBefore is a property of the paragraph rather than a break character in
        // the text, so the page still starts here after the text above it is edited.
        if (op.PageBreakBefore is bool breakBefore)
        {
            if (breakBefore) ReplaceChild(pPr, new PageBreakBefore());
            else pPr.GetFirstChild<PageBreakBefore>()?.Remove();
        }

        if (op.Alignment is not null)
            ReplaceChild(pPr, new Justification { Val = ParseAlignment(op.Alignment) });

        // Cloned from whatever is already there so an indent set earlier is not lost when
        // only one edge is changed now. The clone must come *before* the removal: calling
        // Remove on an element that was never in the tree throws "The parent of this
        // element is null", which is what a paragraph with no existing indent produced.
        if (op.IndentLeftTwips is not null || op.IndentRightTwips is not null || op.IndentFirstLineTwips is not null)
        {
            var existing = pPr.GetFirstChild<Indentation>();
            var ind = existing is null ? new Indentation() : (Indentation)existing.CloneNode(deep: true);
            if (op.IndentLeftTwips is int left) ind.Left = left.ToString();
            if (op.IndentRightTwips is int right) ind.Right = right.ToString();

            // WordprocessingML has no negative first-line indent: w:firstLine is unsigned,
            // and a hanging indent - the first line set back from the rest, which is how
            // every bullet and numbered list is built - is the separate w:hanging attribute.
            // Writing a negative into w:firstLine produces a document Word offers to repair.
            if (op.IndentFirstLineTwips is int firstLine)
            {
                if (firstLine < 0)
                {
                    ind.FirstLine = null;
                    ind.Hanging = (-firstLine).ToString();
                }
                else
                {
                    ind.Hanging = null;
                    ind.FirstLine = firstLine.ToString();
                }
            }
            ReplaceChild(pPr, ind);
        }

        if (op.SpacingBeforeTwips is not null || op.SpacingAfterTwips is not null)
        {
            var existing = pPr.GetFirstChild<SpacingBetweenLines>();
            var sp = existing is null ? new SpacingBetweenLines() : (SpacingBetweenLines)existing.CloneNode(deep: true);
            if (op.SpacingBeforeTwips is int b) sp.Before = b.ToString();
            if (op.SpacingAfterTwips is int a) sp.After = a.ToString();
            ReplaceChild(pPr, sp);
        }

        if (HasBorder(op))
            ReplaceChild(pPr, BuildParagraphBorders(op));
    }

    // ── Borders ──────────────────────────────────────────────────────────

    private static bool HasBorder(FormatOp op) =>
        op.BorderStyle is not null || op.BorderSizeEighths is not null ||
        op.BorderColor is not null || op.BorderEdges is not null;

    /// <summary>The edge names a plan may ask for, for an error message that helps.</summary>
    public const string BorderEdgeNames = "top, left, bottom, right, insideH, insideV";

    /// <summary>Whether every name in the list is one this module knows.</summary>
    public static bool AreEdges(string? edges) =>
        edges is null || Edges(edges).Count > 0;

    /// <summary>
    /// Parses the edge list. An unset list means all four sides plus, on a table, the
    /// inside rules - the behaviour a border had before edges could be named.
    /// </summary>
    private static HashSet<string> Edges(string? edges)
    {
        var all = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            { "top", "left", "bottom", "right", "insideh", "insidev" };

        if (string.IsNullOrWhiteSpace(edges)) return all;

        var named = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var edge in edges!.Split(','))
        {
            var trimmed = edge.Trim();
            if (trimmed.Length == 0) continue;
            if (!all.Contains(trimmed)) return new HashSet<string>();   // unknown: refuse
            named.Add(trimmed);
        }
        return named;
    }

    /// <summary>
    /// Builds the border elements for the named edges, in the order the schema declares
    /// them: top, left, bottom, right, insideH, insideV. Any other order and Word offers
    /// to repair the document.
    /// </summary>
    private static IEnumerable<OpenXmlElement> BorderEdgesOf(FormatOp op, bool inside)
    {
        var style = ParseBorderStyle(op.BorderStyle);
        var size = (uint)(op.BorderSizeEighths ?? 4);
        var color = op.BorderColor ?? "auto";
        var wanted = Edges(op.BorderEdges);

        if (wanted.Contains("top")) yield return new TopBorder { Val = style, Size = size, Color = color };
        if (wanted.Contains("left")) yield return new LeftBorder { Val = style, Size = size, Color = color };
        if (wanted.Contains("bottom")) yield return new BottomBorder { Val = style, Size = size, Color = color };
        if (wanted.Contains("right")) yield return new RightBorder { Val = style, Size = size, Color = color };

        if (!inside) yield break;
        if (wanted.Contains("insideH")) yield return new InsideHorizontalBorder { Val = style, Size = size, Color = color };
        if (wanted.Contains("insideV")) yield return new InsideVerticalBorder { Val = style, Size = size, Color = color };
    }

    private static ParagraphBorders BuildParagraphBorders(FormatOp op) =>
        new(BorderEdgesOf(op, inside: false).ToArray());

    private static TableBorders BuildTableBorders(FormatOp op) =>
        new(BorderEdgesOf(op, inside: true).ToArray());

    private static TableCellBorders BuildTableCellBorders(FormatOp op) =>
        new(BorderEdgesOf(op, inside: false).ToArray());

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
        op.PageBreakBefore is not null || op.ListStyle is not null ||
        op.ColumnWidthsPx is not null ||
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

    /// <summary>
    /// Replaces a property element, keeping the parent's children in the order the schema
    /// declares them.
    /// </summary>
    /// <remarks>
    /// <c>w:rPr</c>, <c>w:pPr</c>, <c>w:tblPr</c> and their kin are <em>sequences</em>, not
    /// bags: <c>w:tblStyle</c> has to precede <c>w:tblW</c>, <c>w:b</c> has to precede
    /// <c>w:color</c>, and so on. Appending produces a document Word offers to repair -
    /// and one that validates only by accident, when the properties happen to be applied
    /// in schema order. <see cref="OpenXmlCompositeElement.InsertAt"/> is positional, so
    /// the insertion point is computed from the SDK's own declared child order.
    /// </remarks>
    private static void ReplaceChild<T>(OpenXmlElement parent, T newChild) where T : OpenXmlElement
    {
        parent.GetFirstChild<T>()?.Remove();

        var order = ChildOrder(parent);
        var rank = Rank(order, newChild);
        if (rank < 0)
        {
            parent.AppendChild(newChild);
            return;
        }

        // The first existing child that belongs after this one.
        foreach (var child in parent.ChildElements)
        {
            var childRank = Rank(order, child);
            if (childRank >= 0 && childRank > rank)
            {
                parent.InsertBefore(newChild, child);
                return;
            }
        }

        parent.AppendChild(newChild);
    }

    private static int Rank(IReadOnlyList<Type> order, OpenXmlElement element)
    {
        for (var i = 0; i < order.Count; i++)
            if (order[i] == element.GetType()) return i;
        return -1;
    }

    /// <summary>
    /// The schema's child order for the property containers this handler writes into. Only
    /// the elements it actually produces need listing - anything else keeps its place.
    /// </summary>
    private static IReadOnlyList<Type> ChildOrder(OpenXmlElement parent) => parent switch
    {
        RunProperties => RunPropertyOrder,
        ParagraphProperties => ParagraphPropertyOrder,
        TableProperties => TablePropertyOrder,
        TableCellProperties => TableCellPropertyOrder,
        _ => Array.Empty<Type>()
    };

    private static readonly Type[] RunPropertyOrder =
    {
        typeof(RunStyle), typeof(RunFonts), typeof(Bold), typeof(BoldComplexScript),
        typeof(Italic), typeof(ItalicComplexScript), typeof(WColor), typeof(Spacing),
        typeof(FontSize), typeof(FontSizeComplexScript), typeof(Highlight), typeof(Underline)
    };

    /// <summary>
    /// The order <c>CT_PPr</c> declares these in. <c>w:pageBreakBefore</c> comes <em>before</em>
    /// <c>w:numPr</c>, not after it - putting a page break on a numbered paragraph is what
    /// turns that into a document Word offers to repair.
    /// </summary>
    internal static readonly Type[] ParagraphPropertyOrder =
    {
        typeof(ParagraphStyleId), typeof(PageBreakBefore), typeof(NumberingProperties),
        typeof(ParagraphBorders), typeof(SpacingBetweenLines), typeof(Indentation),
        typeof(Justification), typeof(OutlineLevel)
    };

    /// <summary>
    /// Places a child of <c>w:pPr</c> where the schema wants it. Shared so the numbering
    /// writer cannot drift from this one - the two disagreeing is exactly the bug this
    /// order exists to prevent.
    /// </summary>
    internal static void PlaceParagraphProperty(ParagraphProperties properties, OpenXmlElement child)
    {
        var rank = Rank(ParagraphPropertyOrder, child);
        if (rank < 0)
        {
            properties.AppendChild(child);
            return;
        }

        foreach (var existing in properties.ChildElements)
        {
            var existingRank = Rank(ParagraphPropertyOrder, existing);
            if (existingRank >= 0 && existingRank > rank)
            {
                properties.InsertBefore(child, existing);
                return;
            }
        }

        properties.AppendChild(child);
    }

    private static readonly Type[] TablePropertyOrder =
    {
        typeof(TableStyle), typeof(TableWidth), typeof(TableJustification),
        typeof(TableIndentation), typeof(TableBorders), typeof(Shading),
        typeof(TableLayout), typeof(TableCellMarginDefault), typeof(TableLook)
    };

    private static readonly Type[] TableCellPropertyOrder =
    {
        typeof(TableCellWidth), typeof(GridSpan), typeof(TableCellBorders),
        typeof(Shading), typeof(TableCellMargin), typeof(TableCellVerticalAlignment)
    };

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
