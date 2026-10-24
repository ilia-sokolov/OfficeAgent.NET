using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;

namespace OfficeAgent.Word;

/// <summary>
/// Replaces every cell of a Word table, growing or shrinking it by rows, and keeps each cell's
/// paragraph and character formatting.
/// </summary>
internal sealed class FillTableHandler : IOperationHandler
{
    private readonly TableNodeProvider _tables = new();

    public bool CanHandle(PlanOperation operation) =>
        operation is FillTableOp { Target: NodeAnchor { Kind: "table" } };

    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (FillTableOp)operation;
        var anchor = (NodeAnchor)op.Target;
        if (_tables.Resolve(anchor, new WordObjectMap(context.Package))?.Elements[0] is not Table table)
            return Fail(ValidationErrorCodes.AnchorNotFound, $"No table at path '{anchor.Path}'.", anchor);

        var rows = table.Elements<TableRow>().ToList();
        var columns = rows.FirstOrDefault()?.Elements<TableCell>().Count() ?? 0;
        var ragged = rows.Any(row => row.Elements<TableCell>().Count() != columns);
        if (TableFill.Check(op.Rows, columns, rows.Count > 0, ragged || Merged(table)) is { } error)
            return Fail(ValidationErrorCodes.InvalidOperation, error, anchor);

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "fillTable",
            Before = $"[table {rows.Count}×{columns}]",
            After = TableFill.Describe(op.Rows),
            Context = anchor.Path,
            BlastRadius = op.Rows.Count * columns
        });
    }

    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (FillTableOp)operation;
        var anchor = (NodeAnchor)op.Target;
        var table = (Table)(_tables.Resolve(anchor, new WordObjectMap(context.Package))?.Elements[0]
            ?? throw new InvalidOperationException($"Table '{anchor.Path}' vanished before apply."));

        var rows = table.Elements<TableRow>().ToList();
        while (rows.Count < op.Rows.Count)
        {
            var copy = (TableRow)rows[rows.Count - 1].CloneNode(deep: true);
            InsertTableRowsHandler.ClearParagraphIds(copy);
            InsertTableRowsHandler.ClearRevisionMarkup(copy);
            rows[rows.Count - 1].InsertAfterSelf(copy);
            rows.Add(copy);
        }

        while (rows.Count > op.Rows.Count)
        {
            rows[rows.Count - 1].Remove();
            rows.RemoveAt(rows.Count - 1);
        }

        for (var r = 0; r < rows.Count; r++)
        {
            var cells = rows[r].Elements<TableCell>().ToList();
            for (var c = 0; c < cells.Count; c++)
                SetText(cells[c], op.Rows[r][c]);
        }
    }

    /// <summary>
    /// Writes a cell's text into its first paragraph with that paragraph's first run's
    /// formatting. Further paragraphs go, so the cell holds the value and nothing left over.
    /// </summary>
    private static void SetText(TableCell cell, string text)
    {
        var paragraphs = cell.Elements<Paragraph>().ToList();
        var first = paragraphs.FirstOrDefault() ?? cell.AppendChild(new Paragraph());
        foreach (var extra in paragraphs.Skip(1)) extra.Remove();

        var pattern = WordModel.Dialect.GetRuns(first).OfType<Run>().FirstOrDefault(WordModel.Dialect.IsTextRun)
                      ?? first.Descendants<Run>().FirstOrDefault();
        var properties = pattern?.RunProperties?.CloneNode(deep: true);

        foreach (var child in first.ChildElements.ToList())
            if (child is not ParagraphProperties)
                child.Remove();

        if (text.Length == 0) return;
        var run = new Run();
        if (properties is not null) run.AppendChild(properties);
        run.AppendChild(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        first.AppendChild(run);
    }

    private static bool Merged(Table table) =>
        table.Descendants<TableCellProperties>().Any(properties =>
            (properties.GridSpan?.Val?.Value ?? 1) > 1 || properties.VerticalMerge is not null || properties.HorizontalMerge is not null);

    private static OperationPreview Fail(string code, string message, Anchor anchor) =>
        OperationPreview.Fail(new ValidationError(code, message, anchor));
}

/// <summary>
/// Links a Word text span or table to a workbook range: the span is wrapped in a content
/// control tagged with the marker, and a table gets the marker as its alt-text title.
/// </summary>
internal sealed class BindDataHandler : IOperationHandler
{
    private readonly TableNodeProvider _tables = new();

    public bool CanHandle(PlanOperation operation) =>
        operation is BindDataOp { Target: TextSpanAnchor or NodeAnchor { Kind: "table" } };

    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (BindDataOp)operation;
        if (!WorkbookReference.TryParse(op.Source, out _))
            return Fail(ValidationErrorCodes.InvalidOperation,
                $"'{op.Source}' is not a workbook range or name. Use a reference such as Summary!B2 or 'Q3 Sales'!A1:D6, " +
                "or a defined name or Excel table name.", op.Target);

        var marker = DataBindingMarker.For(op.Source.Trim());
        if (op.Target is NodeAnchor node)
        {
            if (_tables.Resolve(node, new WordObjectMap(context.Package))?.Elements[0] is not Table table)
                return Fail(ValidationErrorCodes.AnchorNotFound, $"No table at path '{node.Path}'.", node);
            return Ok(node, TableNodeProvider.CaptionOf(table) ?? string.Empty, marker, node.Path);
        }

        var anchor = (TextSpanAnchor)op.Target;
        var paragraph = WordModel.ResolveParagraph(context, anchor.ParaId);
        if (paragraph is null)
            return Fail(ValidationErrorCodes.AnchorNotFound, $"No paragraph with id '{anchor.ParaId}'.", anchor);
        if (string.IsNullOrEmpty(anchor.Expect))
            return Fail(ValidationErrorCodes.InvalidOperation,
                "bindData needs the text to link in 'expect', such as the number the workbook cell supplies.", anchor);

        var text = WordModel.Text.GetLogicalText(paragraph);
        var start = WordModel.Text.IndexOfOccurrence(text, anchor.Expect, anchor.Occurrence, WordModel.Comparison(caseSensitive: true));
        if (start < 0)
            return Fail(ValidationErrorCodes.ExpectMismatch,
                $"Expected text '{anchor.Expect}' not found in paragraph '{anchor.ParaId}' (document drifted).", anchor);

        var covered = ChangeTextHandler.CoveredRuns(paragraph, start, anchor.Expect.Length);
        if (covered.Select(run => run.Parent).Distinct().Count() > 1)
            return Fail(ValidationErrorCodes.InvalidOperation,
                $"'{anchor.Expect}' crosses a hyperlink, content control or tracked change, so it cannot be wrapped in one " +
                "content control. Link a span that lies wholly inside or wholly outside it.", anchor);
        if (covered[0].Parent is SdtContentRun && !WholeControl(covered, anchor.Expect))
            return Fail(ValidationErrorCodes.InvalidOperation,
                $"'{anchor.Expect}' is part of an existing content control. Link the control's whole text, which retags it, " +
                "or text outside it.", anchor);

        return Ok(anchor, anchor.Expect, marker, WordModel.Snippet(text, start, anchor.Expect.Length));
    }

    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (BindDataOp)operation;
        var source = op.Source.Trim();

        if (op.Target is NodeAnchor node)
        {
            var table = (Table)(_tables.Resolve(node, new WordObjectMap(context.Package))?.Elements[0]
                ?? throw new InvalidOperationException($"Table '{node.Path}' vanished before apply."));
            var properties = table.GetFirstChild<TableProperties>() ?? table.PrependChild(new TableProperties());
            properties.TableCaption = new TableCaption { Val = DataBindingMarker.For(source) };
            return;
        }

        var anchor = (TextSpanAnchor)op.Target;
        var paragraph = WordModel.ResolveParagraph(context, anchor.ParaId)
            ?? throw new InvalidOperationException($"Paragraph '{anchor.ParaId}' vanished before apply.");
        var text = WordModel.Text.GetLogicalText(paragraph);
        var start = WordModel.Text.IndexOfOccurrence(text, anchor.Expect, anchor.Occurrence, WordModel.Comparison(caseSensitive: true));
        var covered = WordModel.Text.IsolateSpan(paragraph, start, anchor.Expect.Length);

        // Content-control tags must be unique for fill to find one, so a second link to the
        // same range carries a suffix that the marker reader ignores.
        var tags = new HashSet<string>(AllControls(context.Package).Select(TagOf).OfType<string>(), StringComparer.Ordinal);

        if (covered[0].Parent is SdtContentRun content && content.Parent is SdtRun existing && WholeControl(covered, anchor.Expect))
        {
            var current = TagOf(existing);
            if (DataBindingMarker.TryParse(current, out var bound) && string.Equals(bound, source, StringComparison.OrdinalIgnoreCase))
                return;
            if (current is not null) tags.Remove(current);
            Label(existing.SdtProperties ?? existing.PrependChild(new SdtProperties()), Unique(tags, source), source);
            return;
        }

        var nextId = AllControls(context.Package)
            .Select(sdt => sdt.SdtProperties?.GetFirstChild<SdtId>()?.Val?.Value ?? 0)
            .DefaultIfEmpty(0).Max() + 1;
        var properties2 = new SdtProperties();
        Label(properties2, Unique(tags, source), source);
        properties2.AppendChild(new SdtId { Val = nextId });

        var control = new SdtRun(properties2, new SdtContentRun());
        covered[0].InsertBeforeSelf(control);
        foreach (var run in covered)
        {
            run.Remove();
            control.SdtContentRun!.AppendChild(run);
        }
    }

    /// <summary>Gives a control the marker as its tag and the range as the title Word shows on it.</summary>
    private static void Label(SdtProperties properties, string tag, string source)
    {
        properties.GetFirstChild<SdtAlias>()?.Remove();
        properties.GetFirstChild<Tag>()?.Remove();

        // w:rPr, when a control has one, must stay the first child of w:sdtPr.
        var alias = new SdtAlias { Val = source };
        if (properties.GetFirstChild<RunProperties>() is { } runProperties) runProperties.InsertAfterSelf(alias);
        else properties.PrependChild(alias);
        alias.InsertAfterSelf(new Tag { Val = tag });
    }

    private static string Unique(ISet<string> tags, string source)
    {
        var marker = DataBindingMarker.For(source);
        var tag = marker;
        for (var n = 2; tags.Contains(tag); n++) tag = $"{marker}#{n}";
        return tag;
    }

    private static bool WholeControl(IReadOnlyList<OpenXmlElement> covered, string expect) =>
        covered[0].Parent is SdtContentRun content &&
        string.Equals(string.Concat(content.Descendants<Text>().Select(t => t.Text)), expect, StringComparison.Ordinal);

    private static IEnumerable<SdtElement> AllControls(IOpenXmlPackage package) =>
        WordModel.TextHosts(package).SelectMany(host => host.Root.Descendants<SdtElement>());

    private static string? TagOf(SdtElement sdt) => sdt.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value;

    private static OperationPreview Ok(Anchor target, string before, string after, string context) =>
        OperationPreview.Ok(new ProposedChange
        {
            Target = target,
            Verb = "bindData",
            Before = before,
            After = after,
            Context = context,
            BlastRadius = 1
        });

    private static OperationPreview Fail(string code, string message, Anchor anchor) =>
        OperationPreview.Fail(new ValidationError(code, message, anchor));
}
