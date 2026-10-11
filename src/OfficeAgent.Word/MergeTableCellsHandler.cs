using DocumentFormat.OpenXml;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.Word;

/// <summary>
/// Merges a table cell with the cells to its right in the same row (ECMA-376 Part 1,
/// 17.4.17): the target takes a <c>w:gridSpan</c> over the combined grid columns, the
/// absorbed cells are removed, and their content is appended to the target, as Word does.
/// The table grid is unchanged.
/// </summary>
internal sealed class MergeTableCellsHandler : IOperationHandler
{
    public bool CanHandle(PlanOperation operation) =>
        operation is MergeTableCellsOp { Target: NodeAnchor { Kind: "tableCell" } };

    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (MergeTableCellsOp)operation;
        var anchor = (NodeAnchor)op.Target;

        if (op.Mode == ChangeMode.Tracked)
            return Fail(anchor, ValidationErrorCodes.InvalidOperation,
                "A cell merge is not recorded as a tracked change here, so mode 'Tracked' cannot be honoured. " +
                "Send this operation with mode 'Direct'.");

        var cell = TableLocator.FindCell(context.Package, anchor.Path);
        if (cell is null)
            return Fail(anchor, ValidationErrorCodes.AnchorNotFound,
                $"No table cell at '{anchor.Path}'. Cell paths are table#N/cell#R/C, from inspect_document.");

        var plan = Plan(cell, op.ColumnSpan, out var error);
        if (plan is null)
            return Fail(anchor, ValidationErrorCodes.InvalidOperation, error!);

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "mergeTableCells",
            Before = $"{plan.Count + 1} cells",
            After = $"1 cell spanning {op.ColumnSpan} columns",
            Context = anchor.Path,
            BlastRadius = plan.Count + 1,
            Capability = Capability.Deterministic
        });
    }

    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (MergeTableCellsOp)operation;
        var anchor = (NodeAnchor)op.Target;
        var cell = TableLocator.FindCell(context.Package, anchor.Path)
            ?? throw new InvalidOperationException($"Cell '{anchor.Path}' vanished before apply.");
        var absorbed = Plan(cell, op.ColumnSpan, out var error)
            ?? throw new InvalidOperationException(error);

        var properties = cell.TableCellProperties ?? cell.PrependChild(new TableCellProperties());
        properties.GridSpan = new GridSpan { Val = op.ColumnSpan };

        // A width the cells all state in twips adds up; any other mix is left for Word to lay out.
        var widths = new[] { cell }.Concat(absorbed).Select(c => c.TableCellProperties?.TableCellWidth).ToList();
        if (widths.All(w => w?.Type?.Value == TableWidthUnitValues.Dxa && int.TryParse(w.Width?.Value, out _)))
            properties.TableCellWidth = new TableCellWidth
            {
                Type = TableWidthUnitValues.Dxa,
                Width = widths.Sum(w => int.Parse(w!.Width!.Value!)).ToString(System.Globalization.CultureInfo.InvariantCulture)
            };

        foreach (var other in absorbed)
        {
            var content = other.ChildElements.Where(child => child is not TableCellProperties).ToList();
            if (!IsEmpty(content))
                foreach (var block in content)
                {
                    block.Remove();
                    cell.AppendChild(block);
                }
            other.Remove();
        }
    }

    /// <summary>
    /// The cells the target absorbs to span <paramref name="columnSpan"/> grid columns, or null
    /// with the reason when the merge is not possible.
    /// </summary>
    private static List<TableCell>? Plan(TableCell cell, int columnSpan, out string? error)
    {
        error = null;
        var span = Span(cell);
        if (columnSpan <= span)
        {
            error = $"columnSpan {columnSpan} must be greater than the {span} column(s) the cell already spans.";
            return null;
        }
        if (cell.TableCellProperties?.VerticalMerge is not null)
        {
            error = "The cell is part of a vertical merge; merging it horizontally is not supported.";
            return null;
        }

        var absorbed = new List<TableCell>();
        foreach (var next in cell.ElementsAfter().OfType<TableCell>())
        {
            if (span == columnSpan) break;
            if (next.TableCellProperties?.VerticalMerge is not null)
            {
                error = "A cell the merge would absorb is part of a vertical merge; that is not supported.";
                return null;
            }
            span += Span(next);
            absorbed.Add(next);
        }

        if (span < columnSpan)
        {
            error = $"The row has only {span} grid column(s) from this cell to its end, fewer than columnSpan {columnSpan}.";
            return null;
        }
        if (span > columnSpan)
        {
            error = $"Spanning exactly {columnSpan} columns would split a cell that already spans several; " +
                    $"use columnSpan {span} to include it whole.";
            return null;
        }
        return absorbed;
    }

    private static int Span(TableCell cell) => cell.TableCellProperties?.GridSpan?.Val?.Value ?? 1;

    /// <summary>A cell holding only an empty paragraph contributes nothing when absorbed.</summary>
    private static bool IsEmpty(IReadOnlyList<OpenXmlElement> content) =>
        content.Count == 1 && content[0] is Paragraph paragraph &&
        string.IsNullOrEmpty(WordModel.Text.GetLogicalText(paragraph)) &&
        !paragraph.Descendants().Any(e => e is Drawing or EmbeddedObject or FieldChar or SimpleField);

    private static OperationPreview Fail(Anchor anchor, string code, string message) =>
        OperationPreview.Fail(new ValidationError(code, message, anchor));
}
