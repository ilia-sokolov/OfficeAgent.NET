using DocumentFormat.OpenXml;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.Word;

/// <summary>
/// Inserts one or more rows into an existing table addressed by a
/// <see cref="NodeAnchor"/> with <c>Kind="table"</c> and <c>Path="table#N"</c>.
/// Placement is controlled by <see cref="InsertTableRowsOp.Position"/> +
/// <see cref="InsertTableRowsOp.RowIndex"/>. The handler clones the last
/// existing row as a template so column widths and cell formatting are preserved.
/// </summary>
internal sealed class InsertTableRowsHandler : IOperationHandler
{
    private readonly TableNodeProvider _tables = new();

    public bool CanHandle(PlanOperation operation) =>
        operation is InsertTableRowsOp { Target: NodeAnchor { Kind: "table" } };

    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (InsertTableRowsOp)operation;
        var anchor = (NodeAnchor)op.Target;
        var node = _tables.Resolve(anchor, new WordObjectMap(context.Package));
        if (node is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No table at path '{anchor.Path}'.", anchor));

        if (op.Rows.Count == 0)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "insertTableRows requires at least one row.", anchor));

        var preview = string.Join(" / ", op.Rows.Take(2).Select(r => string.Join(" | ", r)));
        if (op.Rows.Count > 2) preview += $" / (+{op.Rows.Count - 2} more)";

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "insertTableRows",
            Before = $"position={op.Position} rowIndex={op.RowIndex}",
            After = preview,
            Context = anchor.Path,
            BlastRadius = op.Rows.Count
        });
    }

    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (InsertTableRowsOp)operation;
        var anchor = (NodeAnchor)op.Target;
        var node = _tables.Resolve(anchor, new WordObjectMap(context.Package))
            ?? throw new InvalidOperationException($"Table '{anchor.Path}' vanished before apply.");
        var table = (Table)node.Elements[0];

        var existing = table.Elements<TableRow>().ToList();
        var templateRow = existing.LastOrDefault();
        var newRows = op.Rows.Select(cells => BuildRow(templateRow, cells)).ToList();

        InsertAt(table, existing, newRows, op.Position, op.RowIndex);
    }

    internal static void InsertAt(
        OpenXmlElement parent,
        IReadOnlyList<TableRow> existingRows,
        IReadOnlyList<TableRow> newRows,
        TablePosition position,
        int index)
    {
        TableRow? pivot = null;
        bool insertBefore = false;

        switch (position)
        {
            case TablePosition.End:
                foreach (var row in newRows) parent.AppendChild(row);
                return;

            case TablePosition.Start:
                pivot = existingRows.FirstOrDefault();
                insertBefore = true;
                break;

            case TablePosition.Before:
            case TablePosition.After:
                if (existingRows.Count == 0)
                {
                    foreach (var row in newRows) parent.AppendChild(row);
                    return;
                }
                var raw = index < 0 ? existingRows.Count + index : index;
                var resolvedIndex = Math.Min(Math.Max(raw, 0), existingRows.Count - 1);
                pivot = existingRows[resolvedIndex];
                insertBefore = position == TablePosition.Before;
                break;
        }

        if (pivot is null)
        {
            foreach (var row in newRows) parent.AppendChild(row);
            return;
        }

        // Insert in order so they keep their relative ordering in the final document.
        if (insertBefore)
        {
            foreach (var row in newRows) pivot.InsertBeforeSelf(row);
        }
        else
        {
            // Walk forward so each inserted row becomes the new pivot.
            OpenXmlElement cursor = pivot;
            foreach (var row in newRows)
            {
                cursor.InsertAfterSelf(row);
                cursor = row;
            }
        }
    }

    private static TableRow BuildRow(TableRow? templateRow, IReadOnlyList<string> cells)
    {
        if (templateRow is null)
        {
            var row = new TableRow();
            foreach (var cell in cells)
                row.AppendChild(NewCell(cell));
            return row;
        }

        var clone = (TableRow)templateRow.CloneNode(deep: true);
        var templateCells = clone.Elements<TableCell>().ToList();
        int i = 0;
        foreach (var cell in templateCells)
        {
            ReplaceCellText(cell, i < cells.Count ? cells[i] : string.Empty);
            i++;
        }
        for (; i < cells.Count; i++)
            clone.AppendChild(NewCell(cells[i]));
        return clone;
    }

    private static void ReplaceCellText(TableCell cell, string text)
    {
        var paragraph = cell.Elements<Paragraph>().FirstOrDefault();
        if (paragraph is null)
        {
            cell.AppendChild(new Paragraph(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve })));
            return;
        }

        foreach (var run in paragraph.Elements<Run>().ToList())
            run.Remove();
        paragraph.AppendChild(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
    }

    private static TableCell NewCell(string text) =>
        new TableCell(new Paragraph(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve })));
}
