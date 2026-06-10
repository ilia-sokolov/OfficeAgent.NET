using DocumentFormat.OpenXml;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.Word;

/// <summary>
/// Inserts one or more columns into an existing table addressed by a
/// <see cref="NodeAnchor"/> with <c>Kind="table"</c>. Each entry in
/// <see cref="InsertTableColumnsOp.Columns"/> is a column-major list: one cell
/// text per row, header first. Shorter columns are padded with empty strings.
/// </summary>
internal sealed class InsertTableColumnsHandler : IOperationHandler
{
    private readonly TableNodeProvider _tables = new();

    public bool CanHandle(PlanOperation operation) =>
        operation is InsertTableColumnsOp { Target: NodeAnchor { Kind: "table" } };

    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (InsertTableColumnsOp)operation;
        var anchor = (NodeAnchor)op.Target;
        var node = _tables.Resolve(anchor, new WordObjectMap(context.Package));
        if (node is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No table at path '{anchor.Path}'.", anchor));

        if (op.Columns.Count == 0)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "insertTableColumns requires at least one column.", anchor));

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "insertTableColumns",
            Before = $"position={op.Position} columnIndex={op.ColumnIndex}",
            After = $"+{op.Columns.Count} column(s)",
            Context = anchor.Path,
            BlastRadius = op.Columns.Count
        });
    }

    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (InsertTableColumnsOp)operation;
        var anchor = (NodeAnchor)op.Target;
        var node = _tables.Resolve(anchor, new WordObjectMap(context.Package))
            ?? throw new InvalidOperationException($"Table '{anchor.Path}' vanished before apply.");
        var table = (Table)node.Elements[0];

        var rows = table.Elements<TableRow>().ToList();
        if (rows.Count == 0) return;

        // Compute insertion column index per row.
        int totalColumns = rows[0].Elements<TableCell>().Count();
        int resolved = op.Position switch
        {
            TablePosition.End => totalColumns,
            TablePosition.Start => 0,
            TablePosition.Before => Clamp(op.ColumnIndex < 0 ? totalColumns + op.ColumnIndex : op.ColumnIndex, 0, totalColumns),
            TablePosition.After => Clamp((op.ColumnIndex < 0 ? totalColumns + op.ColumnIndex : op.ColumnIndex) + 1, 0, totalColumns),
            _ => totalColumns
        };

        static int Clamp(int value, int min, int max) => Math.Min(Math.Max(value, min), max);

        for (int r = 0; r < rows.Count; r++)
        {
            var row = rows[r];
            var existingCells = row.Elements<TableCell>().ToList();

            for (int c = 0; c < op.Columns.Count; c++)
            {
                var text = r < op.Columns[c].Count ? op.Columns[c][r] : string.Empty;
                var newCell = NewCell(text);
                int columnSlot = resolved + c;

                if (columnSlot >= existingCells.Count)
                {
                    row.AppendChild(newCell);
                }
                else
                {
                    existingCells[columnSlot].InsertBeforeSelf(newCell);
                }
                existingCells = row.Elements<TableCell>().ToList();
            }
        }
    }

    private static TableCell NewCell(string text) =>
        new TableCell(new Paragraph(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve })));
}
