using DocumentFormat.OpenXml;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.Word;

/// <summary>
/// Removes rows from an existing table addressed by a <see cref="NodeAnchor"/>
/// with <c>Kind="table"</c> and <c>Path="table#N"</c>. Negative indices count from
/// the end (-1 = last). When <see cref="RemoveTableRowsOp.OnlyIfEmpty"/> is true,
/// the handler keeps any non-blank row even if its index was listed - which is the
/// safe behaviour when the LLM asks to "clean up empty rows".
/// </summary>
internal sealed class RemoveTableRowsHandler : IOperationHandler
{
    private readonly TableNodeProvider _tables = new();

    public bool CanHandle(PlanOperation operation) =>
        operation is RemoveTableRowsOp { Target: NodeAnchor { Kind: "table" } };

    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (RemoveTableRowsOp)operation;
        var anchor = (NodeAnchor)op.Target;
        var node = _tables.Resolve(anchor, new WordObjectMap(context.Package));
        if (node is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No table at path '{anchor.Path}'.", anchor));

        var table = (Table)node.Elements[0];
        var rows = table.Elements<TableRow>().ToList();
        var toRemove = SelectRowsToRemove(op, rows);

        if (toRemove.Count == 0)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                op.OnlyIfEmpty
                    ? $"No empty rows to remove in '{anchor.Path}'."
                    : $"RemoveTableRowsOp resolved zero rows to remove in '{anchor.Path}'. Provide explicit RowIndices or set OnlyIfEmpty=true.",
                anchor));

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "removeTableRows",
            Before = $"{rows.Count} row(s)",
            After = $"{rows.Count - toRemove.Count} row(s)",
            Context = anchor.Path,
            BlastRadius = toRemove.Count
        });
    }

    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (RemoveTableRowsOp)operation;
        var anchor = (NodeAnchor)op.Target;
        var node = _tables.Resolve(anchor, new WordObjectMap(context.Package))
            ?? throw new InvalidOperationException($"Table '{anchor.Path}' vanished before apply.");
        var table = (Table)node.Elements[0];
        var rows = table.Elements<TableRow>().ToList();
        var toRemove = SelectRowsToRemove(op, rows);

        // Highest index first so prior removals don't shift later indices.
        foreach (var index in toRemove.OrderByDescending(i => i))
            rows[index].Remove();
    }

    private static List<int> SelectRowsToRemove(RemoveTableRowsOp op, List<TableRow> rows)
    {
        var resolved = new HashSet<int>();
        if (op.RowIndices.Count > 0)
        {
            foreach (var raw in op.RowIndices)
            {
                var i = raw < 0 ? rows.Count + raw : raw;
                if (i < 0 || i >= rows.Count) continue;
                if (op.OnlyIfEmpty && !IsRowEmpty(rows[i])) continue;
                resolved.Add(i);
            }
        }
        else if (op.OnlyIfEmpty)
        {
            for (int i = 0; i < rows.Count; i++)
                if (IsRowEmpty(rows[i])) resolved.Add(i);
        }
        return resolved.ToList();
    }

    private static bool IsRowEmpty(TableRow row)
    {
        foreach (var text in row.Descendants<Text>())
            if (!string.IsNullOrWhiteSpace(text.Text))
                return false;
        return true;
    }
}
