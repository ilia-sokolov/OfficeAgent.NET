using DocumentFormat.OpenXml;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.Word;

/// <summary>
/// Removes one or more columns from an existing table addressed by a
/// <see cref="NodeAnchor"/> with <c>Kind="table"</c>. Indices are zero-based;
/// negative values count from the right (-1 = last column).
/// </summary>
internal sealed class RemoveTableColumnsHandler : IOperationHandler
{
    private readonly TableNodeProvider _tables = new();

    public bool CanHandle(PlanOperation operation) =>
        operation is RemoveTableColumnsOp { Target: NodeAnchor { Kind: "table" } };

    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (RemoveTableColumnsOp)operation;
        var anchor = (NodeAnchor)op.Target;
        var node = _tables.Resolve(anchor, new WordObjectMap(context.Package));
        if (node is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No table at path '{anchor.Path}'.", anchor));

        if (op.ColumnIndices.Count == 0)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "removeTableColumns requires at least one ColumnIndices entry.", anchor));

        var table = (Table)node.Elements[0];
        var totalColumns = table.Elements<TableRow>().FirstOrDefault()?.Elements<TableCell>().Count() ?? 0;
        var resolved = ResolveIndices(op.ColumnIndices, totalColumns);

        if (resolved.Count == 0)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"None of the requested column indices fall within the table's {totalColumns} column(s).",
                anchor));

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "removeTableColumns",
            Before = $"{totalColumns} column(s)",
            After = $"{totalColumns - resolved.Count} column(s)",
            Context = anchor.Path,
            BlastRadius = resolved.Count
        });
    }

    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (RemoveTableColumnsOp)operation;
        var anchor = (NodeAnchor)op.Target;
        var node = _tables.Resolve(anchor, new WordObjectMap(context.Package))
            ?? throw new InvalidOperationException($"Table '{anchor.Path}' vanished before apply.");
        var table = (Table)node.Elements[0];
        var rows = table.Elements<TableRow>().ToList();
        var totalColumns = rows[0].Elements<TableCell>().Count();
        var resolved = ResolveIndices(op.ColumnIndices, totalColumns);

        // Highest index first so prior removals don't shift later ones.
        foreach (var row in rows)
        {
            var cells = row.Elements<TableCell>().ToList();
            foreach (var index in resolved.OrderByDescending(i => i))
                if (index < cells.Count) cells[index].Remove();
        }
    }

    private static List<int> ResolveIndices(IReadOnlyList<int> requested, int totalColumns)
    {
        var set = new HashSet<int>();
        foreach (var raw in requested)
        {
            var i = raw < 0 ? totalColumns + raw : raw;
            if (i >= 0 && i < totalColumns) set.Add(i);
        }
        return set.ToList();
    }
}
