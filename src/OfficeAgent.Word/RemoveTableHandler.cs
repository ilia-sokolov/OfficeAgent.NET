using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.Word;

/// <summary>
/// Removes an entire table addressed by a <see cref="NodeAnchor"/> with
/// <c>Kind="table"</c> and <c>Path="table#N"</c>. The table element and every row it
/// contains are deleted; to drop only some rows or columns use the row/column verbs.
/// </summary>
internal sealed class RemoveTableHandler : IOperationHandler
{
    private readonly TableNodeProvider _tables = new();

    public bool CanHandle(PlanOperation operation) =>
        operation is RemoveTableOp { Target: NodeAnchor { Kind: "table" } };

    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (RemoveTableOp)operation;
        var anchor = (NodeAnchor)op.Target;
        var node = _tables.Resolve(anchor, new WordObjectMap(context.Package));
        if (node is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No table at path '{anchor.Path}'.", anchor));

        var table = (Table)node.Elements[0];
        var rowCount = table.Elements<TableRow>().Count();

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "removeTable",
            Before = $"(table, {rowCount} row(s))",
            After = "(removed)",
            Context = anchor.Path,
            BlastRadius = 1
        });
    }

    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (RemoveTableOp)operation;
        var anchor = (NodeAnchor)op.Target;
        var node = _tables.Resolve(anchor, new WordObjectMap(context.Package))
            ?? throw new InvalidOperationException($"Table '{anchor.Path}' vanished before apply.");

        var table = (Table)node.Elements[0];
        table.Remove();
    }
}
