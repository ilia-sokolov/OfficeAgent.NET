using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.Word;

/// <summary>Inserts a new table built from <see cref="TableData"/> relative to an anchored paragraph.</summary>
internal sealed class InsertTableHandler : IOperationHandler
{
    private readonly TableBinder _tables = new();

    public bool CanHandle(PlanOperation operation) =>
        operation is InsertTableOp { Target: TextSpanAnchor };

    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (InsertTableOp)operation;
        var anchor = (TextSpanAnchor)op.Target;

        var paragraph = WordModel.ResolveParagraph(context, anchor.ParaId);
        if (paragraph is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No paragraph with id '{anchor.ParaId}'.", anchor));

        var table = op.Table;
        var rowCount = table.Rows.Count + (table.Headers.Count > 0 ? 1 : 0);
        var colCount = Math.Max(table.Headers.Count, table.Rows.FirstOrDefault()?.Count ?? 0);

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "insertTable",
            Before = string.Empty,
            After = $"[table {rowCount}×{colCount}]",
            Context = $"{op.Position.ToString().ToLowerInvariant()} paragraph '{anchor.ParaId}'",
            BlastRadius = 1
        });
    }

    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (InsertTableOp)operation;
        var anchor = (TextSpanAnchor)op.Target;

        var paragraph = WordModel.ResolveParagraph(context, anchor.ParaId)
            ?? throw new InvalidOperationException($"Paragraph '{anchor.ParaId}' vanished before apply.");

        var table = _tables.Build(op.Table);

        if (op.Position == InsertPosition.Before)
            paragraph.InsertBeforeSelf(table);
        else
            paragraph.InsertAfterSelf(table);
    }
}
