using DocumentFormat.OpenXml;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.Word;

/// <summary>Inserts a new paragraph or a structured table relative to an anchored paragraph.</summary>
internal sealed class InsertHandler : IOperationHandler
{
    private readonly TableBinder _tables = new();

    public bool CanHandle(PlanOperation operation) =>
        operation is InsertOp { Target: TextSpanAnchor };

    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (InsertOp)operation;
        var anchor = (TextSpanAnchor)op.Target;

        var paragraph = WordModel.ResolveParagraph(context, anchor.ParaId);
        if (paragraph is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No paragraph with id '{anchor.ParaId}'.", anchor));

        string after = op.Table is { } table
            ? $"[table {table.Rows.Count + (table.Headers.Count > 0 ? 1 : 0)}×{Math.Max(table.Headers.Count, table.Rows.FirstOrDefault()?.Count ?? 0)}]"
            : op.Text ?? string.Empty;

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "insert",
            Before = string.Empty,
            After = after,
            Context = $"{op.Position.ToString().ToLowerInvariant()} paragraph '{anchor.ParaId}'",
            BlastRadius = 1
        });
    }

    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (InsertOp)operation;
        var anchor = (TextSpanAnchor)op.Target;

        var paragraph = WordModel.ResolveParagraph(context, anchor.ParaId)
            ?? throw new InvalidOperationException($"Paragraph '{anchor.ParaId}' vanished before apply.");

        OpenXmlElement element = op.Table is { } table
            ? _tables.Build(table)
            : BuildParagraph(op);

        if (op.Position == InsertPosition.Before)
            paragraph.InsertBeforeSelf(element);
        else
            paragraph.InsertAfterSelf(element);
    }

    private static Paragraph BuildParagraph(InsertOp op)
    {
        var paragraph = new Paragraph();
        if (!string.IsNullOrEmpty(op.StyleId))
            paragraph.ParagraphProperties = new ParagraphProperties(new ParagraphStyleId { Val = op.StyleId });

        paragraph.AppendChild(new Run(new Text(op.Text ?? string.Empty) { Space = SpaceProcessingModeValues.Preserve }));
        return paragraph;
    }
}
