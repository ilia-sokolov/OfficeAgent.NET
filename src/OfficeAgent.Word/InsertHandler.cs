using DocumentFormat.OpenXml;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.Word;

/// <summary>Inserts a new paragraph relative to an anchored paragraph. Tables are inserted by <see cref="InsertTableHandler"/>.</summary>
internal sealed class InsertHandler : IOperationHandler
{
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

        // level is DrawingML's bullet depth and has no WordprocessingML counterpart -
        // numbering here comes from the paragraph's style and numbering definition. Losing
        // it silently would produce a list that looks flat for no visible reason.
        if (op.Level is not null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "The Word module cannot apply 'level'; it is a slide's bullet depth. " +
                "Use styleId with a list style instead.", anchor));

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "insert",
            Before = string.Empty,
            After = op.Text ?? string.Empty,
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

        var element = BuildParagraph(op);

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
