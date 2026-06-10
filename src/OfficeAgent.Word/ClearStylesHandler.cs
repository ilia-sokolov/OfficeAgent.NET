using DocumentFormat.OpenXml;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.Word;

/// <summary>
/// Removes direct formatting from the target span. With <c>scope="run"</c> only
/// <c>w:rPr</c> children of the matched runs are removed; <c>scope="paragraph"</c>
/// clears the paragraph's <c>w:pPr</c> (style assignment and all paragraph-level
/// properties); <c>scope="all"</c> does both. Empty <c>Expect</c> means the whole
/// paragraph.
/// </summary>
internal sealed class ClearStylesHandler : IOperationHandler
{
    public bool CanHandle(PlanOperation operation) =>
        operation is ClearStylesOp { Target: TextSpanAnchor };

    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (ClearStylesOp)operation;
        var anchor = (TextSpanAnchor)op.Target;

        if (op.Scope is not ("run" or "paragraph" or "all"))
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"clearStyles scope must be 'run', 'paragraph', or 'all'; got '{op.Scope}'.", anchor));

        if (WordModel.ResolveParagraph(context, anchor.ParaId) is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No paragraph with id '{anchor.ParaId}'.", anchor));

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "clearStyles",
            Before = "(direct formatting)",
            After = "(cleared)",
            Context = anchor.ParaId,
            BlastRadius = 1
        });
    }

    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (ClearStylesOp)operation;
        var anchor = (TextSpanAnchor)op.Target;
        var paragraph = WordModel.ResolveParagraph(context, anchor.ParaId)
            ?? throw new InvalidOperationException($"Paragraph '{anchor.ParaId}' vanished before apply.");

        if (op.Scope is "paragraph" or "all")
            paragraph.ParagraphProperties = null;

        if (op.Scope is "run" or "all")
        {
            var runs = TargetRuns(paragraph, anchor);
            foreach (var run in runs)
                run.RunProperties = null;
        }
    }

    private static IReadOnlyList<Run> TargetRuns(Paragraph paragraph, TextSpanAnchor anchor)
    {
        if (string.IsNullOrEmpty(anchor.Expect))
            return paragraph.Elements<Run>().ToList();

        var text = WordModel.Text.GetLogicalText(paragraph);
        int start = WordModel.Text.IndexOfOccurrence(text, anchor.Expect, anchor.Occurrence, WordModel.Comparison(true));
        if (start < 0) return Array.Empty<Run>();
        return WordModel.Text.IsolateSpan(paragraph, start, anchor.Expect.Length).OfType<Run>().ToList();
    }
}
