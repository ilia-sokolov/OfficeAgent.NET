using DocumentFormat.OpenXml;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.Word;

/// <summary>
/// Copies direct formatting (and the assigned paragraph style id, when copying
/// paragraph scope) from a source <see cref="TextSpanAnchor"/> to the destination
/// <see cref="TextSpanAnchor"/> in <see cref="PlanOperation.Target"/>. When the
/// span anchor's <c>Expect</c> is empty, the entire paragraph is treated as the
/// span. Scope is one of <c>run</c>, <c>paragraph</c>, or <c>all</c>.
/// </summary>
internal sealed class CopyStylesHandler : IOperationHandler
{
    public bool CanHandle(PlanOperation operation) =>
        operation is CopyStylesOp { Target: TextSpanAnchor, Source: TextSpanAnchor };

    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (CopyStylesOp)operation;
        var src = (TextSpanAnchor)op.Source;
        var dst = (TextSpanAnchor)op.Target;

        if (!IsValidScope(op.Scope))
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"copyStyles scope must be 'run', 'paragraph', or 'all'; got '{op.Scope}'.", dst));

        if (WordModel.ResolveParagraph(context, src.ParaId) is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"copyStyles source paragraph '{src.ParaId}' was not found.", src));

        if (WordModel.ResolveParagraph(context, dst.ParaId) is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"copyStyles target paragraph '{dst.ParaId}' was not found.", dst));

        return OperationPreview.Ok(new ProposedChange
        {
            Target = dst,
            Verb = "copyStyles",
            Before = $"source={src.ParaId}",
            After = $"scope={op.Scope}",
            Context = $"{src.ParaId} → {dst.ParaId}",
            BlastRadius = 1
        });
    }

    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (CopyStylesOp)operation;
        var src = (TextSpanAnchor)op.Source;
        var dst = (TextSpanAnchor)op.Target;

        var sourceParagraph = WordModel.ResolveParagraph(context, src.ParaId)!;
        var targetParagraph = WordModel.ResolveParagraph(context, dst.ParaId)!;

        var copyRun = op.Scope is "run" or "all";
        var copyPara = op.Scope is "paragraph" or "all";

        if (copyPara)
            CopyParagraphProperties(sourceParagraph, targetParagraph);

        if (copyRun)
        {
            var sourceRun = FirstRunOf(sourceParagraph, src);
            if (sourceRun is null) return;

            var targetRuns = TargetRuns(targetParagraph, dst);
            foreach (var run in targetRuns)
                ReplaceRunProperties(run, sourceRun);
        }
    }

    private static bool IsValidScope(string s) => s is "run" or "paragraph" or "all";

    private static void CopyParagraphProperties(Paragraph source, Paragraph target)
    {
        var existing = source.ParagraphProperties;
        if (existing is null)
        {
            target.ParagraphProperties = null;
            return;
        }
        var clone = (ParagraphProperties)existing.CloneNode(deep: true);
        target.ParagraphProperties = clone;
    }

    private static void ReplaceRunProperties(Run target, Run source)
    {
        var existing = source.RunProperties;
        if (existing is null)
        {
            target.RunProperties = null;
            return;
        }
        target.RunProperties = (RunProperties)existing.CloneNode(deep: true);
    }

    private static Run? FirstRunOf(Paragraph paragraph, TextSpanAnchor anchor)
    {
        if (string.IsNullOrEmpty(anchor.Expect))
            return paragraph.Elements<Run>().FirstOrDefault(r => r.Elements<Text>().Any());

        var text = WordModel.Text.GetLogicalText(paragraph);
        var comparison = WordModel.Comparison(caseSensitive: true);
        int start = WordModel.Text.IndexOfOccurrence(text, anchor.Expect, anchor.Occurrence, comparison);
        if (start < 0) return null;
        var covered = WordModel.Text.IsolateSpan(paragraph, start, anchor.Expect.Length);
        return covered.OfType<Run>().FirstOrDefault();
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
