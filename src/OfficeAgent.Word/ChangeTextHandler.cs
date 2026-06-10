using DocumentFormat.OpenXml;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.Word;

/// <summary>
/// Replaces an anchored, content-verified text span. Direct mode rewrites the
/// runs; tracked mode lands the edit as a Word redline (w:del + w:ins). Handles
/// run-spanning text via the Core <see cref="TextBodyEngine"/>.
/// </summary>
internal sealed class ChangeTextHandler : IOperationHandler
{
    private readonly TimeProvider _clock;

    public ChangeTextHandler(TimeProvider clock) => _clock = clock;

    public bool CanHandle(PlanOperation operation) =>
        operation is ChangeTextOp { Target: TextSpanAnchor };

    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (ChangeTextOp)operation;
        var anchor = (TextSpanAnchor)op.Target;

        if (string.IsNullOrEmpty(anchor.Expect))
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "ChangeTextOp requires a non-empty 'expect' value identifying the text to replace. " +
                "To remove an entire paragraph, set 'with' to the empty string and 'expect' to the current paragraph text.",
                anchor));

        var paragraph = WordModel.ResolveParagraph(context, anchor.ParaId);
        if (paragraph is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No paragraph with id '{anchor.ParaId}'.", anchor));

        var text = WordModel.Text.GetLogicalText(paragraph);
        var comparison = WordModel.Comparison(caseSensitive: true);
        int occurrences = WordModel.Text.CountOccurrences(text, anchor.Expect, comparison);

        if (occurrences == 0)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.ExpectMismatch,
                $"Expected text '{anchor.Expect}' not found in paragraph '{anchor.ParaId}' (document drifted).", anchor));

        int start = WordModel.Text.IndexOfOccurrence(text, anchor.Expect, anchor.Occurrence, comparison);
        if (start < 0)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AmbiguousAnchor,
                $"Occurrence {anchor.Occurrence} of '{anchor.Expect}' does not exist ({occurrences} found).", anchor));

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "changeText",
            Before = anchor.Expect,
            After = op.With,
            Context = WordModel.Snippet(text, start, anchor.Expect.Length),
            BlastRadius = 1
        });
    }

    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (ChangeTextOp)operation;
        var anchor = (TextSpanAnchor)op.Target;

        var paragraph = WordModel.ResolveParagraph(context, anchor.ParaId)
            ?? throw new InvalidOperationException($"Paragraph '{anchor.ParaId}' vanished before apply.");

        var text = WordModel.Text.GetLogicalText(paragraph);
        var comparison = WordModel.Comparison(caseSensitive: true);
        int start = WordModel.Text.IndexOfOccurrence(text, anchor.Expect, anchor.Occurrence, comparison);
        if (start < 0)
            throw new InvalidOperationException($"Expected text '{anchor.Expect}' not found at apply time.");

        var covered = WordModel.Text.IsolateSpan(paragraph, start, anchor.Expect.Length);
        if (covered.Count == 0)
            throw new InvalidOperationException("Span isolation produced no runs.");

        if (op.Mode == ChangeMode.Direct)
            ApplyDirect(covered, op.With);
        else
            ApplyTracked(context.Package, covered, op.With);
    }

    private static void ApplyDirect(IReadOnlyList<OpenXmlElement> covered, string replacement)
    {
        WordModel.Dialect.SetRunText(covered[0], replacement);
        for (int i = 1; i < covered.Count; i++)
            covered[i].Remove();
    }

    private void ApplyTracked(IOpenXmlPackage package, IReadOnlyList<OpenXmlElement> covered, string replacement)
    {
        var first = (Run)covered[0];
        var parent = first.Parent
            ?? throw new InvalidOperationException("Run has no parent paragraph.");

        var allocator = new WordRevisionIdAllocator(package);
        var author = "OfficeAgent";
        var stamp = _clock.GetUtcNow().UtcDateTime;

        var deleted = new DeletedRun
        {
            Author = author,
            Date = stamp,
            Id = allocator.Next().ToString()
        };
        foreach (var element in covered)
        {
            var clone = (Run)element.CloneNode(deep: true);
            foreach (var t in clone.Elements<Text>().ToList())
            {
                var delText = new DeletedText(t.Text) { Space = SpaceProcessingModeValues.Preserve };
                t.InsertAfterSelf(delText);
                t.Remove();
            }
            deleted.AppendChild(clone);
        }

        var insertRun = (Run)first.CloneNode(deep: true);
        WordModel.Dialect.SetRunText(insertRun, replacement);
        var inserted = new InsertedRun
        {
            Author = author,
            Date = stamp,
            Id = allocator.Next().ToString()
        };
        inserted.AppendChild(insertRun);

        parent.InsertBefore(deleted, first);
        parent.InsertBefore(inserted, first);

        foreach (var element in covered)
            element.Remove();
    }
}
