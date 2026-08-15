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

        var paragraph = WordModel.ResolveParagraph(context, anchor.ParaId);
        if (paragraph is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No paragraph with id '{anchor.ParaId}'.", anchor));

        var text = WordModel.Text.GetLogicalText(paragraph);

        // An empty 'expect' names no text, so against a paragraph that has some there is no
        // way to tell "replace all of it" from a caller who forgot to fill the field in -
        // and guessing wrong rewrites a paragraph nobody asked to touch. Against an *empty*
        // paragraph there is nothing to be wrong about: the only thing it can mean is
        // "write here", which is what filling in a blank document consists of.
        if (string.IsNullOrEmpty(anchor.Expect) && text.Length > 0)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "ChangeTextOp requires a non-empty 'expect' value identifying the text to replace. " +
                "To remove an entire paragraph, set 'with' to the empty string and 'expect' to the current paragraph text.",
                anchor));

        if (anchor.Expect.Length == 0)
            return OperationPreview.Ok(new ProposedChange
            {
                Target = anchor,
                Verb = "changeText",
                Before = string.Empty,
                After = op.With,
                Context = "empty paragraph",
                BlastRadius = 1
            });

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

        if (anchor.Expect.Length == 0)
        {
            Fill(context.Package, paragraph, op);
            return;
        }

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

    /// <summary>
    /// Writes text into a paragraph that has none. There is nothing to replace, so this
    /// adds a run rather than isolating a span - and under tracked changes it is recorded
    /// as an insertion, since that is what it is.
    /// </summary>
    private void Fill(IOpenXmlPackage package, OpenXmlElement paragraph, ChangeTextOp op)
    {
        if (op.With.Length == 0) return;

        var run = new Run();
        WordModel.Dialect.SetRunText(run, op.With);

        // A run follows w:pPr, which is the paragraph's first child when it has one.
        var properties = paragraph.GetFirstChild<ParagraphProperties>();

        if (op.Mode == ChangeMode.Direct)
        {
            if (properties is null) paragraph.InsertAt(run, 0);
            else paragraph.InsertAfter(run, properties);
            return;
        }

        var inserted = new InsertedRun
        {
            Author = "OfficeAgent",
            Date = _clock.GetUtcNow().UtcDateTime,
            Id = new WordRevisionIdAllocator(package).Next().ToString()
        };
        inserted.AppendChild(run);

        if (properties is null) paragraph.InsertAt(inserted, 0);
        else paragraph.InsertAfter(inserted, properties);
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
