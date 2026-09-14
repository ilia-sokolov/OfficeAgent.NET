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

        if (PendingRevisionOverlap(paragraph, start, anchor.Expect.Length) is { } overlap)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.RevisionOverlap,
                $"'{anchor.Expect}' in paragraph '{anchor.ParaId}' spans {overlap}. " +
                "Applying the edit here would produce a redline that no longer rejects back to the " +
                "original text. Resolve the pending revisions first with a revision operation " +
                "(accept or reject, addressed by id, by 'author:<name>', or 'all'), re-inspect, " +
                "then reissue the edit. Alternatively target a span that lies wholly inside or " +
                "wholly outside the pending revision.",
                anchor));

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

    /// <summary>
    /// Describes why a span cannot be safely redlined, or null when it can.
    /// </summary>
    /// <remarks>
    /// <para>
    /// A tracked replacement writes one <c>w:del</c> and one <c>w:ins</c> beside the first
    /// run it covers. That is sound while the covered runs are contiguous siblings, and it
    /// is sound when they all sit inside the same earlier <c>w:ins</c> - deleting text a
    /// pending insertion added is exactly the <c>w:ins/w:del</c> nesting Word writes.
    /// </para>
    /// <para>
    /// It is not sound when the span reaches across a pending deletion or out of one
    /// revision container into another. The text view hides <c>w:del</c> and
    /// <c>w:moveFrom</c>, so two runs that read as adjacent can have another author's
    /// deleted words physically between them. Moving the replacement to the front of that
    /// span leaves the earlier deletion after it, and rejecting everything then restores
    /// the old words in the wrong place instead of restoring the original paragraph.
    /// Refusing before any mutation keeps that document unreachable.
    /// </para>
    /// </remarks>
    private static string? PendingRevisionOverlap(OpenXmlElement paragraph, int start, int length)
    {
        var covered = CoveredRuns(paragraph, start, length);
        if (covered.Count < 2) return null;

        // Runs drawn from different containers: part of the span is inside a revision the
        // rest is not, so no single insertion point represents the whole edit.
        var parents = covered.Select(run => run.Parent).Distinct().ToList();
        if (parents.Count > 1)
        {
            var named = parents
                .Select(parent => WordRevisions.TagOf(parent!) ?? "unmarked text")
                .Distinct();
            return $"more than one tracked-change container ({string.Join(" and ", named)})";
        }

        // One container, but with a pending deletion physically between the covered runs.
        var parent = covered[0].Parent!;
        var first = parent.ChildElements.ToList().IndexOf(covered[0]);
        var last = parent.ChildElements.ToList().IndexOf(covered[covered.Count - 1]);
        for (int i = first + 1; i < last; i++)
        {
            var between = parent.ChildElements[i];
            if (between is DeletedRun or MoveFromRun)
                return $"a pending {(between is DeletedRun ? "deletion" : "move")} by " +
                       $"'{WordRevisions.AuthorOf(between) ?? "unknown"}'";
        }

        return null;
    }

    /// <summary>
    /// The runs a logical span touches, found without mutating the paragraph. Span
    /// isolation splits runs, which a preview must never do.
    /// </summary>
    private static List<OpenXmlElement> CoveredRuns(OpenXmlElement paragraph, int start, int length)
    {
        var covered = new List<OpenXmlElement>();
        int end = start + length;
        int offset = 0;

        foreach (var run in WordModel.Dialect.GetRuns(paragraph))
        {
            int runLength = WordModel.Dialect.GetRunText(run).Length;
            int runEnd = offset + runLength;

            // A zero-length run inside the span still participates: the apply path isolates
            // it along with the rest, and an empty inserted run is exactly what an earlier
            // tracked deletion leaves behind.
            bool overlaps = runLength == 0
                ? offset > start && offset < end
                : offset < end && runEnd > start;

            if (overlaps) covered.Add(run);
            offset = runEnd;
            if (offset >= end && covered.Count > 0 && runLength > 0) break;
        }

        return covered;
    }

    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (ChangeTextOp)operation;
        var anchor = (TextSpanAnchor)op.Target;

        var paragraph = WordModel.ResolveParagraph(context, anchor.ParaId)
            ?? throw new InvalidOperationException($"Paragraph '{anchor.ParaId}' vanished before apply.");

        if (anchor.Expect.Length == 0)
        {
            Fill(context, paragraph, op);
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
            ApplyTracked(context, covered, op.With);
    }

    /// <summary>
    /// Writes text into a paragraph that has none. There is nothing to replace, so this
    /// adds a run rather than isolating a span - and under tracked changes it is recorded
    /// as an insertion, since that is what it is.
    /// </summary>
    private static void Fill(ApplyContext context, OpenXmlElement paragraph, ChangeTextOp op)
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

        var revision = context.Revision;
        var inserted = WordRevisions.Stamp(
            new InsertedRun(),
            new WordRevisionIdAllocator(context.Package).Next().ToString(),
            revision.Author,
            revision.TimestampUtc!.Value.UtcDateTime);
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

    private static void ApplyTracked(ApplyContext context, IReadOnlyList<OpenXmlElement> covered, string replacement)
    {
        var first = (Run)covered[0];
        var parent = first.Parent
            ?? throw new InvalidOperationException("Run has no parent paragraph.");

        var allocator = new WordRevisionIdAllocator(context.Package);
        var author = context.Revision.Author;
        var stamp = context.Revision.TimestampUtc!.Value.UtcDateTime;

        var deleted = WordRevisions.Stamp(
            new DeletedRun(), allocator.Next().ToString(), author, stamp);
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
        var inserted = WordRevisions.Stamp(
            new InsertedRun(), allocator.Next().ToString(), author, stamp);
        inserted.AppendChild(insertRun);

        parent.InsertBefore(deleted, first);
        parent.InsertBefore(inserted, first);

        foreach (var element in covered)
            element.Remove();
    }
}
