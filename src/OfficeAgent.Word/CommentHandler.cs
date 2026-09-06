using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using DocumentFormat.OpenXml.Wordprocessing;
using W15 = DocumentFormat.OpenXml.Office2013.Word;

namespace OfficeAgent.Word;

/// <summary>
/// The Word comment lifecycle: attach a comment to an anchored span, reply into an
/// existing thread, mark a thread resolved, or delete one.
/// </summary>
/// <remarks>
/// Word splits a comment across two parts. <c>comments.xml</c> holds the body; the thread
/// - who replied to whom, and whether the conversation is finished - lives in
/// <c>commentsExtended.xml</c>, which addresses each comment by the <c>w14:paraId</c> of
/// its last paragraph. Every comment written here gets that id and an entry in that part,
/// so a comment this engine adds can be replied to and resolved like any other, by Word or
/// by a later plan.
/// </remarks>
internal sealed class CommentHandler : IOperationHandler
{
    private readonly TimeProvider _clock;

    public CommentHandler(TimeProvider clock) => _clock = clock;

    public bool CanHandle(PlanOperation operation) => operation switch
    {
        CommentOp { Target: TextSpanAnchor, Action: CommentAction.Add } => true,
        CommentOp { Target: NodeAnchor { Kind: "comment" }, Action: CommentAction.Reply } => true,
        CommentOp { Target: NodeAnchor { Kind: "comment" }, Action: CommentAction.Resolve } => true,
        CommentOp { Target: NodeAnchor { Kind: "comment" }, Action: CommentAction.Remove } => true,
        _ => false
    };

    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (CommentOp)operation;
        return op.Action == CommentAction.Add
            ? PreviewAdd(context, op, (TextSpanAnchor)op.Target)
            : PreviewOnExisting(context, op, (NodeAnchor)op.Target);
    }

    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (CommentOp)operation;
        var main = WordModel.Doc(context.Package).MainDocumentPart!;

        switch (op.Action)
        {
            case CommentAction.Add:
                ApplyAdd(context, op, (TextSpanAnchor)op.Target);
                break;
            case CommentAction.Reply:
                ApplyReply(context, main, op, (NodeAnchor)op.Target);
                break;
            case CommentAction.Resolve:
                ApplyResolve(context, main, (NodeAnchor)op.Target);
                break;
            case CommentAction.Remove:
                ApplyRemove(main, (NodeAnchor)op.Target);
                break;
        }
    }

    // ── Add ──────────────────────────────────────────────────────────────

    private static OperationPreview PreviewAdd(ApplyContext context, CommentOp op, TextSpanAnchor anchor)
    {
        var paragraph = WordModel.ResolveParagraph(context, anchor.ParaId);
        if (paragraph is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No paragraph with id '{anchor.ParaId}'.", anchor));

        var text = WordModel.Text.GetLogicalText(paragraph);
        if (!string.IsNullOrEmpty(anchor.Expect))
        {
            int start = WordModel.Text.IndexOfOccurrence(text, anchor.Expect, anchor.Occurrence, WordModel.Comparison(true));
            if (start < 0)
                return OperationPreview.Fail(new ValidationError(
                    ValidationErrorCodes.ExpectMismatch,
                    $"Expected text '{anchor.Expect}' not found in paragraph '{anchor.ParaId}'.", anchor));
        }

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "comment",
            Before = string.Empty,
            After = op.Text,
            Context = string.IsNullOrEmpty(anchor.Expect) ? text : anchor.Expect,
            BlastRadius = 1
        });
    }

    private void ApplyAdd(ApplyContext context, CommentOp op, TextSpanAnchor anchor)
    {
        var doc = WordModel.Doc(context.Package);
        var main = doc.MainDocumentPart!;

        var paragraph = WordModel.ResolveParagraph(context, anchor.ParaId)
            ?? throw new InvalidOperationException($"Paragraph '{anchor.ParaId}' vanished before apply.");

        var comment = NewComment(context.Package, main, op.Text, op.Author, op.Initials);
        var id = comment.Id!.Value!;

        var rangeStart = new CommentRangeStart { Id = id };
        var rangeEnd = new CommentRangeEnd { Id = id };
        var reference = new Run(new CommentReference { Id = id });

        if (!string.IsNullOrEmpty(anchor.Expect))
        {
            var text = WordModel.Text.GetLogicalText(paragraph);
            int start = WordModel.Text.IndexOfOccurrence(text, anchor.Expect, anchor.Occurrence, WordModel.Comparison(true));
            var covered = WordModel.Text.IsolateSpan(paragraph, start, anchor.Expect.Length);
            var firstRun = (Run)covered[0];
            var lastRun = (Run)covered[covered.Count - 1];

            firstRun.InsertBeforeSelf(rangeStart);
            lastRun.InsertAfterSelf(rangeEnd);
            rangeEnd.InsertAfterSelf(reference);
        }
        else
        {
            var pPr = paragraph.ParagraphProperties;
            if (pPr is not null)
                pPr.InsertAfterSelf(rangeStart);
            else
                paragraph.InsertAt(rangeStart, 0);

            paragraph.AppendChild(rangeEnd);
            paragraph.AppendChild(reference);
        }

        Extended(main, comment, parentParaId: null);
    }

    // ── Reply, resolve, remove ───────────────────────────────────────────

    private static OperationPreview PreviewOnExisting(ApplyContext context, CommentOp op, NodeAnchor anchor)
    {
        var main = WordModel.Doc(context.Package).MainDocumentPart!;
        var comment = CommentNodeProvider.Locate(main, anchor.Path);
        if (comment is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No comment at path '{anchor.Path}'. Paths come from inspect_document, as 'comment#<id>'.",
                anchor));

        if (op.Action == CommentAction.Reply && string.IsNullOrEmpty(op.Text))
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "A comment reply needs 'text'.", anchor));

        var current = CommentNodeProvider.TextOf(comment);
        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "comment",
            Before = current,
            After = op.Action switch
            {
                CommentAction.Reply => op.Text,
                CommentAction.Resolve => "resolved",
                _ => "removed"
            },
            Context = anchor.Path,
            BlastRadius = op.Action == CommentAction.Remove ? 1 + RepliesTo(main, comment).Count : 1,
            Capability = Capability.Deterministic
        });
    }

    private void ApplyReply(ApplyContext context, MainDocumentPart main, CommentOp op, NodeAnchor anchor)
    {
        var parent = CommentNodeProvider.Locate(main, anchor.Path)
            ?? throw new InvalidOperationException($"Comment '{anchor.Path}' vanished before apply.");

        // The parent needs a paragraph id and an entry of its own before anything can point
        // at it: a comment Word wrote before threads existed has neither.
        var parentParaId = EnsureParaId(context.Package, parent);
        Extended(main, parent, parentParaId: null);

        var reply = NewComment(context.Package, main, op.Text, op.Author, op.Initials);
        Extended(main, reply, parentParaId);

        PlaceReplyMarkers(main, parentId: parent.Id!.Value!, replyId: reply.Id!.Value!);
    }

    private static void ApplyResolve(ApplyContext context, MainDocumentPart main, NodeAnchor anchor)
    {
        var comment = CommentNodeProvider.Locate(main, anchor.Path)
            ?? throw new InvalidOperationException($"Comment '{anchor.Path}' vanished before apply.");

        EnsureParaId(context.Package, comment);
        var entry = Extended(main, comment, parentParaId: null);
        entry.Done = true;
    }

    private static void ApplyRemove(MainDocumentPart main, NodeAnchor anchor)
    {
        var comment = CommentNodeProvider.Locate(main, anchor.Path)
            ?? throw new InvalidOperationException($"Comment '{anchor.Path}' vanished before apply.");

        // A reply whose parent is gone is a comment Word cannot place in a thread, so the
        // thread goes as a unit.
        var doomed = new List<Comment> { comment };
        doomed.AddRange(RepliesTo(main, comment));

        var ids = new HashSet<string>(
            doomed.Select(c => c.Id?.Value).Where(id => !string.IsNullOrEmpty(id))!,
            StringComparer.Ordinal);
        var paraIds = new HashSet<string>(
            doomed.Select(CommentNodeProvider.LastParaId).Where(id => !string.IsNullOrEmpty(id))!,
            StringComparer.OrdinalIgnoreCase);

        RemoveMarkers(main, ids);

        foreach (var c in doomed) c.Remove();

        if (main.WordprocessingCommentsExPart?.CommentsEx is { } extended)
            foreach (var entry in extended.Elements<W15.CommentEx>().ToList())
                if (entry.ParaId?.Value is { } paraId && paraIds.Contains(paraId))
                    entry.Remove();
    }

    /// <summary>Every comment whose <c>commentsExtended</c> entry names this one as its parent.</summary>
    private static List<Comment> RepliesTo(MainDocumentPart main, Comment parent)
    {
        var parentParaId = CommentNodeProvider.LastParaId(parent);
        if (string.IsNullOrEmpty(parentParaId)) return new List<Comment>();

        var index = CommentNodeProvider.ExtendedIndex(main);
        return CommentNodeProvider.Comments(main)
            .Where(c => CommentNodeProvider.LastParaId(c) is { Length: > 0 } paraId
                        && index.TryGetValue(paraId, out var entry)
                        && string.Equals(entry.ParaIdParent?.Value, parentParaId, StringComparison.OrdinalIgnoreCase))
            .ToList();
    }

    // ── Mechanics ────────────────────────────────────────────────────────

    private Comment NewComment(
        IOpenXmlPackage package, MainDocumentPart main, string text, string author, string initials)
    {
        var commentsPart = EnsureCommentsPart(main);
        var comments = commentsPart.Comments!;

        var comment = new Comment
        {
            Id = NextCommentId(comments).ToString(),
            Author = author,
            Initials = initials,
            Date = _clock.GetUtcNow().UtcDateTime
        };

        comment.AppendChild(new Paragraph(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }))
        {
            ParagraphId = WordModel.MintParaId(package)
        });
        comments.AppendChild(comment);
        return comment;
    }

    /// <summary>
    /// Gives a comment's last paragraph a <c>w14:paraId</c> if it has none, and returns it.
    /// </summary>
    private static string EnsureParaId(IOpenXmlPackage package, Comment comment)
    {
        var paragraph = comment.Elements<Paragraph>().LastOrDefault()
            ?? comment.AppendChild(new Paragraph());

        if (string.IsNullOrEmpty(paragraph.ParagraphId?.Value))
            paragraph.ParagraphId = WordModel.MintParaId(package);

        return paragraph.ParagraphId!.Value!;
    }

    /// <summary>
    /// The comment's <c>w15:commentEx</c> entry, created when it has none. Passing a
    /// <paramref name="parentParaId"/> makes it a reply; passing null leaves an existing
    /// parentage alone rather than flattening a thread.
    /// </summary>
    private static W15.CommentEx Extended(MainDocumentPart main, Comment comment, string? parentParaId)
    {
        var paraId = CommentNodeProvider.LastParaId(comment)
            ?? throw new InvalidOperationException("A comment needs a paragraph id before it can join a thread.");

        var part = main.WordprocessingCommentsExPart ?? main.AddNewPart<WordprocessingCommentsExPart>();
        var extended = part.CommentsEx ??= new W15.CommentsEx();

        var entry = extended.Elements<W15.CommentEx>()
            .FirstOrDefault(e => string.Equals(e.ParaId?.Value, paraId, StringComparison.OrdinalIgnoreCase));

        if (entry is null)
        {
            entry = new W15.CommentEx { ParaId = paraId, Done = false };
            extended.AppendChild(entry);
        }

        if (parentParaId is { Length: > 0 })
            entry.ParaIdParent = parentParaId;

        return entry;
    }

    /// <summary>
    /// Places the reply's range and reference markers immediately after the parent's, which
    /// is what puts the two on the same span and therefore in the same thread.
    /// </summary>
    private static void PlaceReplyMarkers(MainDocumentPart main, string parentId, string replyId)
    {
        foreach (var (root, _) in TextHostsOf(main))
        {
            var start = root.Descendants<CommentRangeStart>()
                .FirstOrDefault(s => string.Equals(s.Id?.Value, parentId, StringComparison.Ordinal));
            if (start is null) continue;

            var end = root.Descendants<CommentRangeEnd>()
                .FirstOrDefault(e => string.Equals(e.Id?.Value, parentId, StringComparison.Ordinal));
            var reference = root.Descendants<CommentReference>()
                .FirstOrDefault(r => string.Equals(r.Id?.Value, parentId, StringComparison.Ordinal));

            start.InsertAfterSelf(new CommentRangeStart { Id = replyId });

            // The reply's end and reference go after the parent's reference run, so the
            // parent's span closes before the reply's does - the nesting Word writes.
            var afterReference = reference?.Parent as Run;
            if (end is not null && afterReference is not null)
            {
                afterReference.InsertAfterSelf(new CommentRangeEnd { Id = replyId });
                root.Descendants<CommentRangeEnd>()
                    .First(e => string.Equals(e.Id?.Value, replyId, StringComparison.Ordinal))
                    .InsertAfterSelf(new Run(new CommentReference { Id = replyId }));
            }
            else if (end is not null)
            {
                end.InsertAfterSelf(new CommentRangeEnd { Id = replyId });
            }

            return;
        }
    }

    /// <summary>Strips every range and reference marker belonging to the named comments.</summary>
    private static void RemoveMarkers(MainDocumentPart main, HashSet<string> ids)
    {
        foreach (var (root, _) in TextHostsOf(main))
        {
            foreach (var start in root.Descendants<CommentRangeStart>().ToList())
                if (start.Id?.Value is { } id && ids.Contains(id)) start.Remove();

            foreach (var end in root.Descendants<CommentRangeEnd>().ToList())
                if (end.Id?.Value is { } id && ids.Contains(id)) end.Remove();

            foreach (var reference in root.Descendants<CommentReference>().ToList())
            {
                if (reference.Id?.Value is not { } id || !ids.Contains(id)) continue;

                // The reference sits in a run of its own; leaving that run behind would
                // leave an empty run where the marker was.
                var run = reference.Parent as Run;
                reference.Remove();
                if (run is not null && !run.HasChildren) run.Remove();
            }
        }
    }

    /// <summary>
    /// The parts a comment marker can live in. Comment markers reach into headers, footers
    /// and notes, so removing a comment has to sweep all of them, not just the body.
    /// </summary>
    private static IEnumerable<(OpenXmlElement Root, string Key)> TextHostsOf(MainDocumentPart main)
    {
        if (main.Document?.Body is { } body) yield return (body, string.Empty);

        int h = 0;
        foreach (var header in main.HeaderParts)
            if (header.Header is { } el) yield return (el, $"hdr{h++}");

        int f = 0;
        foreach (var footer in main.FooterParts)
            if (footer.Footer is { } el) yield return (el, $"ftr{f++}");

        if (main.FootnotesPart?.Footnotes is { } footnotes) yield return (footnotes, "fn");
        if (main.EndnotesPart?.Endnotes is { } endnotes) yield return (endnotes, "en");
    }

    private static WordprocessingCommentsPart EnsureCommentsPart(MainDocumentPart mainPart)
    {
        var part = mainPart.WordprocessingCommentsPart ?? mainPart.AddNewPart<WordprocessingCommentsPart>();
        part.Comments ??= new Comments();
        return part;
    }

    private static int NextCommentId(Comments comments)
    {
        int max = comments.Elements<Comment>()
            .Select(c => int.TryParse(c.Id?.Value, out var n) ? n : 0)
            .DefaultIfEmpty(0)
            .Max();
        return max + 1;
    }
}
