using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.Word;

/// <summary>Attaches a review comment to an anchored span (or whole paragraph).</summary>
internal sealed class CommentHandler : IOperationHandler
{
    private readonly TimeProvider _clock;

    public CommentHandler(TimeProvider clock) => _clock = clock;

    public bool CanHandle(PlanOperation operation) =>
        operation is CommentOp { Target: TextSpanAnchor, Action: CommentAction.Add };

    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (CommentOp)operation;
        var anchor = (TextSpanAnchor)op.Target;

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

    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (CommentOp)operation;
        var anchor = (TextSpanAnchor)op.Target;
        var doc = WordModel.Doc(context.Package);

        var paragraph = WordModel.ResolveParagraph(context, anchor.ParaId)
            ?? throw new InvalidOperationException($"Paragraph '{anchor.ParaId}' vanished before apply.");

        var commentsPart = EnsureCommentsPart(doc.MainDocumentPart!);
        var comments = commentsPart.Comments!;
        var id = NextCommentId(comments).ToString();

        var comment = new Comment
        {
            Id = id,
            Author = op.Author,
            Initials = op.Initials,
            Date = _clock.GetUtcNow().UtcDateTime
        };
        comment.AppendChild(new Paragraph(new Run(new Text(op.Text) { Space = SpaceProcessingModeValues.Preserve })));
        comments.AppendChild(comment);

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
