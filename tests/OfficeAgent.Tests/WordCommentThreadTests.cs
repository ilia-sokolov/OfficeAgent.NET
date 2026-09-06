using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Word;
using W15 = DocumentFormat.OpenXml.Office2013.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// Word comments are readable, answerable, and closable - not only writable. A review
/// workflow that can add a comment but cannot see the ones already in the document, reply
/// to them, or mark them dealt with is half a workflow.
/// </summary>
public class WordCommentThreadTests
{
    private static OfficeAgentClient Client() => new(new WordModule());

    [Fact]
    public void Inspect_lists_comments_with_their_author_and_text()
    {
        var document = Commented("Confirm this clause with legal.");

        var node = Client()
            .Inspect(new StreamHandle(new MemoryStream(document)))
            .Nodes.Single(n => n.Kind == "comment");

        Assert.Equal("comment#1", node.Path);
        Assert.Contains("Reviewer", node.Summary);
        Assert.Contains("Confirm this clause with legal.", node.Summary);
        Assert.DoesNotContain("(resolved)", node.Summary);
    }

    [Fact]
    public void A_new_comment_joins_the_commentsExtended_part_so_it_can_be_threaded()
    {
        var document = Commented("Please review.");
        AssertValid(document);

        using var stream = new MemoryStream(document);
        using var package = WordprocessingDocument.Open(stream, isEditable: false);
        var main = package.MainDocumentPart!;

        var comment = main.WordprocessingCommentsPart!.Comments!.Elements<Comment>().Single();
        var paraId = comment.Elements<Paragraph>().Single().ParagraphId?.Value;
        Assert.False(string.IsNullOrEmpty(paraId));

        var entry = main.WordprocessingCommentsExPart!.CommentsEx!.Elements<W15.CommentEx>().Single();
        Assert.Equal(paraId, entry.ParaId?.Value);
        Assert.False(entry.Done?.Value ?? false);
    }

    [Fact]
    public void A_reply_lands_in_the_thread_it_answers()
    {
        var document = Reply(Commented("Is thirty days right?"), "Forty-five, per the MSA.");
        AssertValid(document);

        var node = Client()
            .Inspect(new StreamHandle(new MemoryStream(document)))
            .Nodes.Single(n => n.Kind == "comment" && n.Path == "comment#2");
        Assert.Contains("reply to comment#1", node.Summary);

        using var stream = new MemoryStream(document);
        using var package = WordprocessingDocument.Open(stream, isEditable: false);
        var main = package.MainDocumentPart!;

        var parentParaId = main.WordprocessingCommentsPart!.Comments!
            .Elements<Comment>().Single(c => c.Id?.Value == "1")
            .Elements<Paragraph>().Last().ParagraphId?.Value;

        var replyParaId = main.WordprocessingCommentsPart.Comments
            .Elements<Comment>().Single(c => c.Id?.Value == "2")
            .Elements<Paragraph>().Last().ParagraphId?.Value;

        var replyEntry = main.WordprocessingCommentsExPart!.CommentsEx!
            .Elements<W15.CommentEx>().Single(e => e.ParaId?.Value == replyParaId);
        Assert.Equal(parentParaId, replyEntry.ParaIdParent?.Value);

        // The reply anchors to the same span, which is what makes Word show it as a reply
        // rather than as a second comment that happens to sit nearby.
        var body = main.Document.Body!;
        Assert.Equal(2, body.Descendants<CommentRangeStart>().Count());
        Assert.Equal(2, body.Descendants<CommentReference>().Count());
    }

    [Fact]
    public void Resolving_marks_the_thread_done_and_keeps_its_history()
    {
        var document = Apply(Commented("Check the indemnity cap."), new CommentOp
        {
            Target = new NodeAnchor { Kind = "comment", Path = "comment#1" },
            Action = CommentAction.Resolve
        });

        AssertValid(document);

        var node = Client()
            .Inspect(new StreamHandle(new MemoryStream(document)))
            .Nodes.Single(n => n.Kind == "comment");
        Assert.Contains("(resolved)", node.Summary);

        using var stream = new MemoryStream(document);
        using var package = WordprocessingDocument.Open(stream, isEditable: false);
        // The comment itself is untouched - only its status changed.
        Assert.Single(package.MainDocumentPart!.WordprocessingCommentsPart!.Comments!.Elements<Comment>());
    }

    [Fact]
    public void Removing_a_comment_takes_its_replies_and_its_markers_with_it()
    {
        var threaded = Reply(Commented("Original."), "Answer.");

        var document = Apply(threaded, new CommentOp
        {
            Target = new NodeAnchor { Kind = "comment", Path = "comment#1" },
            Action = CommentAction.Remove
        });

        AssertValid(document);

        using var stream = new MemoryStream(document);
        using var package = WordprocessingDocument.Open(stream, isEditable: false);
        var main = package.MainDocumentPart!;

        Assert.Empty(main.WordprocessingCommentsPart!.Comments!.Elements<Comment>());
        Assert.Empty(main.WordprocessingCommentsExPart!.CommentsEx!.Elements<W15.CommentEx>());

        // No orphaned range markers or empty runs left behind in the body.
        var body = main.Document.Body!;
        Assert.Empty(body.Descendants<CommentRangeStart>());
        Assert.Empty(body.Descendants<CommentRangeEnd>());
        Assert.Empty(body.Descendants<CommentReference>());
        Assert.DoesNotContain(body.Descendants<Run>(), r => !r.HasChildren);
    }

    [Fact]
    public void Removing_only_the_reply_leaves_the_comment_it_answered()
    {
        var threaded = Reply(Commented("Original."), "Answer.");

        var document = Apply(threaded, new CommentOp
        {
            Target = new NodeAnchor { Kind = "comment", Path = "comment#2" },
            Action = CommentAction.Remove
        });

        var comments = Client()
            .Inspect(new StreamHandle(new MemoryStream(document)))
            .Nodes.Where(n => n.Kind == "comment")
            .ToList();

        Assert.Single(comments);
        Assert.Equal("comment#1", comments[0].Path);
    }

    [Fact]
    public void A_comment_path_that_names_nothing_fails_before_anything_is_written()
    {
        var report = Client().Preview(
            new StreamHandle(new MemoryStream(Commented("Anything."))),
            new DocumentPlan
            {
                Operations = new PlanOperation[]
                {
                    new CommentOp
                    {
                        Target = new NodeAnchor { Kind = "comment", Path = "comment#99" },
                        Action = CommentAction.Resolve
                    }
                }
            });

        Assert.False(report.IsValid);
        Assert.Contains(report.Errors, e => e.Code == ValidationErrorCodes.AnchorNotFound);
    }

    [Fact]
    public void A_reply_with_no_text_is_refused()
    {
        var report = Client().Preview(
            new StreamHandle(new MemoryStream(Commented("Anything."))),
            new DocumentPlan
            {
                Operations = new PlanOperation[]
                {
                    new CommentOp
                    {
                        Target = new NodeAnchor { Kind = "comment", Path = "comment#1" },
                        Action = CommentAction.Reply
                    }
                }
            });

        Assert.False(report.IsValid);
        Assert.Contains(report.Errors, e => e.Code == ValidationErrorCodes.InvalidOperation);
    }

    // ── Fixtures ─────────────────────────────────────────────────────────

    private static byte[] Commented(string text)
    {
        var seeded = DocxFactory.Contract();
        var paraId = Client()
            .Inspect(new StreamHandle(new MemoryStream(seeded)))
            .Paragraphs.First(p => p.Text.Contains("shall provide services")).ParaId;

        return Apply(seeded, new CommentOp
        {
            Target = new TextSpanAnchor { ParaId = paraId, Expect = "shall provide services" },
            Text = text,
            Author = "Reviewer",
            Initials = "RV"
        });
    }

    private static byte[] Reply(byte[] document, string text) =>
        Apply(document, new CommentOp
        {
            Target = new NodeAnchor { Kind = "comment", Path = "comment#1" },
            Action = CommentAction.Reply,
            Text = text,
            Author = "Counsel",
            Initials = "CO"
        });

    private static byte[] Apply(byte[] document, PlanOperation operation)
    {
        using var applied = Client().Commit(
            new StreamHandle(new MemoryStream(document)),
            new DocumentPlan { Operations = new[] { operation } });

        Assert.True(applied.Committed,
            string.Join("; ", applied.Report.Errors.Select(e => $"{e.Code}: {e.Message}")));
        return applied.ToBytes();
    }

    private static void AssertValid(byte[] document)
    {
        using var stream = new MemoryStream(document);
        using var opened = WordprocessingDocument.Open(stream, isEditable: false);
        var problems = new OpenXmlValidator(FileFormatVersions.Office2019)
            .Validate(opened)
            .Select(e => $"{e.Path?.XPath}: {e.Description}")
            .ToList();
        Assert.True(problems.Count == 0, string.Join("; ", problems.Take(3)));
    }
}
