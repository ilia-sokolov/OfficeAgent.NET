using OfficeAgent.Abstractions;
using OfficeAgent.Samples.ContractReview;
using Xunit;

namespace ContractReview.Tests;

/// <summary>
/// The design review raised a hazard: a tracked replacement strikes the original
/// wording through, so a comment anchored to that same wording may not resolve if
/// it is applied afterwards. These tests record what the engine actually does, so
/// the ordering rule in <see cref="PlanBuilder"/> rests on measurement.
///
/// The measured behaviour is worth knowing in its own right: <b>preview accepts
/// both orders, but only one of them commits.</b> Preview validates each
/// operation against the pre-apply document, so it cannot see that the first
/// operation invalidates the second one's anchor.
/// </summary>
public sealed class OrderingProbeTests
{
    private static TextSpanAnchor Span(string paraId, string expect) =>
        new() { ParaId = paraId, Expect = expect, Occurrence = 0 };

    private static CommentOp Comment(string paraId) => new()
    {
        Target = Span(paraId, "60 days"),
        Text = "Payment terms exceed the house position.",
        Author = PlanBuilder.CommentAuthor,
        Initials = PlanBuilder.CommentInitials,
        Action = CommentAction.Add,
    };

    private static ChangeTextOp Redline(string paraId) => new()
    {
        Target = Span(paraId, "60 days"),
        With = "30 days",
        Mode = ChangeMode.Tracked,
    };

    private static async Task<string> PaymentParaIdAsync(ReviewHarness harness)
    {
        var hits = await harness.Client.FindAsync(
            ReviewHarness.ConnectionId,
            harness.DocumentId,
            new FindQuery { Pattern = "60 days" });

        return ((TextSpanAnchor)hits.Single().Anchor).ParaId;
    }

    [Fact]
    public async Task Preview_accepts_both_orders()
    {
        using var harness = await ReviewHarness.CreateAsync();
        var paraId = await PaymentParaIdAsync(harness);

        foreach (var commentFirst in new[] { true, false })
        {
            var plan = new DocumentPlan
            {
                Operations = commentFirst
                    ? new PlanOperation[] { Comment(paraId), Redline(paraId) }
                    : new PlanOperation[] { Redline(paraId), Comment(paraId) },
            };

            var preview = await harness.Client.PreviewAsync(
                ReviewHarness.ConnectionId, harness.DocumentId, plan);

            Assert.True(
                preview.IsValid,
                $"commentFirst={commentFirst}: " +
                string.Join("; ", preview.Errors.Select(e => $"{e.Code}: {e.Message}")));
        }
    }

    [Fact]
    public async Task Comment_before_redline_commits_and_produces_both()
    {
        using var harness = await ReviewHarness.CreateAsync();
        var paraId = await PaymentParaIdAsync(harness);

        var applied = await harness.Client.CommitAsync(
            ReviewHarness.ConnectionId,
            harness.DocumentId,
            new DocumentPlan { Operations = new PlanOperation[] { Comment(paraId), Redline(paraId) } },
            new SaveDocumentOptions { Mode = SaveMode.NewDocument, NewName = "comment-first.docx" });

        Assert.True(
            applied.Committed,
            string.Join("; ", applied.Report.Errors.Select(e => $"{e.Code}: {e.Message}")));

        Assert.Contains("<w:ins ", harness.DocumentXml("comment-first.docx"), StringComparison.Ordinal);
        Assert.Contains(
            PlanBuilder.CommentAuthor,
            harness.CommentsXml("comment-first.docx"),
            StringComparison.Ordinal);
    }

    [Fact]
    public async Task Redline_before_comment_fails_at_commit_with_expect_mismatch()
    {
        // This is the trap the PlanBuilder ordering rule exists to avoid. The
        // tracked replacement strikes '60 days' through, struck text is no longer
        // part of the paragraph's visible text, and the comment that follows can
        // no longer find its anchor. The plan is atomic, so nothing is written.
        using var harness = await ReviewHarness.CreateAsync();
        var paraId = await PaymentParaIdAsync(harness);

        var applied = await harness.Client.CommitAsync(
            ReviewHarness.ConnectionId,
            harness.DocumentId,
            new DocumentPlan { Operations = new PlanOperation[] { Redline(paraId), Comment(paraId) } },
            new SaveDocumentOptions { Mode = SaveMode.NewDocument, NewName = "change-first.docx" });

        Assert.False(applied.Committed);
        Assert.Contains(applied.Report.Errors, e => e.Code == "expect-mismatch");
        Assert.False(harness.OutputExists("change-first.docx"));
    }

    [Fact]
    public async Task PlanBuilder_emits_the_order_that_commits()
    {
        using var harness = await ReviewHarness.CreateAsync();

        var result = await new ReviewRunner(harness.Client).RunAsync(
            ReviewHarness.ConnectionId,
            harness.DocumentId,
            TestPlaybooks.PaymentRedline(),
            (screening, _) => Task.FromResult<IReadOnlyList<Finding>>(
                screening.Candidates.Select(c => new Finding
                {
                    CandidateId = c.Id,
                    Violation = true,
                    Rationale = "Sixty days exceeds the house position.",
                    ReplacementText = "30 days",
                }).ToList()),
            "ordered.reviewed.docx");

        Assert.Equal(ReviewStatus.Applied, result.Status);
        Assert.True(harness.OutputExists("ordered.reviewed.docx"));
    }
}
