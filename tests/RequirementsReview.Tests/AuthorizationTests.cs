using OfficeAgent.Abstractions;
using OfficeAgent.Samples.RequirementsReview;
using Xunit;

namespace RequirementsReview.Tests;

/// <summary>
/// The write path. Only an authorized reviewer's decision, bound to the exact plan and bytes
/// they previewed, and revalidated just before the commit, produces a file.
/// </summary>
public sealed class AuthorizationTests
{
    private const string Output = "method-statement.reviewed.docx";

    [Fact]
    public async Task The_proposal_phase_never_writes()
    {
        using var harness = await Harness.CreateAsync();
        var before = harness.SourceSha256();

        var proposal = await harness.ProposeAsync();

        Assert.Equal(ProposalStatus.AwaitingApproval, proposal.Status);
        Assert.Equal(before, harness.SourceSha256());
        Assert.Equal(1, harness.DocxCount());
        Assert.Equal(ApplyOutcome.Previewed, proposal.PreviewReceipt!.Outcome);
    }

    [Fact]
    public async Task No_decision_means_no_write()
    {
        using var harness = await Harness.CreateAsync();
        var proposal = await harness.ProposeAsync();

        var outcome = await new AuthorizedCommitter(harness.Client).CommitAsync(proposal, null, Output);

        Assert.Equal(CommitStatus.Unauthorized, outcome.Status);
        Assert.False(harness.Exists(Output));
    }

    [Fact]
    public async Task A_reviewer_outside_the_policy_cannot_approve()
    {
        using var harness = await Harness.CreateAsync();
        var proposal = await harness.ProposeAsync();

        var outcome = await harness.ApproveAsync(proposal, reviewer: "drafting-agent");

        Assert.Equal(CommitStatus.Unauthorized, outcome.Status);
        Assert.Equal("reviewer is not authorized by the review policy", outcome.Reason);
        Assert.False(harness.Exists(Output));
    }

    [Fact]
    public async Task An_approval_for_a_different_plan_does_not_authorize_this_one()
    {
        using var harness = await Harness.CreateAsync();
        var proposal = await harness.ProposeAsync();
        var approval = ReviewerApproval.For(proposal, Harness.Reviewer, ReviewerDecision.Approved, DateTimeOffset.UtcNow) with
        {
            PlanSha256 = new string('0', 64),
        };

        var outcome = await new AuthorizedCommitter(harness.Client).CommitAsync(proposal, approval, Output);

        Assert.Equal(CommitStatus.Unauthorized, outcome.Status);
        Assert.False(harness.Exists(Output));
    }

    [Fact]
    public async Task An_approval_bound_to_another_snapshot_does_not_authorize_this_one()
    {
        using var harness = await Harness.CreateAsync();
        var proposal = await harness.ProposeAsync();
        var approval = ReviewerApproval.For(proposal, Harness.Reviewer, ReviewerDecision.Approved, DateTimeOffset.UtcNow) with
        {
            SnapshotETag = "other",
        };

        var outcome = await new AuthorizedCommitter(harness.Client).CommitAsync(proposal, approval, Output);

        Assert.Equal(CommitStatus.Unauthorized, outcome.Status);
        Assert.False(harness.Exists(Output));
    }

    [Fact]
    public async Task A_plan_changed_after_approval_is_not_written()
    {
        using var harness = await Harness.CreateAsync();
        var proposal = await harness.ProposeAsync();
        var approval = ReviewerApproval.For(proposal, Harness.Reviewer, ReviewerDecision.Approved, DateTimeOffset.UtcNow);

        // Swap the approved comment for something else after the reviewer said yes.
        var ops = proposal.Built!.Plan.Operations
            .Select(o => o is CommentOp c && c.Text.StartsWith("[REQ-REC-01", StringComparison.Ordinal)
                ? new CommentOp { Target = c.Target, Text = "[REQ-REC-01 v1] Recovery procedure accepted.", Author = c.Author, Initials = c.Initials, Action = c.Action }
                : o)
            .ToList();
        Assert.NotEqual(proposal.Built.Plan.Operations.OfType<CommentOp>().Select(c => c.Text), ops.OfType<CommentOp>().Select(c => c.Text));
        var tampered = proposal with { Built = proposal.Built with { Plan = new DocumentPlan { Snapshot = proposal.Built.Plan.Snapshot, Revision = proposal.Built.Plan.Revision, Operations = ops } } };

        var outcome = await new AuthorizedCommitter(harness.Client).CommitAsync(tampered, approval, Output);

        Assert.Equal((CommitStatus.Unauthorized, "the plan changed since approval"), (outcome.Status, outcome.Reason));
        Assert.False(harness.Exists(Output));
    }

    [Fact]
    public async Task A_rejection_is_recorded_and_writes_nothing()
    {
        using var harness = await Harness.CreateAsync();
        var proposal = await harness.ProposeAsync();
        var rejection = ReviewerApproval.For(proposal, Harness.Reviewer, ReviewerDecision.Rejected, DateTimeOffset.UtcNow);

        var outcome = await new AuthorizedCommitter(harness.Client).CommitAsync(proposal, rejection, Output);

        Assert.Equal(CommitStatus.RejectedByReviewer, outcome.Status);
        Assert.False(harness.Exists(Output));
    }

    [Fact]
    public async Task A_document_edited_after_approval_blocks_the_commit()
    {
        using var harness = await Harness.CreateAsync();
        var proposal = await harness.ProposeAsync();
        var approval = ReviewerApproval.For(proposal, Harness.Reviewer, ReviewerDecision.Approved, DateTimeOffset.UtcNow);

        await File.WriteAllBytesAsync(harness.SourcePath, MethodStatementFixture.Create(new MethodStatementOptions { Scope = "This method statement covers replacing AHU-3" }));

        var outcome = await new AuthorizedCommitter(harness.Client).CommitAsync(proposal, approval, Output);

        Assert.Equal(CommitStatus.Stale, outcome.Status);
        Assert.Contains(ValidationErrorCodes.StaleSnapshot, outcome.Reason);
        Assert.Null(outcome.Receipt);
        Assert.False(harness.Exists(Output));
    }

    [Fact]
    public async Task An_approved_commit_writes_a_new_copy_and_records_the_reviewer()
    {
        using var harness = await Harness.CreateAsync();
        var source = harness.SourceSha256();
        var proposal = await harness.ProposeAsync();

        var outcome = await harness.ApproveAsync(proposal);

        Assert.Equal(CommitStatus.Committed, outcome.Status);
        Assert.True(harness.Exists(Output));
        Assert.Equal(source, harness.SourceSha256());
        Assert.Equal(Harness.Reviewer, outcome.Receipt!.Actor!.Subject);
        Assert.Equal(proposal.PreviewReceipt!.PlanSha256, outcome.Receipt.PlanSha256);
        Assert.Equal(proposal.Evidence.SourceSha256, outcome.Receipt.InputSha256);
        Assert.Equal(ApplyOutcome.Committed, outcome.Receipt.Outcome);
        Assert.NotNull(outcome.Receipt.OutputSha256);
        Assert.Equal(PlanBuilder.Author, outcome.Receipt.Revision.Author);
    }

    [Fact]
    public async Task An_existing_reviewed_copy_is_never_overwritten()
    {
        using var harness = await Harness.CreateAsync();
        var first = await harness.ApproveAsync(await harness.ProposeAsync());
        var firstBytes = File.ReadAllBytes(Path.Combine(harness.Root, Output));

        var second = await harness.ApproveAsync(await harness.ProposeAsync());

        Assert.Equal(CommitStatus.Committed, first.Status);
        Assert.Equal(CommitStatus.CommitRejected, second.Status);
        Assert.StartsWith("the document store refused the save", second.Reason);
        Assert.Equal(firstBytes, File.ReadAllBytes(Path.Combine(harness.Root, Output)));
    }

    [Fact]
    public async Task A_stopped_run_cannot_be_committed()
    {
        using var harness = await Harness.CreateAsync(new MethodStatementOptions { InstallationWindow = "Installation window: tbc." });
        var proposal = await harness.ProposeAsync();

        var outcome = await harness.ApproveAsync(proposal);

        Assert.Equal(CommitStatus.NotCommittable, outcome.Status);
        Assert.Equal(1, harness.DocxCount());
    }
}
