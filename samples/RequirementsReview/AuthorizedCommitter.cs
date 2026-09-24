using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;

namespace OfficeAgent.Samples.RequirementsReview;

/// <summary>A reviewer's decision.</summary>
public enum ReviewerDecision
{
    /// <summary>Write the proposal to a reviewed copy.</summary>
    Approved,

    /// <summary>Do not write it.</summary>
    Rejected,
}

/// <summary>
/// A reviewer's decision, bound to the exact plan, snapshot and input bytes they previewed.
/// The reviewer id must come from the host's authentication, never from model output.
/// </summary>
/// <param name="ReviewerId">Authenticated reviewer id.</param>
/// <param name="Decision">The decision.</param>
/// <param name="PlanSha256">Plan hash from the preview receipt the reviewer saw.</param>
/// <param name="InputSha256">Input hash from the same receipt.</param>
/// <param name="SnapshotETag">The snapshot the plan is bound to.</param>
/// <param name="DecidedAtUtc">When the decision was made.</param>
public sealed record ReviewerApproval(
    string ReviewerId,
    ReviewerDecision Decision,
    string PlanSha256,
    string InputSha256,
    string SnapshotETag,
    DateTimeOffset DecidedAtUtc)
{
    /// <summary>Creates a decision bound to what the proposal previewed.</summary>
    /// <param name="proposal">The proposal the reviewer saw.</param>
    /// <param name="reviewerId">Authenticated reviewer id.</param>
    /// <param name="decision">The decision.</param>
    /// <param name="decidedAtUtc">When.</param>
    /// <returns>The approval.</returns>
    public static ReviewerApproval For(ReviewProposal proposal, string reviewerId, ReviewerDecision decision, DateTimeOffset decidedAtUtc) => new(
        reviewerId,
        decision,
        proposal.PreviewReceipt?.PlanSha256 ?? string.Empty,
        proposal.PreviewReceipt?.InputSha256 ?? string.Empty,
        proposal.Evidence.Snapshot.ETag,
        decidedAtUtc);
}

/// <summary>How a commit attempt ended.</summary>
public enum CommitStatus
{
    /// <summary>A reviewed copy was written.</summary>
    Committed,

    /// <summary>The proposal is not in a state that can be committed.</summary>
    NotCommittable,

    /// <summary>No decision, an unknown reviewer, or a decision for a different plan.</summary>
    Unauthorized,

    /// <summary>The reviewer rejected the proposal.</summary>
    RejectedByReviewer,

    /// <summary>The document changed after approval. Nothing was written.</summary>
    Stale,

    /// <summary>OfficeAgent refused the plan. Nothing was written.</summary>
    CommitRejected,

    /// <summary>A copy was written, but from bytes other than the approved ones. Treat as not approved.</summary>
    CommittedAgainstChangedInput,
}

/// <summary>The result of a commit attempt.</summary>
/// <param name="Status">The status.</param>
/// <param name="Reason">Why.</param>
/// <param name="Approval">The decision presented.</param>
/// <param name="Revalidation">The pre-commit preview receipt, when one ran.</param>
/// <param name="Receipt">The commit receipt, when a commit ran.</param>
/// <param name="Output">The reviewed copy, when one was written.</param>
public sealed record CommitOutcome(
    CommitStatus Status,
    string Reason,
    ReviewerApproval? Approval,
    ApplyReceipt? Revalidation,
    ApplyReceipt? Receipt,
    DocumentReference? Output);

/// <summary>
/// The only code path that writes. It checks who approved, that the approval covers this exact
/// plan and these exact bytes, previews again, and only then commits to a new document.
/// </summary>
public sealed class AuthorizedCommitter
{
    private readonly OfficeAgentClient _client;

    /// <summary>Initializes a new instance of the <see cref="AuthorizedCommitter"/> class.</summary>
    /// <param name="client">The OfficeAgent client.</param>
    public AuthorizedCommitter(OfficeAgentClient client) => _client = client;

    /// <summary>Commits an approved proposal to a new document.</summary>
    /// <param name="proposal">The proposal.</param>
    /// <param name="approval">The reviewer's decision, or null when there is none.</param>
    /// <param name="outputName">Name of the reviewed copy.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    /// <returns>The outcome.</returns>
    public async Task<CommitOutcome> CommitAsync(
        ReviewProposal proposal,
        ReviewerApproval? approval,
        string outputName,
        CancellationToken cancellationToken = default)
    {
        if (proposal.Status != ProposalStatus.AwaitingApproval || proposal.Built is null || proposal.PreviewReceipt is null)
        {
            return new CommitOutcome(CommitStatus.NotCommittable, $"proposal status is {proposal.Status}", approval, null, null, null);
        }

        if (approval is null)
        {
            return new CommitOutcome(CommitStatus.Unauthorized, "no reviewer decision", null, null, null, null);
        }

        if (!proposal.Policy.AuthorizedReviewers.Contains(approval.ReviewerId))
        {
            return new CommitOutcome(CommitStatus.Unauthorized, "reviewer is not authorized by the review policy", approval, null, null, null);
        }

        if (!Same(approval.PlanSha256, proposal.PreviewReceipt.PlanSha256)
            || !Same(approval.InputSha256, proposal.PreviewReceipt.InputSha256)
            || !Same(approval.SnapshotETag, proposal.Evidence.Snapshot.ETag))
        {
            return new CommitOutcome(CommitStatus.Unauthorized, "the decision does not cover this plan, snapshot and input", approval, null, null, null);
        }

        if (approval.Decision != ReviewerDecision.Approved)
        {
            return new CommitOutcome(CommitStatus.RejectedByReviewer, "the reviewer rejected the proposal", approval, null, null, null);
        }

        var connectionId = proposal.Evidence.ConnectionId;
        var documentId = proposal.Evidence.DocumentId;
        var plan = proposal.Built.Plan;

        // Revalidate immediately before the write. Time has passed since the preview the
        // reviewer saw; the document may have moved.
        using (var revalidation = await _client.PreviewWithReceiptAsync(connectionId, documentId, plan, cancellationToken).ConfigureAwait(false))
        {
            if (!revalidation.Report.IsValid)
            {
                return Refused(revalidation.Report, approval, revalidation.Receipt, null);
            }

            if (revalidation.Receipt is null || !Same(revalidation.Receipt.InputSha256, approval.InputSha256))
            {
                return new CommitOutcome(CommitStatus.Stale, "document bytes changed since approval", approval, revalidation.Receipt, null, null);
            }

            if (!Same(revalidation.Receipt.PlanSha256, approval.PlanSha256))
            {
                return new CommitOutcome(CommitStatus.Unauthorized, "the plan changed since approval", approval, revalidation.Receipt, null, null);
            }

            ProviderApplyResult committed;
            try
            {
                committed = await _client.CommitAsync(
                    connectionId,
                    documentId,
                    plan,
                    new SaveDocumentOptions
                    {
                        Mode = SaveMode.NewDocument,
                        NewName = outputName,
                        Actor = new AuditActor { Subject = approval.ReviewerId },
                    },
                    cancellationToken).ConfigureAwait(false);
            }
            catch (DocumentProviderException ex)
            {
                // For example, a reviewed copy with that name already exists. New-document
                // mode never overwrites, so nothing was written.
                return new CommitOutcome(CommitStatus.CommitRejected, $"the document store refused the save: {ex.Message}", approval, revalidation.Receipt, null, null);
            }

            if (!committed.Committed)
            {
                return Refused(committed.Report, approval, revalidation.Receipt, committed.Receipt);
            }

            // The engine checks the snapshot again at commit, but the snapshot does not cover
            // every part. The receipt's input hash does; a mismatch means the copy was made
            // from bytes nobody approved.
            if (committed.Receipt is null || !Same(committed.Receipt.InputSha256, approval.InputSha256))
            {
                return new CommitOutcome(CommitStatus.CommittedAgainstChangedInput,
                    "a copy was written from bytes other than the approved ones; do not use it", approval, revalidation.Receipt, committed.Receipt, committed.Document);
            }

            return new CommitOutcome(CommitStatus.Committed, "reviewed copy written", approval, revalidation.Receipt, committed.Receipt, committed.Document);
        }
    }

    private static CommitOutcome Refused(ChangeReport report, ReviewerApproval approval, ApplyReceipt? revalidation, ApplyReceipt? receipt)
    {
        var stale = report.Errors.Any(e => e.Code == ValidationErrorCodes.StaleSnapshot);
        return new CommitOutcome(
            stale ? CommitStatus.Stale : CommitStatus.CommitRejected,
            string.Join("; ", report.Errors.Select(e => $"{e.Code}: {e.Message}")),
            approval,
            revalidation,
            receipt,
            null);
    }

    private static bool Same(string? a, string? b) =>
        !string.IsNullOrEmpty(a) && string.Equals(a, b, StringComparison.Ordinal);
}
