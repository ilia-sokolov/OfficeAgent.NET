using System.Text;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;

namespace OfficeAgent.Samples.ContractReview;

/// <summary>
/// Orchestrates one review: screen, judge, build, preview, apply. The host owns
/// the save mode, so the source contract is never a write target.
/// </summary>
public sealed class ReviewRunner
{
    private readonly OfficeAgentClient _client;
    private readonly Screener _screener;

    /// <summary>Initializes a new instance of the <see cref="ReviewRunner"/> class.</summary>
    /// <param name="client">The OfficeAgent client.</param>
    public ReviewRunner(OfficeAgentClient client)
    {
        _client = client;
        _screener = new Screener(client);
    }

    /// <summary>Runs a review end to end.</summary>
    /// <param name="connectionId">Connection holding the contract.</param>
    /// <param name="documentId">Opaque document id of the contract.</param>
    /// <param name="playbook">The playbook to apply.</param>
    /// <param name="judge">
    /// Produces judgements for the screened candidates. In production this is the
    /// MAF agent; in tests it is a scripted stand-in.
    /// </param>
    /// <param name="outputName">Name for the reviewed copy.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    /// <returns>The review outcome.</returns>
    public async Task<ReviewResult> RunAsync(
        string connectionId,
        string documentId,
        Playbook playbook,
        Func<ScreeningResult, CancellationToken, Task<IReadOnlyList<Finding>>> judge,
        string outputName,
        CancellationToken cancellationToken = default)
    {
        // Taken before screening, so it covers the document the candidates were
        // found in. Anchors catch a change to the text being edited; this catches
        // changes in other snapshotted text hosts.
        var inspected = await _client
            .InspectAsync(connectionId, documentId, cancellationToken: cancellationToken)
            .ConfigureAwait(false);
        var snapshot = inspected.Snapshot;

        var screening = await _screener
            .ScreenAsync(connectionId, documentId, playbook, cancellationToken)
            .ConfigureAwait(false);

        if (screening.Candidates.Count == 0)
        {
            var freshness = await _client.PreviewAsync(
                    connectionId,
                    documentId,
                    SnapshotOnlyPlan(snapshot),
                    cancellationToken)
                .ConfigureAwait(false);
            if (!freshness.IsValid)
            {
                return ReviewResult.Rejected(
                    screening,
                    Array.Empty<ReviewItem>(),
                    freshness,
                    Array.Empty<string>());
            }

            return ReviewResult.NothingToDo(screening);
        }

        var findings = await judge(screening, cancellationToken).ConfigureAwait(false);

        var judged = findings.Select(f => f.CandidateId).ToHashSet(StringComparer.Ordinal);
        var unjudged = screening.Candidates
            .Where(c => !judged.Contains(c.Id))
            .Select(c => c.Id)
            .ToList();

        // Fail closed. A candidate the model never judged was reviewed by nobody,
        // so the run cannot claim a result for this contract - and "no breaches"
        // is the most dangerous thing it could claim. Write nothing.
        if (unjudged.Count > 0)
        {
            var partial = new PlanBuilder(playbook).Build(screening.Candidates, findings, snapshot);
            return ReviewResult.Incomplete(screening, partial.Items, unjudged);
        }

        var build = new PlanBuilder(playbook).Build(screening.Candidates, findings, snapshot);

        // Preview even an empty plan. Besides validating operations, preview
        // verifies that the snapshot captured before screening is still current,
        // so a clean result is never reported for a contract that moved.
        var preview = await _client
            .PreviewAsync(connectionId, documentId, build.Plan, cancellationToken)
            .ConfigureAwait(false);

        if (!preview.IsValid)
        {
            return ReviewResult.Rejected(screening, build.Items, preview, unjudged);
        }

        if (build.Plan.Operations.Count == 0)
        {
            return ReviewResult.NoBreaches(screening, build.Items, unjudged);
        }

        var applied = await _client.CommitAsync(
            connectionId,
            documentId,
            build.Plan,
            new SaveDocumentOptions { Mode = SaveMode.NewDocument, NewName = outputName },
            cancellationToken).ConfigureAwait(false);

        // Preview validates each operation against the pre-apply document, so a
        // valid preview is not a promise that the commit succeeds. Check.
        if (!applied.Committed)
        {
            return ReviewResult.Rejected(screening, build.Items, applied.Report, unjudged);
        }

        return ReviewResult.Applied(screening, build.Items, applied, unjudged);
    }

    private static DocumentPlan SnapshotOnlyPlan(SnapshotToken snapshot) => new()
    {
        Snapshot = snapshot,
        Operations = Array.Empty<PlanOperation>(),
    };

    /// <summary>Renders the report a reviewer reads before opening Word.</summary>
    /// <param name="playbook">The playbook applied.</param>
    /// <param name="result">The review outcome.</param>
    /// <returns>A markdown report.</returns>
    public static string Report(Playbook playbook, ReviewResult result)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"# Contract review - {playbook.Name}");
        sb.AppendLine();

        var breaches = result.Items.Where(i => i.Violation).ToList();
        var clear = result.Items.Where(i => !i.Violation).ToList();

        sb.AppendLine($"- Candidates screened: {result.Screening.Candidates.Count}");
        sb.AppendLine($"- Breaches found: {breaches.Count}");
        sb.AppendLine($"- Checked and compliant: {clear.Count}");
        sb.AppendLine($"- Outcome: {result.Status}");
        if (result.OutputName is { Length: > 0 })
        {
            sb.AppendLine($"- Reviewed copy: {result.OutputName}");
        }

        sb.AppendLine();

        if (breaches.Count > 0)
        {
            sb.AppendLine("## Breaches");
            sb.AppendLine();
            sb.AppendLine("| Rule | Severity | Wording | Redline | Reason |");
            sb.AppendLine("| --- | --- | --- | --- | --- |");
            foreach (var item in breaches.OrderByDescending(i => i.Severity))
            {
                sb.AppendLine(
                    $"| {item.RuleId} | {item.Severity} | {Escape(item.MatchedText)} | " +
                    $"{item.Redline} | {Escape(item.Rationale)} |");
            }

            sb.AppendLine();
        }

        if (clear.Count > 0)
        {
            sb.AppendLine("## Checked, no change proposed");
            sb.AppendLine();
            foreach (var item in clear)
            {
                sb.AppendLine($"- {item.RuleId}: {Escape(item.Rationale)}");
            }

            sb.AppendLine();
        }

        if (result.Screening.Undetected.Count > 0)
        {
            sb.AppendLine("## Not detected");
            sb.AppendLine();
            sb.AppendLine(
                "These rules matched nothing. That means the patterns did not fire, " +
                "which is not the same as the contract complying - check them by hand.");
            sb.AppendLine();
            foreach (var rule in result.Screening.Undetected)
            {
                sb.AppendLine($"- {rule.RuleId}: {rule.Title}");
            }

            sb.AppendLine();
        }

        if (result.Unjudged.Count > 0)
        {
            sb.AppendLine("## Not judged");
            sb.AppendLine();
            sb.AppendLine(
                "The model returned no judgement for these screened candidates, so the " +
                "review is incomplete and no document was written. Re-run, or judge " +
                "them by hand.");
            sb.AppendLine();
            foreach (var id in result.Unjudged)
            {
                sb.AppendLine($"- {id}");
            }

            sb.AppendLine();
        }

        if (result.Errors.Count > 0)
        {
            sb.AppendLine("## Plan rejected");
            sb.AppendLine();
            foreach (var error in result.Errors)
            {
                sb.AppendLine($"- {error}");
            }
        }

        return sb.ToString();
    }

    private static string Escape(string value) => value.Replace("|", "\\|", StringComparison.Ordinal);
}

/// <summary>What happened to the review.</summary>
public enum ReviewStatus
{
    /// <summary>Screening matched nothing anywhere.</summary>
    NoCandidates,

    /// <summary>The model left at least one candidate unjudged; nothing was written.</summary>
    Incomplete,

    /// <summary>Candidates were judged, none breached.</summary>
    NoBreaches,

    /// <summary>The plan failed validation; nothing was written.</summary>
    Rejected,

    /// <summary>A reviewed copy was written.</summary>
    Applied,
}

/// <summary>The outcome of one review.</summary>
public sealed record ReviewResult
{
    /// <summary>Gets the status.</summary>
    public ReviewStatus Status { get; init; }

    /// <summary>Gets the screening result.</summary>
    public ScreeningResult Screening { get; init; } = new(Array.Empty<Candidate>(), Array.Empty<UndetectedRule>());

    /// <summary>Gets the report rows.</summary>
    public IReadOnlyList<ReviewItem> Items { get; init; } = Array.Empty<ReviewItem>();

    /// <summary>Gets candidates the model never judged. An audit gap, so it is reported.</summary>
    public IReadOnlyList<string> Unjudged { get; init; } = Array.Empty<string>();

    /// <summary>Gets validation errors when the plan was rejected.</summary>
    public IReadOnlyList<string> Errors { get; init; } = Array.Empty<string>();

    /// <summary>Gets the reviewed copy's document id, when one was written.</summary>
    public string? OutputDocumentId { get; init; }

    /// <summary>Gets the reviewed copy's name, when one was written.</summary>
    public string? OutputName { get; init; }

    internal static ReviewResult NothingToDo(ScreeningResult screening) => new()
    {
        Status = ReviewStatus.NoCandidates,
        Screening = screening,
    };

    internal static ReviewResult Incomplete(
        ScreeningResult screening,
        IReadOnlyList<ReviewItem> items,
        IReadOnlyList<string> unjudged) => new()
    {
        Status = ReviewStatus.Incomplete,
        Screening = screening,
        Items = items,
        Unjudged = unjudged,
    };

    internal static ReviewResult NoBreaches(
        ScreeningResult screening,
        IReadOnlyList<ReviewItem> items,
        IReadOnlyList<string> unjudged) => new()
    {
        Status = ReviewStatus.NoBreaches,
        Screening = screening,
        Items = items,
        Unjudged = unjudged,
    };

    internal static ReviewResult Rejected(
        ScreeningResult screening,
        IReadOnlyList<ReviewItem> items,
        ChangeReport report,
        IReadOnlyList<string> unjudged) => new()
    {
        Status = ReviewStatus.Rejected,
        Screening = screening,
        Items = items,
        Unjudged = unjudged,
        Errors = report.Errors.Select(e => $"{e.Code}: {e.Message}").ToList(),
    };

    internal static ReviewResult Applied(
        ScreeningResult screening,
        IReadOnlyList<ReviewItem> items,
        ProviderApplyResult applied,
        IReadOnlyList<string> unjudged) => new()
    {
        Status = ReviewStatus.Applied,
        Screening = screening,
        Items = items,
        Unjudged = unjudged,
        OutputDocumentId = applied.Document?.ItemId,
        OutputName = applied.Document?.Name,
    };
}
