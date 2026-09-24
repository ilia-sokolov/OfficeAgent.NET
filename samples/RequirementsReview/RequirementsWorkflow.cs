using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;

namespace OfficeAgent.Samples.RequirementsReview;

/// <summary>One requirement after evidence, checks, evaluation, routing and drafting.</summary>
/// <param name="Requirement">The registered requirement.</param>
/// <param name="Evidence">Its anchored evidence.</param>
/// <param name="Deterministic">Deterministic check results.</param>
/// <param name="Decision">The evaluator's typed answers, when the evaluator was called.</param>
/// <param name="Routing">The routing decision.</param>
/// <param name="Draft">The drafted proposal, when the route asked for one.</param>
public sealed record ReviewItem(
    RegisteredRequirement Requirement,
    AnchoredDocumentEvidence Evidence,
    IReadOnlyList<DeterministicResult> Deterministic,
    RequirementDecision? Decision,
    RoutingDecision Routing,
    DraftResult? Draft);

/// <summary>Where a proposal stands.</summary>
public enum ProposalStatus
{
    /// <summary>A valid, previewed plan waits for a reviewer.</summary>
    AwaitingApproval,

    /// <summary>Every requirement passed and the snapshot is still current. Nothing to write.</summary>
    NothingToPropose,

    /// <summary>A hard rule stopped the run before any external call.</summary>
    Stopped,

    /// <summary>OfficeAgent refused the plan in preview. Nothing was written.</summary>
    PreviewRejected,

    /// <summary>The document changed after evidence was collected. Nothing was written.</summary>
    Stale,
}

/// <summary>The result of the proposal phase. Nothing has been written when this exists.</summary>
public sealed record ReviewProposal
{
    /// <summary>Gets the run id.</summary>
    public required string RunId { get; init; }

    /// <summary>Gets the status.</summary>
    public required ProposalStatus Status { get; init; }

    /// <summary>Gets a one-line explanation of the status.</summary>
    public required string StatusReason { get; init; }

    /// <summary>Gets the document evidence the run was bound to.</summary>
    public required DocumentEvidence Evidence { get; init; }

    /// <summary>Gets the requirement set.</summary>
    public required RequirementSet Requirements { get; init; }

    /// <summary>Gets the policy.</summary>
    public required ReviewPolicy Policy { get; init; }

    /// <summary>Gets the routed requirements.</summary>
    public required IReadOnlyList<ReviewItem> Items { get; init; }

    /// <summary>Gets the built plan, when one was built.</summary>
    public BuiltPlan? Built { get; init; }

    /// <summary>Gets the preview report, when the plan was previewed.</summary>
    public ChangeReport? Preview { get; init; }

    /// <summary>Gets the preview receipt: plan and input hashes a reviewer approves.</summary>
    public ApplyReceipt? PreviewReceipt { get; init; }

    /// <summary>Gets when the run started.</summary>
    public required DateTimeOffset StartedAtUtc { get; init; }
}

/// <summary>
/// The host workflow. Application code owns the requirements, evidence, questions, thresholds,
/// routing, plan, and commit. A review agent invokes the evaluator as a tool, but its summary
/// never decides what the host writes.
/// </summary>
public sealed class RequirementsWorkflow
{
    private readonly OfficeAgentClient _client;
    private readonly RequirementSet _requirements;
    private readonly ReviewPolicy _policy;
    private readonly IRequirementReviewAgent _reviewer;
    private readonly IProposalDrafter _drafter;
    private readonly TimeProvider _time;
    private readonly ILogger _logger;

    /// <summary>Initializes a new instance of the <see cref="RequirementsWorkflow"/> class.</summary>
    /// <param name="client">The OfficeAgent client.</param>
    /// <param name="requirements">The requirement set.</param>
    /// <param name="policy">The review policy.</param>
    /// <param name="reviewer">The Agent Framework reviewer with a Jev-backed tool.</param>
    /// <param name="drafter">The proposal drafter.</param>
    /// <param name="time">Clock.</param>
    /// <param name="logger">Logger.</param>
    public RequirementsWorkflow(
        OfficeAgentClient client,
        RequirementSet requirements,
        ReviewPolicy policy,
        IRequirementReviewAgent reviewer,
        IProposalDrafter drafter,
        TimeProvider? time = null,
        ILogger? logger = null)
    {
        _client = client;
        _requirements = requirements;
        _policy = policy;
        _reviewer = reviewer;
        _drafter = drafter;
        _time = time ?? TimeProvider.System;
        _logger = logger ?? NullLogger.Instance;
    }

    /// <summary>Runs everything up to and including preview. Never writes.</summary>
    /// <param name="connectionId">The connection.</param>
    /// <param name="documentId">The document.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    /// <returns>The proposal.</returns>
    public async Task<ReviewProposal> ProposeAsync(string connectionId, string documentId, CancellationToken cancellationToken = default)
    {
        var started = _time.GetUtcNow();
        var runId = $"run_{started:yyyyMMddTHHmmssZ}_{Guid.NewGuid().ToString("N")[..8]}";

        // 1. One read, one hash, one snapshot, evidence cut by the registry's sections.
        var evidence = await EvidenceCollector
            .CollectAsync(_client, connectionId, documentId, _requirements, cancellationToken)
            .ConfigureAwait(false);

        ReviewProposal Result(ProposalStatus status, string reason, IReadOnlyList<ReviewItem> items, BuiltPlan? built = null, ChangeReport? preview = null, ApplyReceipt? receipt = null) => new()
        {
            RunId = runId,
            Status = status,
            StatusReason = reason,
            Evidence = evidence,
            Requirements = _requirements,
            Policy = _policy,
            Items = items,
            Built = built,
            Preview = preview,
            PreviewReceipt = receipt,
            StartedAtUtc = started,
        };

        // 2. Hard rules first, locally. A stop leaves before any content leaves the host.
        var items = new List<ReviewItem>();
        foreach (var requirement in _requirements.Requirements.Where(r => r.Kind == RequirementKind.Deterministic))
        {
            var ev = evidence.ByRequirement[requirement.Id];
            var results = HardRuleEvaluator.Evaluate(requirement, ev);
            items.Add(new ReviewItem(requirement, ev, results, null, Router.RouteHardRule(requirement, results), null));
        }

        if (items.FirstOrDefault(i => i.Routing.Route == Route.Stop) is { } stop)
        {
            _logger.LogWarning("Run {RunId} stopped by {RequirementId}: {Reason}", runId, stop.Requirement.Id, stop.Routing.Reason);
            return Result(ProposalStatus.Stopped, $"{stop.Requirement.Id}: {stop.Routing.Reason}", items);
        }

        // 3. Check completeness before offering any requirement to the agent. The tool receives
        // only eligible IDs; it looks up the immutable requirement and anchored passage here.
        var semantic = _requirements.Requirements.Where(r => r.Kind == RequirementKind.Semantic).ToList();
        var eligible = new List<RegisteredRequirement>();
        var incomplete = new Dictionary<string, RoutingDecision>(StringComparer.Ordinal);
        var signalsById = new Dictionary<string, IReadOnlyList<DeterministicResult>>(StringComparer.Ordinal);
        foreach (var requirement in semantic)
        {
            var ev = evidence.ByRequirement[requirement.Id];
            var signals = HardRuleEvaluator.PassSignals(requirement, ev);
            signalsById[requirement.Id] = signals;

            if (Completeness(ev) is { } gap)
            {
                incomplete[requirement.Id] = Router.Incomplete(gap.Rule, gap.Reason, gap.Note);
            }
            else
            {
                eligible.Add(requirement);
            }
        }

        IReadOnlyDictionary<string, RequirementDecision> decisions;
        try
        {
            decisions = await _reviewer.ReviewAsync(eligible, evidence.ByRequirement, cancellationToken).ConfigureAwait(false);
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            throw;
        }
        catch (Exception ex)
        {
            decisions = eligible.ToDictionary(
                r => r.Id,
                r => RequirementDecision.Failed(r.Id, "agent-tool",
                    new EvaluationFailure(EvaluationFailureKind.EvaluatorError, ex.GetType().Name)),
                StringComparer.Ordinal);
        }

        foreach (var requirement in semantic)
        {
            var ev = evidence.ByRequirement[requirement.Id];
            var signals = signalsById[requirement.Id];
            if (incomplete.TryGetValue(requirement.Id, out var gap))
            {
                items.Add(new ReviewItem(requirement, ev, signals, null, gap, null));
                continue;
            }

            var decision = decisions.TryGetValue(requirement.Id, out var returned) ? returned
                : RequirementDecision.Failed(requirement.Id, "agent-tool",
                    new EvaluationFailure(EvaluationFailureKind.NotSent, "agent did not call evaluation"));
            if (!string.Equals(decision.RequirementId, requirement.Id, StringComparison.Ordinal))
            {
                decision = RequirementDecision.Failed(requirement.Id, decision.Provider,
                    new EvaluationFailure(EvaluationFailureKind.MalformedResponse, "decision is for a different requirement"));
            }
            items.Add(new ReviewItem(requirement, ev, signals, decision, Router.RouteSemantic(_policy, requirement, decision, signals), null));
        }

        // 4. Drafting only where the router already decided a violation.
        for (var i = 0; i < items.Count; i++)
        {
            var item = items[i];
            if (item.Routing.Action is WordAction.Comment or WordAction.CommentAndTrackedChange && item.Evidence.Items.Count > 0)
            {
                var draft = await _drafter.DraftAsync(
                    item.Requirement,
                    item.Evidence,
                    allowReplacement: item.Routing.Action == WordAction.CommentAndTrackedChange,
                    cancellationToken).ConfigureAwait(false);
                items[i] = item with { Draft = draft };
            }
        }

        // 5. One plan, bound to the snapshot taken before any of the above.
        var built = PlanBuilder.Build(evidence, items);

        // 6. Preview even an empty plan: it is also the freshness check.
        using var preview = await _client
            .PreviewWithReceiptAsync(connectionId, documentId, built.Plan, cancellationToken)
            .ConfigureAwait(false);

        if (!preview.Report.IsValid)
        {
            var stale = preview.Report.Errors.Any(e => e.Code == ValidationErrorCodes.StaleSnapshot);
            return Result(
                stale ? ProposalStatus.Stale : ProposalStatus.PreviewRejected,
                string.Join("; ", preview.Report.Errors.Select(e => $"{e.Code}: {e.Message}")),
                items, built, preview.Report, preview.Receipt);
        }

        // The snapshot covers the text hosts. The input hash covers every byte, including
        // comments and properties, so a change there is caught as well.
        if (preview.Receipt is null || !string.Equals(preview.Receipt.InputSha256, evidence.SourceSha256, StringComparison.Ordinal))
        {
            return Result(ProposalStatus.Stale, "document bytes changed after evidence was collected", items, built, preview.Report, preview.Receipt);
        }

        if (built.Plan.Operations.Count == 0)
        {
            return Result(ProposalStatus.NothingToPropose, "every requirement passed; snapshot still current", items, built, preview.Report, preview.Receipt);
        }

        return Result(ProposalStatus.AwaitingApproval, $"{built.Plan.Operations.Count} operation(s) previewed; waiting for a reviewer", items, built, preview.Report, preview.Receipt);
    }

    private (string Rule, string Reason, string Note)? Completeness(AnchoredDocumentEvidence evidence)
    {
        if (evidence.SectionMatches == 0)
        {
            return ("C1", $"no section headed \"{evidence.Section}\"", "No section with this heading was found.");
        }

        if (evidence.SectionMatches > 1)
        {
            return ("C2", $"{evidence.SectionMatches} sections headed \"{evidence.Section}\"", "This heading appears more than once.");
        }

        if (evidence.Items.Count == 0)
        {
            return ("C3", $"section \"{evidence.Section}\" is empty", "This section is empty.");
        }

        if (evidence.Items.FirstOrDefault(i => !i.ReviewStateKnown) is { } unknown)
        {
            return ("C4", $"review state of evidence {unknown.EvidenceId} could not be read", "The review state of this section could not be read.");
        }

        if (evidence.Items.FirstOrDefault(i => i.HasPendingRevision) is { } pending)
        {
            // Inspection shows pending insertions as if accepted and hides pending deletions.
            // Judging that text would judge a proposal, not the document.
            return ("C5", $"evidence {pending.EvidenceId} contains an unaccepted tracked change", "This section contains unaccepted tracked changes. Review it once they are accepted or rejected.");
        }

        if (evidence.Characters > _policy.MaxEvidenceCharacters)
        {
            // Truncating could drop the sentence that decides the requirement.
            return ("C6", $"evidence is {evidence.Characters} characters, over the {_policy.MaxEvidenceCharacters}-character external-evaluation budget", "This section is too long for the automated check.");
        }

        return null;
    }
}
