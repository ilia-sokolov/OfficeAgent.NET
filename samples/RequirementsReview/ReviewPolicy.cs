using System.Globalization;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace OfficeAgent.Samples.RequirementsReview;

/// <summary>Thresholds the router applies. Owned by the application, never by a model.</summary>
/// <param name="MinEvidenceConfidence">Minimum confidence for <c>evidence_status = sufficient</c>.</param>
/// <param name="MinResultConfidence">Minimum confidence for a satisfied or violated result.</param>
/// <param name="MaterialityProposalScore">Expected materiality level at or above which a violation gets a tracked change.</param>
/// <param name="MinMaterialityConfidence">Below this, materiality is treated as material rather than trusted to downgrade.</param>
public sealed record Thresholds(
    double MinEvidenceConfidence,
    double MinResultConfidence,
    double MaterialityProposalScore,
    double MinMaterialityConfidence);

/// <summary>The versioned review policy: evaluator settings, thresholds, and who may approve.</summary>
public sealed record ReviewPolicy
{
    /// <summary>Gets the policy version recorded with every route.</summary>
    public required string Version { get; init; }

    /// <summary>Gets the SHA-256 of the JSON the policy was loaded from.</summary>
    public required string Sha256 { get; init; }

    /// <summary>Gets the model requested from the evaluator.</summary>
    public required string EvaluatorModel { get; init; }

    /// <summary>Gets the evaluator timeout.</summary>
    public required TimeSpan EvaluatorTimeout { get; init; }

    /// <summary>Gets the largest evidence payload that may leave the host for one requirement.</summary>
    public required int MaxEvidenceCharacters { get; init; }

    /// <summary>Gets the routing thresholds.</summary>
    public required Thresholds Thresholds { get; init; }

    /// <summary>Gets the reviewer ids allowed to approve a write.</summary>
    public required IReadOnlySet<string> AuthorizedReviewers { get; init; }

    /// <summary>Loads a policy from a file.</summary>
    /// <param name="path">The path.</param>
    /// <returns>The policy.</returns>
    public static ReviewPolicy FromFile(string path) => FromJson(File.ReadAllText(path));

    /// <summary>Loads and validates a policy. Out-of-range values throw.</summary>
    /// <param name="json">The JSON text.</param>
    /// <returns>The policy.</returns>
    public static ReviewPolicy FromJson(string json)
    {
        var dto = JsonSerializer.Deserialize<PolicyDto>(json, JsonOptions)
                  ?? throw new InvalidDataException("Policy is empty.");

        if (string.IsNullOrWhiteSpace(dto.Version) || dto.Evaluator is null || dto.Thresholds is null)
        {
            throw new InvalidDataException("Policy needs version, evaluator and thresholds.");
        }

        var t = dto.Thresholds;
        foreach (var (name, value) in new[]
                 {
                     ("minEvidenceConfidence", t.MinEvidenceConfidence),
                     ("minResultConfidence", t.MinResultConfidence),
                     ("minMaterialityConfidence", t.MinMaterialityConfidence),
                 })
        {
            if (value is not (> 0 and <= 1))
            {
                throw new InvalidDataException($"Policy threshold {name} must be in (0,1].");
            }
        }

        if (t.MaterialityProposalScore is not (>= 0 and <= 3))
        {
            throw new InvalidDataException("Policy threshold materialityProposalScore must be in [0,3].");
        }

        if (string.IsNullOrWhiteSpace(dto.Evaluator.Model)
            || dto.Evaluator.TimeoutSeconds is not (> 0 and <= 300)
            || dto.Evaluator.MaxEvidenceCharacters is not (> 0 and <= 100_000))
        {
            throw new InvalidDataException("Policy evaluator settings are missing or out of range.");
        }

        return new ReviewPolicy
        {
            Version = dto.Version!,
            Sha256 = RequirementSet.Hash(json),
            EvaluatorModel = dto.Evaluator.Model!,
            EvaluatorTimeout = TimeSpan.FromSeconds(dto.Evaluator.TimeoutSeconds),
            MaxEvidenceCharacters = dto.Evaluator.MaxEvidenceCharacters,
            Thresholds = new Thresholds(t.MinEvidenceConfidence, t.MinResultConfidence, t.MaterialityProposalScore, t.MinMaterialityConfidence),
            AuthorizedReviewers = (dto.AuthorizedReviewers ?? new List<string>()).ToHashSet(StringComparer.Ordinal),
        };
    }

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow,
    };

    private sealed class PolicyDto
    {
        public string? Version { get; set; }
        public EvaluatorDto? Evaluator { get; set; }
        public ThresholdsDto? Thresholds { get; set; }
        public List<string>? AuthorizedReviewers { get; set; }
    }

    private sealed class EvaluatorDto
    {
        public string? Model { get; set; }
        public int TimeoutSeconds { get; set; }
        public int MaxEvidenceCharacters { get; set; }
    }

    private sealed class ThresholdsDto
    {
        public double MinEvidenceConfidence { get; set; }
        public double MinResultConfidence { get; set; }
        public double MaterialityProposalScore { get; set; }
        public double MinMaterialityConfidence { get; set; }
    }
}

/// <summary>What the workflow does next with a requirement.</summary>
public enum Route
{
    /// <summary>The finding is settled enough to continue: record a pass, or propose a change for approval.</summary>
    Proceed,

    /// <summary>Stop the run. Nothing is proposed or written.</summary>
    Stop,

    /// <summary>A person decides. No change is proposed for this requirement.</summary>
    NeedsHumanReview,
}

/// <summary>The requirement's outcome as far as this run can say.</summary>
public enum Outcome
{
    /// <summary>The evidence shows the requirement is met.</summary>
    Satisfied,

    /// <summary>The evidence shows the requirement is not met.</summary>
    Violated,

    /// <summary>This run cannot say.</summary>
    NotDetermined,
}

/// <summary>What the host may put in Word for a requirement.</summary>
public enum WordAction
{
    /// <summary>Nothing.</summary>
    None,

    /// <summary>A host-written "needs human review" comment on the section heading.</summary>
    ReviewComment,

    /// <summary>A comment on the evidence.</summary>
    Comment,

    /// <summary>A comment plus one tracked change.</summary>
    CommentAndTrackedChange,
}

/// <summary>A routing decision and the rule that produced it.</summary>
/// <param name="Route">The route.</param>
/// <param name="Outcome">The outcome.</param>
/// <param name="Action">What the host may put in Word.</param>
/// <param name="Rule">The id of the routing rule that fired.</param>
/// <param name="Reason">The technical reason, recorded in the audit.</param>
/// <param name="Note">The same reason in plain words, for comments a document author reads.</param>
public sealed record RoutingDecision(Route Route, Outcome Outcome, WordAction Action, string Rule, string Reason, string Note);

/// <summary>
/// Turns typed answers into a route. Ordered rules, first match wins, and every path that is
/// not an explicit, confident, uncontradicted pass ends somewhere other than a pass.
/// </summary>
public static class Router
{
    /// <summary>Routes a deterministic requirement.</summary>
    /// <param name="requirement">The requirement.</param>
    /// <param name="results">Its check results.</param>
    /// <returns>The routing decision.</returns>
    public static RoutingDecision RouteHardRule(RegisteredRequirement requirement, IReadOnlyList<DeterministicResult> results)
    {
        if (results.Count > 0 && results.All(r => r.Passed))
        {
            return new RoutingDecision(Route.Proceed, Outcome.Satisfied, WordAction.None, "H1", "all deterministic checks passed", "All mandatory elements are present.");
        }

        var failed = string.Join("; ", results.Where(r => !r.Passed).Select(r => $"{r.Check}: {r.Detail}"));
        return requirement.OnFail == HardRuleFailure.Stop
            ? new RoutingDecision(Route.Stop, Outcome.Violated, WordAction.None, "H2", $"hard rule failed ({failed}); run stopped before any external call", "A mandatory element is missing, so the review stopped.")
            : new RoutingDecision(Route.NeedsHumanReview, Outcome.Violated, WordAction.ReviewComment, "H3", $"hard rule failed ({failed})", "A mandatory element is missing.");
    }

    /// <summary>Routes a semantic requirement from its typed answers.</summary>
    /// <param name="policy">The policy.</param>
    /// <param name="requirement">The requirement.</param>
    /// <param name="decision">The evaluator's decision.</param>
    /// <param name="passSignals">Deterministic signals that can block a pass.</param>
    /// <returns>The routing decision.</returns>
    public static RoutingDecision RouteSemantic(
        ReviewPolicy policy,
        RegisteredRequirement requirement,
        RequirementDecision decision,
        IReadOnlyList<DeterministicResult> passSignals)
    {
        var t = policy.Thresholds;

        if (decision.Failure is { } failure)
        {
            return Human("R1", $"evaluator failed: {failure.Kind} ({failure.Detail})", "The automated check could not be completed.");
        }

        if (!decision.IsComplete)
        {
            return Human("R2", "evaluator returned incomplete answers", "The automated check could not be completed.");
        }

        var evidence = decision.EvidenceStatus!;
        var result = decision.Result!;
        var materiality = decision.Materiality!;

        if (evidence.Label != "sufficient")
        {
            return Human("R3", $"evidence_status = {evidence.Label}", evidence.Label == "conflicting"
                ? "This section contains statements that contradict each other."
                : "This section does not contain enough information to decide.");
        }

        if (evidence.Confidence < t.MinEvidenceConfidence)
        {
            return Human("R4", $"evidence_status confidence {F(evidence.Confidence)} < {F(t.MinEvidenceConfidence)}", "The automated check was not confident enough to decide.");
        }

        if (result.Label == "not_determined")
        {
            return Human("R5", "requirement_result = not_determined", "The automated check could not decide.");
        }

        if (result.Confidence < t.MinResultConfidence)
        {
            return Human("R6", $"requirement_result confidence {F(result.Confidence)} < {F(t.MinResultConfidence)}", "The automated check was not confident enough to decide.");
        }

        if (result.Label == "satisfied")
        {
            var blocking = passSignals.Where(s => !s.Passed).ToList();
            if (blocking.Count > 0)
            {
                // Disagreement: the evaluator says met, a registered deterministic signal
                // says a required element is absent. Neither side wins; a person decides.
                return Human("R7", $"disagreement: evaluator says satisfied, {string.Join(", ", blocking.Select(b => b.Check))} found no match", "The automated check and a keyword check disagree.");
            }

            return new RoutingDecision(Route.Proceed, Outcome.Satisfied, WordAction.None, "R8",
                $"satisfied with confidence {F(result.Confidence)}; evidence sufficient with confidence {F(evidence.Confidence)}",
                "Meets the requirement.");
        }

        // violated. Uncertain materiality never downgrades a violation.
        var trusted = materiality.Confidence >= t.MinMaterialityConfidence;
        var material = !trusted || materiality.Score >= t.MaterialityProposalScore;
        var why = trusted
            ? $"materiality {F(materiality.Score)} {(material ? ">=" : "<")} {F(t.MaterialityProposalScore)}"
            : $"materiality confidence {F(materiality.Confidence)} < {F(t.MinMaterialityConfidence)}, treated as material";

        var action = material && requirement.Proposal == ProposalMode.TrackedChange
            ? WordAction.CommentAndTrackedChange
            : WordAction.Comment;

        return new RoutingDecision(Route.Proceed, Outcome.Violated, action, material ? "R9" : "R10",
            $"violated with confidence {F(result.Confidence)}; {why}",
            "Does not meet the requirement.");
    }

    /// <summary>Routes a requirement whose evidence failed the completeness check.</summary>
    /// <param name="rule">The rule id.</param>
    /// <param name="reason">The technical reason.</param>
    /// <param name="note">The reason in plain words.</param>
    /// <returns>The routing decision.</returns>
    public static RoutingDecision Incomplete(string rule, string reason, string note) => Human(rule, reason, note);

    private static RoutingDecision Human(string rule, string reason, string note) =>
        new(Route.NeedsHumanReview, Outcome.NotDetermined, WordAction.ReviewComment, rule, reason, note);

    private static string F(double value) => value.ToString("0.00", CultureInfo.InvariantCulture);
}
