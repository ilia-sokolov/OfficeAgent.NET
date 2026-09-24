using OfficeAgent.Samples.RequirementsReview;
using Xunit;

namespace RequirementsReview.Tests;

/// <summary>
/// The router in isolation. Every way to reach a pass is enumerated; nothing else gets there.
/// </summary>
public sealed class RoutingTests
{
    private static readonly ReviewPolicy Policy = ReviewPolicy.FromFile(Path.Combine(AppContext.BaseDirectory, "review-policy.json"));
    private static readonly RequirementSet Set = RequirementSet.FromFile(Path.Combine(AppContext.BaseDirectory, "requirements.json"));
    private static readonly RegisteredRequirement Backup = Set.Requirements.Single(r => r.Id == "REQ-REC-01");
    private static readonly RegisteredRequirement Tracked = Backup with { Proposal = ProposalMode.TrackedChange };
    private static readonly RegisteredRequirement Rollback = Set.Requirements.Single(r => r.Id == "REQ-INS-01");

    private static readonly IReadOnlyList<DeterministicResult> NoSignals = Array.Empty<DeterministicResult>();

    private static RoutingDecision Route(RequirementDecision decision, RegisteredRequirement? requirement = null, IReadOnlyList<DeterministicResult>? signals = null) =>
        Router.RouteSemantic(Policy, requirement ?? Backup, decision, signals ?? NoSignals);

    [Fact]
    public void A_confident_satisfied_answer_with_sufficient_evidence_passes()
    {
        var route = Route(Script.Decision("REQ-REC-01"));
        Assert.Equal((OfficeAgent.Samples.RequirementsReview.Route.Proceed, Outcome.Satisfied, "R8"), (route.Route, route.Outcome, route.Rule));
    }

    [Theory]
    [InlineData(EvaluationFailureKind.Timeout)]
    [InlineData(EvaluationFailureKind.ProviderUnavailable)]
    [InlineData(EvaluationFailureKind.Authentication)]
    [InlineData(EvaluationFailureKind.RateLimited)]
    [InlineData(EvaluationFailureKind.RequestRejected)]
    [InlineData(EvaluationFailureKind.MalformedResponse)]
    [InlineData(EvaluationFailureKind.NotSent)]
    [InlineData(EvaluationFailureKind.EvaluatorError)]
    public void Any_failure_goes_to_a_person_even_with_satisfied_answers_attached(EvaluationFailureKind kind)
    {
        // A failure wins over any answers a buggy evaluator also attached.
        var decision = Script.Decision("REQ-REC-01") with { Failure = new EvaluationFailure(kind, "x") };
        var route = Route(decision);
        Assert.Equal((OfficeAgent.Samples.RequirementsReview.Route.NeedsHumanReview, "R1"), (route.Route, route.Rule));
    }

    [Fact]
    public void Missing_answers_go_to_a_person()
    {
        var route = Route(Script.Decision("REQ-REC-01") with { Materiality = null });
        Assert.Equal("R2", route.Rule);
    }

    [Theory]
    [InlineData("insufficient")]
    [InlineData("conflicting")]
    public void Evidence_that_is_not_sufficient_goes_to_a_person(string label)
    {
        var route = Route(Script.Decision("REQ-REC-01", evidence: label));
        Assert.Equal(("R3", Outcome.NotDetermined), (route.Rule, route.Outcome));
    }

    [Fact]
    public void Low_evidence_confidence_goes_to_a_person() =>
        Assert.Equal("R4", Route(Script.Decision("REQ-REC-01", evidenceConfidence: 0.79)).Rule);

    [Fact]
    public void Not_determined_goes_to_a_person() =>
        Assert.Equal("R5", Route(Script.Decision("REQ-REC-01", result: "not_determined")).Rule);

    [Fact]
    public void Low_result_confidence_goes_to_a_person() =>
        Assert.Equal("R6", Route(Script.Decision("REQ-REC-01", resultConfidence: 0.79)).Rule);

    [Fact]
    public void A_deterministic_signal_can_block_a_pass()
    {
        var signals = new[] { new DeterministicResult("passRequires[0]", "\\brestor", false, Array.Empty<string>(), "no match in section") };
        var route = Route(Script.Decision("REQ-REC-01"), signals: signals);
        Assert.Equal((OfficeAgent.Samples.RequirementsReview.Route.NeedsHumanReview, "R7"), (route.Route, route.Rule));
        Assert.Contains("disagreement", route.Reason);
    }

    [Fact]
    public void A_deterministic_signal_cannot_produce_a_pass()
    {
        var signals = new[] { new DeterministicResult("passRequires[0]", "\\brestor", true, new[] { "REQ-REC-01:1" }, null) };
        var route = Route(Script.Decision("REQ-REC-01", result: "not_determined"), signals: signals);
        Assert.Equal(OfficeAgent.Samples.RequirementsReview.Route.NeedsHumanReview, route.Route);
    }

    [Fact]
    public void A_material_violation_gets_a_tracked_change_when_the_requirement_allows_one()
    {
        var route = Route(Script.Decision("REQ-REC-01", result: "violated", materiality: 2.1), Tracked);
        Assert.Equal((Outcome.Violated, WordAction.CommentAndTrackedChange, "R9"), (route.Outcome, route.Action, route.Rule));
    }

    [Fact]
    public void A_minor_violation_gets_a_comment_only()
    {
        var route = Route(Script.Decision("REQ-REC-01", result: "violated", materiality: 1.0), Tracked);
        Assert.Equal((Outcome.Violated, WordAction.Comment, "R10"), (route.Outcome, route.Action, route.Rule));
    }

    [Fact]
    public void Uncertain_materiality_never_downgrades_a_violation()
    {
        var route = Route(Script.Decision("REQ-REC-01", result: "violated", materiality: 0.2, materialityConfidence: 0.4), Tracked);
        Assert.Equal((WordAction.CommentAndTrackedChange, "R9"), (route.Action, route.Rule));
    }

    [Fact]
    public void A_comment_only_requirement_never_gets_a_tracked_change()
    {
        var route = Route(Script.Decision("REQ-INS-01", result: "violated", materiality: 3.0), Rollback);
        Assert.Equal(WordAction.Comment, route.Action);
    }

    [Fact]
    public void Nothing_but_an_explicit_confident_uncontradicted_answer_passes()
    {
        var labels = new[] { "sufficient", "conflicting", "insufficient" };
        var results = new[] { "satisfied", "violated", "not_determined" };
        var confidences = new[] { 0.5, 0.79, 0.8, 0.99 };
        var blocked = new[] { new DeterministicResult("passRequires[0]", "x", false, Array.Empty<string>(), "no match") };

        var passes = 0;
        foreach (var label in labels)
        foreach (var result in results)
        foreach (var ec in confidences)
        foreach (var rc in confidences)
        foreach (var failed in new[] { false, true })
        foreach (var signal in new[] { NoSignals, blocked })
        {
            var decision = Script.Decision("REQ-REC-01", label, ec, result, rc);
            if (failed)
            {
                decision = decision with { Failure = new EvaluationFailure(EvaluationFailureKind.Timeout, "x") };
            }

            var route = Route(decision, signals: signal);
            if (route.Outcome != Outcome.Satisfied)
            {
                continue;
            }

            passes++;
            Assert.False(failed);
            Assert.Equal(("sufficient", "satisfied"), (label, result));
            Assert.True(ec >= Policy.Thresholds.MinEvidenceConfidence && rc >= Policy.Thresholds.MinResultConfidence);
            Assert.Same(NoSignals, signal);
        }

        Assert.Equal(4, passes);
    }

    [Fact]
    public void Hard_rules_pass_stop_or_go_to_a_person()
    {
        var gov = Set.Requirements.Single(r => r.Id == "REQ-GOV-01");
        var ok = new[] { new DeterministicResult("pattern[0]", "x", true, new[] { "a" }, null) };
        var bad = new[] { new DeterministicResult("pattern[0]", "x", false, Array.Empty<string>(), "no match in section") };

        Assert.Equal(OfficeAgent.Samples.RequirementsReview.Route.Proceed, Router.RouteHardRule(gov, ok).Route);
        Assert.Equal(OfficeAgent.Samples.RequirementsReview.Route.Stop, Router.RouteHardRule(gov, bad).Route);
        Assert.Equal(OfficeAgent.Samples.RequirementsReview.Route.NeedsHumanReview, Router.RouteHardRule(gov with { OnFail = HardRuleFailure.HumanReview }, bad).Route);
        Assert.Equal(OfficeAgent.Samples.RequirementsReview.Route.Stop, Router.RouteHardRule(gov, Array.Empty<DeterministicResult>()).Route);
    }
}
