using OfficeAgent.Samples.RequirementsReview;
using Xunit;

namespace RequirementsReview.Tests;

public sealed class AgentReviewerTests
{
    [Fact]
    public async Task Agent_calls_Jev_tool_for_each_registered_requirement()
    {
        var (requirements, evidence) = Inputs();
        var model = new ScriptedReviewChatClient();
        var jev = Script.Always(r => Script.Decision(r.Id, result: "violated"));

        var decisions = await new AgentRequirementReviewer(model, jev)
            .ReviewAsync(requirements, evidence, default);

        Assert.Equal(requirements.Select(r => r.Id), jev.Calls);
        Assert.All(model.OfferedTools, tools => Assert.Equal(new[] { AgentRequirementReviewer.ToolName }, tools));
        Assert.All(decisions.Values, decision => Assert.True(decision.IsComplete));
    }

    [Fact]
    public async Task Skipped_tool_call_is_not_treated_as_a_pass()
    {
        var (requirements, evidence) = Inputs();
        var model = new ScriptedReviewChatClient(ids => ids.Take(1).ToArray());
        var jev = Script.Always(r => Script.Decision(r.Id));

        var decisions = await new AgentRequirementReviewer(model, jev)
            .ReviewAsync(requirements, evidence, default);

        Assert.Equal(new[] { "REQ-PRIV-ACCESS" }, jev.Calls);
        Assert.Equal(EvaluationFailureKind.NotSent, decisions["REQ-PRIV-RETENTION"].Failure?.Kind);
    }

    [Fact]
    public async Task Duplicate_tool_call_requires_human_review_and_does_not_repeat_Jev_request()
    {
        var (requirements, evidence) = Inputs();
        var model = new ScriptedReviewChatClient(ids => new[] { ids[0], ids[0] });
        var jev = Script.Always(r => Script.Decision(r.Id));

        var decisions = await new AgentRequirementReviewer(model, jev)
            .ReviewAsync(requirements, evidence, default);

        Assert.Equal(new[] { "REQ-PRIV-ACCESS" }, jev.Calls);
        Assert.Equal(EvaluationFailureKind.MalformedResponse, decisions["REQ-PRIV-ACCESS"].Failure?.Kind);
        Assert.Equal(EvaluationFailureKind.NotSent, decisions["REQ-PRIV-RETENTION"].Failure?.Kind);
    }

    [Fact]
    public async Task Invented_ID_never_reaches_Jev()
    {
        var (requirements, evidence) = Inputs();
        var model = new ScriptedReviewChatClient(_ => new[] { "REQ-NOT-REGISTERED" });
        var jev = Script.Always(r => Script.Decision(r.Id));

        var decisions = await new AgentRequirementReviewer(model, jev)
            .ReviewAsync(requirements, evidence, default);

        Assert.Empty(jev.Calls);
        Assert.All(decisions.Values, decision => Assert.Equal(EvaluationFailureKind.MalformedResponse, decision.Failure?.Kind));
    }

    [Fact]
    public async Task Workflow_routes_requirements_skipped_by_agent_to_human_review()
    {
        using var harness = await Harness.CreateAsync();
        var model = new ScriptedReviewChatClient(ids => ids.Take(1).ToArray());

        var proposal = await harness.ProposeAsync(reviewChat: model);

        var skipped = proposal.Items.Single(item => item.Requirement.Id == "REQ-REC-01");
        Assert.Equal(EvaluationFailureKind.NotSent, skipped.Decision?.Failure?.Kind);
        Assert.Equal(Route.NeedsHumanReview, skipped.Routing.Route);
        Assert.Null(skipped.Draft);
    }

    private static (IReadOnlyList<RegisteredRequirement>, IReadOnlyDictionary<string, AnchoredDocumentEvidence>) Inputs()
    {
        var requirements = new[]
        {
            new RegisteredRequirement
            {
                Id = "REQ-PRIV-ACCESS", Version = "1", Kind = RequirementKind.Semantic,
                Title = "Access", Text = "Name allowed roles.", Section = "Data privacy",
            },
            new RegisteredRequirement
            {
                Id = "REQ-PRIV-RETENTION", Version = "1", Kind = RequirementKind.Semantic,
                Title = "Retention", Text = "Name a deletion date.", Section = "Data privacy",
            },
        };

        var evidence = requirements.ToDictionary(
            r => r.Id,
            r => new AnchoredDocumentEvidence(r.Id, "Data privacy", 1, "w14:30000002", "Data privacy",
                new[] { new EvidenceItem(r.Id + ":1", "w14:30000003", PrivacyProposalFixture.Passage, null, false, false) }),
            StringComparer.Ordinal);
        return (requirements, evidence);
    }
}
