using OfficeAgent.Abstractions;
using OfficeAgent.Samples.RequirementsReview;
using Xunit;
using static RequirementsReview.Tests.PipelineTests;

namespace RequirementsReview.Tests;

/// <summary>
/// The Agent Framework drafter's boundary: one tool, no view of the policy, and arguments
/// treated as untrusted input.
/// </summary>
public sealed class DrafterTests
{
    private static ScriptedChatClient Proposing(string evidenceId, string target, string replacement = "", string comment = "Please name the recovery actions and their owner.") =>
        new(id => id == "REQ-REC-01"
            ? new ScriptedCall(AgentProposalDrafter.ToolName, new Dictionary<string, object?>
            {
                ["evidenceId"] = evidenceId,
                ["targetText"] = target,
                ["replacementText"] = replacement,
                ["commentText"] = comment,
            })
            : null);

    [Fact]
    public async Task The_agent_is_offered_exactly_one_tool_and_never_sees_the_policy()
    {
        using var harness = await Harness.CreateAsync();
        var chat = new ScriptedChatClient();

        await harness.ProposeAsync(chat: chat);

        Assert.NotEmpty(chat.OfferedTools);
        Assert.All(chat.OfferedTools, tools => Assert.Equal(new[] { AgentProposalDrafter.ToolName }, tools));

        var seen = string.Join("\n", chat.Transcript);
        Assert.Contains("REQ-REC-01", seen);
        Assert.DoesNotContain(harness.Policy.Version, seen);
        Assert.DoesNotContain("0.80", seen);
        Assert.DoesNotContain("confidence", seen, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("REQ-RSP-01", seen);
    }

    [Fact]
    public async Task Only_violations_reach_the_drafter()
    {
        using var harness = await Harness.CreateAsync();
        var chat = new ScriptedChatClient();

        var proposal = await harness.ProposeAsync(chat: chat);

        Assert.Equal(new[] { "REQ-REC-01" }, proposal.Items.Where(i => i.Draft is not null).Select(i => i.Requirement.Id));
    }

    [Theory]
    [InlineData("REQ-INS-01:1", "the team will assess the situation and agree next steps", "evidenceId is not one of the supplied evidence ids")]
    [InlineData("w14:10000010", "the team will assess the situation and agree next steps", "evidenceId is not one of the supplied evidence ids")]
    [InlineData("REQ-REC-01:1", "if there is time", "targetText is not an exact substring of the cited evidence")]
    [InlineData("REQ-REC-01:1", "the", "targetText appears more than once in the cited evidence")]
    public async Task Invalid_citations_are_refused_and_nothing_is_redlined(string evidenceId, string target, string error)
    {
        using var harness = await Harness.CreateAsync();

        var proposal = await harness.ProposeAsync(chat: Proposing(evidenceId, target));

        var item = Item(proposal, "REQ-REC-01");
        Assert.Null(item.Draft!.Proposal);
        Assert.Equal(error, Assert.Single(item.Draft.Rejections));
        Assert.Equal(Outcome.Violated, item.Routing.Outcome);

        // The violation still reaches the document, as a host comment, not the refused text.
        var comments = proposal.Built!.Plan.Operations.OfType<CommentOp>().Select(c => c.Text).ToList();
        Assert.Contains(comments, c => c.Contains("No valid wording was drafted", StringComparison.Ordinal));
        Assert.DoesNotContain(comments, c => c.Contains("Please name the recovery actions", StringComparison.Ordinal));
    }

    [Fact]
    public async Task Overlong_comments_are_refused()
    {
        using var harness = await Harness.CreateAsync();

        var proposal = await harness.ProposeAsync(chat: Proposing("REQ-REC-01:1", "the team will assess the situation and agree next steps", comment: new string('x', ProposalCollector.MaxComment + 1)));

        Assert.Equal($"commentText must be 1-{ProposalCollector.MaxComment} characters", Assert.Single(Item(proposal, "REQ-REC-01").Draft!.Rejections));
    }

    [Fact]
    public async Task Replacement_wording_is_refused_for_a_supplier_owned_section()
    {
        using var harness = await Harness.CreateAsync();

        var proposal = await harness.ProposeAsync(chat: Proposing("REQ-REC-01:1", "the team will assess the situation and agree next steps", "the supplier does whatever it likes"));

        Assert.Equal("replacement wording is not allowed for this requirement", Assert.Single(Item(proposal, "REQ-REC-01").Draft!.Rejections));
        Assert.DoesNotContain(proposal.Built!.Plan.Operations, o => o is ChangeTextOp);
    }

    [Fact]
    public async Task A_call_to_a_tool_the_agent_was_not_given_does_nothing()
    {
        using var harness = await Harness.CreateAsync();
        var chat = new ScriptedChatClient(id => id == "REQ-REC-01"
            ? new ScriptedCall("apply_plan", new Dictionary<string, object?> { ["plan"] = "{}" })
            : null);

        var proposal = await harness.ProposeAsync(chat: chat);

        var item = Item(proposal, "REQ-REC-01");
        Assert.Null(item.Draft!.Proposal);
        Assert.DoesNotContain(proposal.Built!.Plan.Operations, o => o is ChangeTextOp);
        Assert.Equal(1, harness.DocxCount());
    }

    [Fact]
    public async Task Replacement_wording_is_refused_where_the_requirement_allows_comments_only()
    {
        var requirement = RoutingTestsData.Set.Requirements.Single(r => r.Id == "REQ-INS-01");
        var evidence = new AnchoredDocumentEvidence("REQ-INS-01", "Inspections", 1, "h", "Inspections",
            new[] { new EvidenceItem("REQ-INS-01:1", "w14:1", MethodStatementFixture.HoldPoint1, null, false, false) });
        var collector = new ProposalCollector(requirement, evidence, allowReplacement: false);
        var tool = collector.AsTool();

        var refused = await tool.InvokeAsync(new Microsoft.Extensions.AI.AIFunctionArguments(new Dictionary<string, object?>
        {
            ["evidenceId"] = "REQ-INS-01:1",
            ["targetText"] = "structural engineer",
            ["replacementText"] = "site engineer",
            ["commentText"] = "Different inspector.",
        }));

        Assert.Contains("replacement wording is not allowed", refused!.ToString());
        Assert.Null(collector.Accepted);
    }

    [Fact]
    public async Task Only_the_first_valid_submission_counts()
    {
        var requirement = RoutingTestsData.Set.Requirements.Single(r => r.Id == "REQ-REC-01");
        var evidence = new AnchoredDocumentEvidence("REQ-REC-01", "Recovery procedure", 1, "h", "Recovery procedure",
            new[] { new EvidenceItem("REQ-REC-01:1", "w14:10000010", MethodStatementFixture.RecoveryText, null, false, false) });
        var collector = new ProposalCollector(requirement, evidence, allowReplacement: true);
        var tool = collector.AsTool();

        async Task<object?> Submit(string replacement) => await tool.InvokeAsync(new Microsoft.Extensions.AI.AIFunctionArguments(new Dictionary<string, object?>
        {
            ["evidenceId"] = "REQ-REC-01:1",
            ["targetText"] = "the team will assess the situation and agree next steps",
            ["replacementText"] = replacement,
            ["commentText"] = "Make it mandatory.",
        }));

        await Submit("and verify it with a test restore");
        var second = await Submit("and skip it");

        Assert.Contains("already accepted", second!.ToString());
        Assert.Equal("and verify it with a test restore", collector.Accepted!.ReplacementText);
        Assert.StartsWith("prop_", collector.Accepted.ProposalId);
    }
}

internal static class RoutingTestsData
{
    public static readonly RequirementSet Set = RequirementSet.FromFile(Path.Combine(AppContext.BaseDirectory, "requirements.json"));
}
