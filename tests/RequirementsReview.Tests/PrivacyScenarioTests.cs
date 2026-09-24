using System.Net;
using OfficeAgent.Abstractions;
using OfficeAgent.Samples.RequirementsReview;
using Xunit;

namespace RequirementsReview.Tests;

public sealed class PrivacyScenarioTests
{
    [Fact]
    public async Task Supplier_proposal_is_checked_by_MAF_tool_and_approved_comments_land_in_new_copy()
    {
        using var harness = await Harness.CreatePrivacyAsync();
        var sourceHash = harness.SourceSha256();
        var transport = new ScriptedJevHandler(id => ScriptedAnswer.PrivacyDemo.TryGetValue(id, out var answer)
            ? ScriptedJevHandler.Json(HttpStatusCode.OK, answer.ToResponse(ScriptedJevHandler.ScriptedModel).ToJsonString())
            : ScriptedJevHandler.Json(HttpStatusCode.InternalServerError, "{\"detail\":\"no scripted answer\"}"));
        var reviewModel = new ScriptedReviewChatClient();

        var proposal = await harness.ProposeAsync(evaluator: harness.Jev(transport), reviewChat: reviewModel);

        Assert.Equal(ProposalStatus.AwaitingApproval, proposal.Status);
        Assert.Equal(3, transport.Bodies.Count);
        Assert.All(reviewModel.OfferedTools, tools => Assert.Equal(new[] { AgentRequirementReviewer.ToolName }, tools));
        Assert.All(proposal.Items, item =>
        {
            Assert.Equal(Route.Proceed, item.Routing.Route);
            Assert.Equal(Outcome.Violated, item.Routing.Outcome);
            Assert.NotNull(item.Draft?.Proposal);
        });
        Assert.Equal(3, proposal.Built!.Plan.Operations.Count);
        Assert.All(proposal.Built.Plan.Operations, operation => Assert.IsType<CommentOp>(operation));
        Assert.Equal(1, harness.DocxCount());

        const string reviewed = "supplier-proposal.reviewed.docx";
        var commit = await harness.ApproveAsync(proposal, reviewed);

        Assert.Equal(CommitStatus.Committed, commit.Status);
        Assert.True(harness.Exists(reviewed));
        Assert.Equal(sourceHash, harness.SourceSha256());
        var comments = harness.CommentsXml(reviewed);
        Assert.Contains("REQ-PRIV-ACCESS", comments);
        Assert.Contains("REQ-PRIV-STORAGE", comments);
        Assert.Contains("REQ-PRIV-RETENTION", comments);
    }
}
