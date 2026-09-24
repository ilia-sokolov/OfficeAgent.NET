using System.Xml.Linq;
using OfficeAgent.Abstractions;
using OfficeAgent.Samples.RequirementsReview;
using Xunit;

namespace RequirementsReview.Tests;

/// <summary>
/// The demonstration run: one supplier method statement, one requirement set, three outcomes,
/// through the real Jev adapter (scripted transport), the Agent Framework drafter (scripted
/// model), and OfficeAgent.
/// </summary>
public sealed class PipelineTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string Output = "method-statement.reviewed.docx";
    private const string Gap = "the team will assess the situation and agree next steps";

    [Fact]
    public async Task One_run_shows_a_pass_a_violation_and_a_human_review()
    {
        using var harness = await Harness.CreateAsync();

        var proposal = await harness.ProposeAsync();

        Assert.Equal(ProposalStatus.AwaitingApproval, proposal.Status);

        var gov = Item(proposal, "REQ-GOV-01");
        Assert.Equal((Route.Proceed, Outcome.Satisfied, WordAction.None), (gov.Routing.Route, gov.Routing.Outcome, gov.Routing.Action));
        Assert.All(gov.Deterministic, d => Assert.True(d.Passed));

        var pass = Item(proposal, "REQ-INS-01");
        Assert.Equal((Route.Proceed, Outcome.Satisfied, "R8"), (pass.Routing.Route, pass.Routing.Outcome, pass.Routing.Rule));
        Assert.Equal(new[] { "w14:1000000D", "w14:1000000E" }, pass.Evidence.Items.Select(i => i.ParaId));
        Assert.Equal(MethodStatementFixture.HoldPoint2, pass.Evidence.Items[1].Text);

        var fail = Item(proposal, "REQ-REC-01");
        Assert.Equal((Route.Proceed, Outcome.Violated, WordAction.Comment), (fail.Routing.Route, fail.Routing.Outcome, fail.Routing.Action));
        var draft = Assert.IsType<Proposal>(fail.Draft?.Proposal);
        Assert.Equal(("REQ-REC-01:1", "w14:10000010", Gap), (draft.EvidenceId, draft.ParaId, draft.TargetText));
        Assert.Null(draft.ReplacementText);

        var human = Item(proposal, "REQ-RSP-01");
        Assert.Equal((Route.NeedsHumanReview, Outcome.NotDetermined, "R3"), (human.Routing.Route, human.Routing.Outcome, human.Routing.Rule));
        Assert.Null(human.Draft);

        // A supplier owns its text: the customer comments, it does not redline.
        var ops = proposal.Built!.Plan.Operations;
        Assert.All(ops, o => Assert.IsType<CommentOp>(o));
        Assert.Equal(2, ops.Count);
        Assert.Equal(proposal.Evidence.Snapshot.ETag, proposal.Built.Plan.Snapshot!.ETag);

        // Nothing is written until a reviewer approves.
        Assert.Equal(1, harness.DocxCount());
    }

    [Fact]
    public async Task Approved_proposal_lands_as_comments_in_a_new_copy()
    {
        using var harness = await Harness.CreateAsync();
        var sourceHash = harness.SourceSha256();

        var proposal = await harness.ProposeAsync();
        var commit = await harness.ApproveAsync(proposal);

        Assert.Equal(CommitStatus.Committed, commit.Status);
        Assert.Equal(sourceHash, harness.SourceSha256());

        var doc = XDocument.Parse(harness.DocumentXml(Output));
        Assert.DoesNotContain(doc.Descendants(), e => (string?)e.Attribute(W + "author") == PlanBuilder.Author);

        var comments = XDocument.Parse(harness.CommentsXml(Output)).Descendants(W + "comment").ToList();
        Assert.Equal(3, comments.Count);
        Assert.Contains(comments, c => (string?)c.Attribute(W + "author") == "Maria Jensen" && Text(c, "t") == MethodStatementFixture.ExistingComment);
        Assert.Contains(comments, c => (string?)c.Attribute(W + "author") == PlanBuilder.Author && Text(c, "t").StartsWith("[REQ-REC-01 v1] Please list the recovery actions"));
        Assert.Contains(comments, c => (string?)c.Attribute(W + "author") == PlanBuilder.Author
            && Text(c, "t") == "[REQ-RSP-01 v1] Needs human review. This section does not contain enough information to decide. No change proposed.");

        // The drafted comment is anchored on the vague sentence itself.
        var recovery = Paragraph(doc, "10000010");
        Assert.Single(recovery.Descendants(W + "commentRangeStart"));
    }

    [Fact]
    public async Task Existing_table_comment_and_tracked_change_survive_the_commit()
    {
        using var harness = await Harness.CreateAsync();

        var proposal = await harness.ProposeAsync();
        Assert.Equal((1, 1, 1), (proposal.Evidence.ExistingCommentCount, proposal.Evidence.ExistingRevisionCount, proposal.Evidence.TableCount));

        await harness.ApproveAsync(proposal);
        var doc = XDocument.Parse(harness.DocumentXml(Output));

        var table = Assert.Single(doc.Descendants(W + "tbl"));
        Assert.Equal(3, table.Elements(W + "tr").Count());
        Assert.Contains("Priya Raman", table.Descendants(W + "t").Select(t => t.Value));

        var theirs = doc.Descendants(W + "ins").Where(e => (string?)e.Attribute(W + "author") == "Tom Berger").ToList();
        Assert.Equal(MethodStatementFixture.ExistingInsertion, Text(Assert.Single(theirs), "t"));

        // Maria's comment is still anchored where it was.
        Assert.Contains(Paragraph(doc, "10000007").Descendants(W + "commentRangeStart"), c => (string?)c.Attribute(W + "id") == "0");
    }

    [Fact]
    public async Task A_requirement_that_allows_tracked_changes_gets_a_redline_after_the_comment()
    {
        // Internal documents can allow proposed wording. Same pipeline, one registry switch.
        using var harness = await Harness.CreateAsync();
        harness.Requirements = Variants.RecoveryAllowsTrackedChange(harness.Requirements);
        var chat = new ScriptedChatClient(id => id == "REQ-REC-01"
            ? new ScriptedCall(AgentProposalDrafter.ToolName, new Dictionary<string, object?>
            {
                ["evidenceId"] = "REQ-REC-01:1",
                ["targetText"] = Gap,
                ["replacementText"] = "the site supervisor makes the unit safe and reinstates the existing unit",
                ["commentText"] = "Recovery actions and their owner must be explicit.",
            })
            : null);

        var proposal = await harness.ProposeAsync(chat: chat);

        Assert.Equal(WordAction.CommentAndTrackedChange, Item(proposal, "REQ-REC-01").Routing.Action);
        Assert.Equal(new[] { typeof(CommentOp), typeof(CommentOp), typeof(ChangeTextOp) }, proposal.Built!.Plan.Operations.Select(o => o.GetType()));

        Assert.Equal(CommitStatus.Committed, (await harness.ApproveAsync(proposal)).Status);
        var ours = XDocument.Parse(harness.DocumentXml(Output)).Descendants().Where(e => (string?)e.Attribute(W + "author") == PlanBuilder.Author).ToList();
        Assert.Contains(ours, e => e.Name == W + "del" && Text(e, "delText") == Gap);
        Assert.Contains(ours, e => e.Name == W + "ins" && Text(e, "t") == "the site supervisor makes the unit safe and reinstates the existing unit");
    }

    [Fact]
    public async Task All_requirements_passing_writes_nothing_but_still_checks_freshness()
    {
        // The recovery text must name a role, or REQ-REC-01's deterministic signal blocks the pass.
        using var harness = await Harness.CreateAsync(new MethodStatementOptions
        {
            Recovery = "If installation fails, the site supervisor makes the unit safe and reinstates the existing unit within 24 hours.",
        });

        var proposal = await harness.ProposeAsync(Script.Always(r => Script.Decision(r.Id)));

        Assert.Equal(ProposalStatus.NothingToPropose, proposal.Status);
        Assert.Empty(proposal.Built!.Plan.Operations);
        Assert.True(proposal.Preview!.IsValid);
        Assert.Equal(proposal.Evidence.SourceSha256, proposal.PreviewReceipt!.InputSha256);
        Assert.Equal(1, harness.DocxCount());

        var commit = await harness.ApproveAsync(proposal);
        Assert.Equal(CommitStatus.NotCommittable, commit.Status);
    }

    internal static ReviewItem Item(ReviewProposal proposal, string id) => proposal.Items.Single(i => i.Requirement.Id == id);

    private static XElement Paragraph(XDocument doc, string paraId) =>
        doc.Descendants(W + "p").Single(p => (string?)p.Attribute(XNamespace.Get("http://schemas.microsoft.com/office/word/2010/wordml") + "paraId") == paraId);

    private static string Text(XElement element, string local) => string.Concat(element.Descendants(W + local).Select(t => t.Value));
}

/// <summary>Requirement-set variants for tests.</summary>
internal static class Variants
{
    /// <summary>The bundled set with REQ-REC-01 allowed to propose a tracked change.</summary>
    public static RequirementSet RecoveryAllowsTrackedChange(RequirementSet _)
    {
        var json = File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "requirements.json"));
        const string from = "\"section\": \"Recovery procedure\",\n      \"passRequires\": [ \"\\\\b(responsible|owner|supervisor|manager|engineer)\\\\b\" ],\n      \"proposal\": \"commentOnly\"";
        var normalized = json.Replace("\r\n", "\n", StringComparison.Ordinal);
        Assert.Contains(from, normalized);
        return RequirementSet.FromJson(normalized.Replace(from, from.Replace("commentOnly", "trackedChange", StringComparison.Ordinal), StringComparison.Ordinal));
    }
}
