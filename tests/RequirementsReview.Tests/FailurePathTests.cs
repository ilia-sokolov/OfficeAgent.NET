using System.Net;
using DocumentFormat.OpenXml.Packaging;
using OfficeAgent.Abstractions;
using OfficeAgent.Samples.RequirementsReview;
using Xunit;
using static RequirementsReview.Tests.PipelineTests;

namespace RequirementsReview.Tests;

/// <summary>
/// Missing evidence, provider failure, disagreement, stale snapshots and unsupported document
/// state. None of them may produce a pass, and none of them may write.
/// </summary>
public sealed class FailurePathTests
{
    private static readonly string[] Semantic = { "REQ-INS-01", "REQ-REC-01", "REQ-RSP-01" };

    [Fact]
    public async Task A_provider_outage_turns_every_semantic_requirement_into_human_review()
    {
        using var harness = await Harness.CreateAsync();
        var handler = new ScriptedJevHandler(_ => ScriptedJevHandler.Json(HttpStatusCode.ServiceUnavailable, "{}"));

        var proposal = await harness.ProposeAsync(harness.Jev(handler));

        foreach (var id in Semantic)
        {
            var item = Item(proposal, id);
            Assert.Equal((Route.NeedsHumanReview, Outcome.NotDetermined, "R1"), (item.Routing.Route, item.Routing.Outcome, item.Routing.Rule));
        }

        Assert.DoesNotContain(proposal.Built!.Plan.Operations, o => o is ChangeTextOp);
        Assert.Equal(1, harness.DocxCount());
    }

    [Fact]
    public async Task An_evaluator_that_throws_cannot_produce_a_pass()
    {
        using var harness = await Harness.CreateAsync();
        var evaluator = new ScriptedRequirementEvaluator((_, _, _) => throw new InvalidOperationException("boom"));

        var proposal = await harness.ProposeAsync(evaluator);

        Assert.All(Semantic, id => Assert.Equal("R1", Item(proposal, id).Routing.Rule));
        Assert.Equal(EvaluationFailureKind.EvaluatorError, Item(proposal, "REQ-INS-01").Decision!.Failure!.Kind);
    }

    [Fact]
    public async Task A_decision_for_another_requirement_is_refused()
    {
        using var harness = await Harness.CreateAsync();
        var evaluator = Script.Always(_ => Script.Decision("REQ-INS-01"));

        var proposal = await harness.ProposeAsync(evaluator);

        Assert.Equal(Outcome.Satisfied, Item(proposal, "REQ-INS-01").Routing.Outcome);
        Assert.Equal(EvaluationFailureKind.MalformedResponse, Item(proposal, "REQ-REC-01").Decision!.Failure!.Kind);
        Assert.Equal(Route.NeedsHumanReview, Item(proposal, "REQ-REC-01").Routing.Route);
    }

    [Fact]
    public async Task Malformed_provider_output_goes_to_a_person()
    {
        using var harness = await Harness.CreateAsync();
        var handler = new ScriptedJevHandler(_ => Script.Ok("{\"model\":\"jev\",\"answers\":{\"compliant\":true}}"));

        var proposal = await harness.ProposeAsync(harness.Jev(handler));

        Assert.All(Semantic, id => Assert.Equal(EvaluationFailureKind.MalformedResponse, Item(proposal, id).Decision!.Failure!.Kind));
        Assert.DoesNotContain(proposal.Items, i => i.Requirement.Kind == RequirementKind.Semantic && i.Routing.Outcome == Outcome.Satisfied);
    }

    [Fact]
    public async Task A_missing_section_is_not_sent_and_not_passed()
    {
        using var harness = await Harness.CreateAsync(new MethodStatementOptions { Responsibilities = null });
        var handler = new ScriptedJevHandler();

        var proposal = await harness.ProposeAsync(harness.Jev(handler));

        var item = Item(proposal, "REQ-RSP-01");
        Assert.Equal(("C1", Route.NeedsHumanReview), (item.Routing.Rule, item.Routing.Route));
        Assert.Null(item.Decision);
        Assert.DoesNotContain(handler.Bodies, b => b.Contains("REQ-RSP-01", StringComparison.Ordinal));
        Assert.Contains(proposal.Built!.Changes, c => c.RequirementId == "REQ-RSP-01" && c.Operations.Count == 0 && c.Note is not null);
    }

    [Fact]
    public async Task An_empty_section_is_not_sent_and_not_passed()
    {
        using var harness = await Harness.CreateAsync(new MethodStatementOptions { Responsibilities = "" });
        var handler = new ScriptedJevHandler();

        var proposal = await harness.ProposeAsync(harness.Jev(handler));

        var item = Item(proposal, "REQ-RSP-01");
        Assert.Equal(("C3", Route.NeedsHumanReview), (item.Routing.Rule, item.Routing.Route));
        Assert.DoesNotContain(handler.Bodies, b => b.Contains("REQ-RSP-01", StringComparison.Ordinal));
    }

    [Fact]
    public async Task An_ambiguous_section_is_not_guessed()
    {
        using var harness = await Harness.CreateAsync(new MethodStatementOptions { DuplicateInspectionsHeading = true });
        var handler = new ScriptedJevHandler();

        var proposal = await harness.ProposeAsync(harness.Jev(handler));

        Assert.Equal("C2", Item(proposal, "REQ-INS-01").Routing.Rule);
        Assert.DoesNotContain(handler.Bodies, b => b.Contains("REQ-INS-01", StringComparison.Ordinal));
    }

    [Fact]
    public async Task Evidence_with_a_pending_tracked_change_is_unsupported_state()
    {
        // Inspection would show the supplier's pending insertion as if accepted. That is a
        // proposal, not the submitted text, so it is not judged at all.
        using var harness = await Harness.CreateAsync(new MethodStatementOptions { PendingRevisionInRecovery = true });
        var handler = new ScriptedJevHandler();

        var proposal = await harness.ProposeAsync(harness.Jev(handler));

        var item = Item(proposal, "REQ-REC-01");
        Assert.Equal(("C5", Route.NeedsHumanReview), (item.Routing.Rule, item.Routing.Route));
        Assert.True(item.Evidence.Items[0].HasPendingRevision);
        Assert.EndsWith("agree next steps." + MethodStatementFixture.PendingRecoveryInsertion, item.Evidence.Items[0].Text);
        Assert.DoesNotContain(handler.Bodies, b => b.Contains("REQ-REC-01", StringComparison.Ordinal));

        // The review comment still lands, and both of Tom's pending insertions survive.
        var commit = await harness.ApproveAsync(proposal);
        Assert.Equal(CommitStatus.Committed, commit.Status);
        var xml = harness.DocumentXml("method-statement.reviewed.docx");
        Assert.Contains(MethodStatementFixture.PendingRecoveryInsertion.Trim(), xml);
        Assert.Contains(MethodStatementFixture.ExistingInsertion, xml);
    }

    [Fact]
    public async Task Evidence_over_the_budget_is_routed_before_any_call()
    {
        using var harness = await Harness.CreateAsync();
        var handler = new ScriptedJevHandler();
        var tight = harness.Policy with { MaxEvidenceCharacters = 50 };

        var proposal = await harness.ProposeAsync(harness.Jev(handler), policy: tight);

        Assert.All(Semantic, id => Assert.Equal("C6", Item(proposal, id).Routing.Rule));
        Assert.Empty(handler.Bodies);
    }

    [Fact]
    public async Task A_failed_hard_rule_stops_before_any_content_leaves_the_host()
    {
        using var harness = await Harness.CreateAsync(new MethodStatementOptions { InstallationWindow = "Installation window: to be confirmed." });
        var handler = new ScriptedJevHandler();
        var chat = new ScriptedChatClient();

        var proposal = await harness.ProposeAsync(harness.Jev(handler), chat);

        Assert.Equal(ProposalStatus.Stopped, proposal.Status);
        Assert.Equal((Route.Stop, "H2"), (Item(proposal, "REQ-GOV-01").Routing.Route, Item(proposal, "REQ-GOV-01").Routing.Rule));
        Assert.Empty(handler.Bodies);
        Assert.Empty(chat.Transcript);
        Assert.Null(proposal.Built);
        Assert.Equal(1, harness.DocxCount());
    }

    [Fact]
    public async Task An_edit_during_evaluation_makes_the_proposal_stale()
    {
        using var harness = await Harness.CreateAsync();
        var evaluator = new ScriptedRequirementEvaluator(async (r, _, ct) =>
        {
            if (r.Id == "REQ-INS-01")
            {
                // A colleague edits an unrelated paragraph while the evaluation runs.
                await File.WriteAllBytesAsync(harness.SourcePath, MethodStatementFixture.Create(new MethodStatementOptions { Scope = "This method statement covers replacing AHU-3" }), ct);
            }

            return await Script.Demo().EvaluateAsync(r, null!, ct);
        });

        var proposal = await harness.ProposeAsync(evaluator);

        Assert.Equal(ProposalStatus.Stale, proposal.Status);
        Assert.Contains(ValidationErrorCodes.StaleSnapshot, proposal.StatusReason);
        Assert.Equal(1, harness.DocxCount());
        Assert.Equal(CommitStatus.NotCommittable, (await harness.ApproveAsync(proposal)).Status);
    }

    [Fact]
    public async Task A_change_the_snapshot_does_not_cover_is_caught_by_the_input_hash()
    {
        // Editing only an existing comment's text leaves every text host, and so the
        // snapshot, unchanged. The preview receipt's input hash still differs.
        using var harness = await Harness.CreateAsync();
        var evaluator = new ScriptedRequirementEvaluator(async (r, _, ct) =>
        {
            if (r.Id == "REQ-INS-01")
            {
                using var doc = WordprocessingDocument.Open(harness.SourcePath, isEditable: true);
                var text = doc.MainDocumentPart!.WordprocessingCommentsPart!.Comments!.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().First();
                text.Text = "Lag threshold confirmed.";
                doc.MainDocumentPart.WordprocessingCommentsPart.Comments.Save();
            }

            return await Script.Demo().EvaluateAsync(r, null!, ct);
        });

        var proposal = await harness.ProposeAsync(evaluator);

        Assert.Equal(ProposalStatus.Stale, proposal.Status);
        Assert.Equal("document bytes changed after evidence was collected", proposal.StatusReason);
        Assert.Equal(1, harness.DocxCount());
    }

    [Fact]
    public async Task A_drafter_failure_keeps_the_violation_and_drops_only_the_wording()
    {
        using var harness = await Harness.CreateAsync();
        var chat = new ThrowingChatClient();

        var proposal = await harness.ProposeAsync(chat: chat);

        var item = Item(proposal, "REQ-REC-01");
        Assert.Equal(Outcome.Violated, item.Routing.Outcome);
        Assert.Null(item.Draft!.Proposal);
        Assert.StartsWith("drafting failed", item.Draft.Error);
        Assert.DoesNotContain(proposal.Built!.Plan.Operations, o => o is ChangeTextOp);
        Assert.Contains(proposal.Built.Plan.Operations.OfType<CommentOp>(), c => c.Text.Contains("No valid wording was drafted", StringComparison.Ordinal));
    }

    [Fact]
    public void A_runaway_pattern_fails_its_check_instead_of_hanging()
    {
        var requirement = new RegisteredRequirement
        {
            Id = "REQ-X",
            Version = "1",
            Kind = RequirementKind.Deterministic,
            Title = "x",
            Text = "x",
            Section = "s",
            Patterns = new[] { new System.Text.RegularExpressions.Regex("(a+)+$", System.Text.RegularExpressions.RegexOptions.None, RequirementSet.PatternTimeout) },
        };
        var evidence = new AnchoredDocumentEvidence("REQ-X", "s", 1, "h", "s", new[] { new EvidenceItem("REQ-X:1", "p", new string('a', 40) + "!", null, false, false) });

        var result = Assert.Single(HardRuleEvaluator.Evaluate(requirement, evidence));

        Assert.False(result.Passed);
        Assert.Equal("pattern timed out", result.Detail);
    }
}

/// <summary>A chat client whose model is down.</summary>
internal sealed class ThrowingChatClient : Microsoft.Extensions.AI.IChatClient
{
    public Task<Microsoft.Extensions.AI.ChatResponse> GetResponseAsync(IEnumerable<Microsoft.Extensions.AI.ChatMessage> messages, Microsoft.Extensions.AI.ChatOptions? options = null, CancellationToken cancellationToken = default) =>
        throw new HttpRequestException("model endpoint unavailable");

    public IAsyncEnumerable<Microsoft.Extensions.AI.ChatResponseUpdate> GetStreamingResponseAsync(IEnumerable<Microsoft.Extensions.AI.ChatMessage> messages, Microsoft.Extensions.AI.ChatOptions? options = null, CancellationToken cancellationToken = default) =>
        throw new HttpRequestException("model endpoint unavailable");

    public object? GetService(Type serviceType, object? serviceKey = null) => null;

    public void Dispose()
    {
    }
}
