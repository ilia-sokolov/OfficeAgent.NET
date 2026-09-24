using System.Net;
using System.Text.Json.Nodes;
using OfficeAgent.Samples.RequirementsReview;
using Xunit;

namespace RequirementsReview.Tests;

/// <summary>
/// The audit record: complete enough to reconstruct every route, and free of secrets.
/// </summary>
public sealed class AuditTests
{
    private static readonly EvaluatorInfo Info = new("scripted-jev-wire", "test");

    private static async Task<(JsonObject Audit, ReviewProposal Proposal, CommitOutcome Commit, Harness Harness)> RunAsync()
    {
        var harness = await Harness.CreateAsync();
        var proposal = await harness.ProposeAsync();
        var commit = await harness.ApproveAsync(proposal);
        var json = AuditRecord.ToJson(AuditRecord.Build(proposal, Info, commit));
        return (JsonNode.Parse(json)!.AsObject(), proposal, commit, harness);
    }

    [Fact]
    public async Task The_record_holds_every_minimum_field()
    {
        var (audit, proposal, commit, harness) = await RunAsync();
        using var _h = harness;

        Assert.Equal(proposal.Evidence.SourceSha256, S(audit, "document.sha256"));
        Assert.Equal(proposal.Evidence.Snapshot.ETag, S(audit, "document.snapshotETag"));
        Assert.Equal("2026.4", S(audit, "requirementSet.version"));
        Assert.Equal("rp-2026.09.1", S(audit, "policy.version"));
        Assert.Equal(QuestionSet.Version, S(audit, "evaluator.questionSetVersion"));
        Assert.Equal(QuestionSet.RequirementResultLabels, audit["evaluator"]!["questions"]!["requirement_result"]!["criteria"]!.AsObject().Select(p => p.Key));

        var backup = Requirement(audit, "REQ-REC-01");
        Assert.Equal("1", S(backup, "version"));
        Assert.Equal("Does not meet the requirement.", S(backup, "note"));
        Assert.Equal("w14:10000010", S(backup["evidence"]![0]!, "paraId"));
        Assert.Equal(MethodStatementFixture.RecoveryText, S(backup["evidence"]![0]!, "excerpt"));
        Assert.Equal("violated", S(backup, "evaluation.answers.requirement_result.label"));
        Assert.NotNull(backup["evaluation"]!["answers"]!["requirement_result"]!["probabilities"]);
        Assert.StartsWith("req_scripted_", S(backup, "evaluation.requestId"));
        Assert.Equal(ScriptedJevHandler.ScriptedModel, S(backup, "evaluation.model"));
        Assert.Equal(ScriptedJevHandler.ProviderLabel, S(backup, "evaluation.provider"));
        Assert.Equal("Proceed", S(backup, "route"));
        Assert.Equal("R9", S(backup, "routingRule"));
        Assert.Equal(ScriptedChatClient.ScriptedModel, S(backup, "draft.model"));
        Assert.StartsWith("prop_", S(backup, "draft.proposal.proposalId"));

        var gov = Requirement(audit, "REQ-GOV-01");
        Assert.Equal(2, gov["deterministic"]!.AsArray().Count);

        Assert.Equal(proposal.PreviewReceipt!.PlanSha256, S(audit, "preview.receipt.planSha256"));
        Assert.Equal(Harness.Reviewer, S(audit, "humanDecision.reviewerId"));
        Assert.Equal("Approved", S(audit, "humanDecision.decision"));
        Assert.NotNull(audit["humanDecision"]!["decidedAtUtc"]);
        Assert.Equal("Committed", S(audit, "commit.status"));
        Assert.Equal(commit.Receipt!.OutputSha256, S(audit, "commit.receipt.outputSha256"));
        Assert.Equal(Harness.Reviewer, S(audit, "commit.receipt.actor"));
    }

    [Fact]
    public async Task Every_semantic_route_can_be_recomputed_from_the_record()
    {
        var (audit, _, _, harness) = await RunAsync();
        using var _h = harness;

        var t = audit["policy"]!["thresholds"]!;
        var policy = harness.Policy with
        {
            Thresholds = new Thresholds(
                D(t["minEvidenceConfidence"]),
                D(t["minResultConfidence"]),
                D(t["materialityProposalScore"]),
                D(t["minMaterialityConfidence"])),
        };

        var recomputed = 0;
        foreach (var node in audit["requirements"]!.AsArray())
        {
            if (node!["evaluation"]?["answers"] is not JsonObject answers)
            {
                continue;
            }

            var requirement = harness.Requirements.Requirements.Single(r => r.Id == S(node, "requirementId"));
            var decision = new RequirementDecision
            {
                RequirementId = requirement.Id,
                Provider = S(node, "evaluation.provider"),
                EvidenceStatus = Choice(answers["evidence_status"]!),
                Result = Choice(answers["requirement_result"]!),
                Materiality = new ScoreJudgement(D(answers["materiality"]!["score"]), D(answers["materiality"]!["confidence"]), new Dictionary<string, double>()),
            };
            var signals = node["deterministic"]!.AsArray()
                .Select(d => new DeterministicResult(S(d!, "check"), S(d!, "pattern"), d!["passed"]!.GetValue<bool>(), Array.Empty<string>(), null))
                .ToList();

            var route = Router.RouteSemantic(policy, requirement, decision, signals);

            Assert.Equal(S(node, "route"), route.Route.ToString());
            Assert.Equal(S(node, "routingRule"), route.Rule);
            Assert.Equal(S(node, "reason"), route.Reason);
            recomputed++;
        }

        Assert.Equal(3, recomputed);
    }

    [Fact]
    public async Task No_secret_reaches_the_record_or_the_logs_on_success_or_failure()
    {
        using var harness = await Harness.CreateAsync();

        var ok = await harness.ProposeAsync();
        var denied = await harness.ProposeAsync(harness.Jev(new ScriptedJevHandler(_ => ScriptedJevHandler.Json(HttpStatusCode.Forbidden,
            "{\"detail\":{\"error_type\":\"authentication_error\",\"message\":\"Must supply an API key!\"}}"))));

        foreach (var proposal in new[] { ok, denied })
        {
            var json = AuditRecord.ToJson(AuditRecord.Build(proposal, Info));
            Assert.DoesNotContain(Harness.FakeKey, json);
            Assert.DoesNotContain("Bearer", json);
        }

        Assert.NotEmpty(harness.Log.Lines);
        Assert.DoesNotContain(harness.Log.Lines, l => l.Contains(Harness.FakeKey, StringComparison.Ordinal));
        Assert.Contains(harness.Log.Lines, l => l.Contains("Authentication", StringComparison.Ordinal));
    }

    private static JsonNode Requirement(JsonObject audit, string id) =>
        audit["requirements"]!.AsArray().Single(r => S(r!, "requirementId") == id)!;

    private static string S(JsonNode node, string path)
    {
        var current = node;
        foreach (var part in path.Split('.'))
        {
            current = current![part];
        }

        return current!.GetValue<string>();
    }

    private static double D(JsonNode? node) => node!.GetValue<double>();

    private static ChoiceJudgement Choice(JsonNode node) =>
        new(node["label"]!.GetValue<string>(), node["confidence"]!.GetValue<double>(),
            node["probabilities"]!.AsObject().ToDictionary(p => p.Key, p => p.Value!.GetValue<double>()));
}
