using Microsoft.Extensions.AI;
using OfficeAgent.Samples.ContractReview;
using Xunit;

namespace ContractReview.Tests;

/// <summary>
/// The paths that matter when something goes wrong: a plan the engine refuses, a
/// model that judges the same candidate twice, and a model that judges nothing.
/// </summary>
public sealed class FailurePathTests
{
    [Fact]
    public async Task A_plan_the_engine_refuses_writes_nothing_and_surfaces_the_error()
    {
        using var harness = await ReviewHarness.CreateAsync();

        // A finding whose replacement targets an anchor the screener never
        // produced: the candidate is real, but the run doctors the expect text so
        // the engine cannot resolve it.
        var screener = new Screener(harness.Client);
        var playbook = TestPlaybooks.PaymentRedline();
        var screening = await screener.ScreenAsync(ReviewHarness.ConnectionId, harness.DocumentId, playbook);

        var doctored = new ScreeningResult(
            screening.Candidates
                .Select(c => c with { MatchedText = "ninety days" })
                .ToList(),
            screening.Undetected);

        var build = new PlanBuilder(playbook).Build(
            doctored.Candidates,
            doctored.Candidates.Select(c => new Finding
            {
                CandidateId = c.Id,
                Violation = true,
                Rationale = "Too long.",
                ReplacementText = "30 days",
            }).ToList());

        var preview = await harness.Client.PreviewAsync(
            ReviewHarness.ConnectionId, harness.DocumentId, build.Plan);

        Assert.False(preview.IsValid);
        Assert.Contains(preview.Errors, e => e.Code is "anchor-not-found" or "expect-mismatch");
    }

    [Fact]
    public async Task Document_changing_under_the_review_is_refused_and_writes_nothing()
    {
        // Drift between screening and applying is the realistic failure: someone
        // edits the contract while the model is judging. The anchors the screener
        // produced no longer resolve, preview refuses the plan, and the run must
        // write nothing rather than a partly redlined contract.
        using var harness = await ReviewHarness.CreateAsync();

        var result = await new ReviewRunner(harness.Client).RunAsync(
            ReviewHarness.ConnectionId,
            harness.DocumentId,
            TestPlaybooks.PaymentRedline(),
            async (screening, ct) =>
            {
                // The judging step is where the wall-clock time goes, so this is
                // exactly when a concurrent edit lands.
                var rewritten = ContractFixture.Create(paymentDays: 30);
                await File.WriteAllBytesAsync(harness.ContractPath, rewritten, ct);

                return screening.Candidates.Select(c => new Finding
                {
                    CandidateId = c.Id,
                    Violation = true,
                    Rationale = "Sixty days exceeds the house position.",
                    ReplacementText = "30 days",
                }).ToList();
            },
            "drifted.reviewed.docx");

        Assert.Equal(ReviewStatus.Rejected, result.Status);
        Assert.NotEmpty(result.Errors);
        Assert.False(harness.OutputExists("drifted.reviewed.docx"));
    }

    [Fact]
    public async Task Judging_one_candidate_twice_is_refused()
    {
        using var harness = await ReviewHarness.CreateAsync();

        await Assert.ThrowsAsync<InvalidOperationException>(() =>
            new ReviewRunner(harness.Client).RunAsync(
                ReviewHarness.ConnectionId,
                harness.DocumentId,
                TestPlaybooks.PaymentRedline(),
                (screening, _) => Task.FromResult<IReadOnlyList<Finding>>(
                    screening.Candidates
                        .SelectMany(c => new[]
                        {
                            new Finding { CandidateId = c.Id, Violation = true, Rationale = "One.", ReplacementText = "30 days" },
                            new Finding { CandidateId = c.Id, Violation = true, Rationale = "Twice.", ReplacementText = "45 days" },
                        })
                        .ToList()),
                "double.reviewed.docx"));

        Assert.False(harness.OutputExists("double.reviewed.docx"));
    }

    [Fact]
    public async Task An_unjudged_candidate_makes_the_run_incomplete_and_writes_nothing()
    {
        using var harness = await ReviewHarness.CreateAsync();

        var result = await new ReviewRunner(harness.Client).RunAsync(
            ReviewHarness.ConnectionId,
            harness.DocumentId,
            TestPlaybooks.PaymentRedline(),
            (_, _) => Task.FromResult<IReadOnlyList<Finding>>(Array.Empty<Finding>()),
            "silent.reviewed.docx");

        // Fail closed: a model that judged nothing must not read as "no breaches".
        Assert.Equal(ReviewStatus.Incomplete, result.Status);
        Assert.Single(result.Unjudged);
        Assert.False(harness.OutputExists("silent.reviewed.docx"));

        var report = ReviewRunner.Report(TestPlaybooks.PaymentRedline(), result);
        Assert.Contains("Not judged", report, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Only_one_tracked_change_lands_per_paragraph()
    {
        using var harness = await ReviewHarness.CreateAsync();

        // Two rules whose patterns match different text in the same paragraph.
        var playbook = Playbook.FromJson(
            """
            {
              "name": "Same paragraph",
              "rules": [
                {
                  "id": "A-01", "title": "Days", "severity": "high",
                  "patterns": ["60 days"], "regex": false,
                  "rationale": "Too long.", "action": "redline"
                },
                {
                  "id": "B-01", "title": "Undisputed", "severity": "medium",
                  "patterns": ["undisputed"], "regex": false,
                  "rationale": "Define undisputed.", "action": "redline"
                }
              ]
            }
            """);

        var result = await new ReviewRunner(harness.Client).RunAsync(
            ReviewHarness.ConnectionId,
            harness.DocumentId,
            playbook,
            (screening, _) => Task.FromResult<IReadOnlyList<Finding>>(
                screening.Candidates.Select(c => new Finding
                {
                    CandidateId = c.Id,
                    Violation = true,
                    Rationale = $"Breach per {c.RuleId}.",
                    ReplacementText = c.RuleId == "A-01" ? "30 days" : "properly disputed",
                }).ToList()),
            "oneperpara.reviewed.docx");

        Assert.Equal(ReviewStatus.Applied, result.Status);

        var xml = harness.DocumentXml("oneperpara.reviewed.docx");
        var insertions = xml.Split("<w:ins ", StringSplitOptions.None).Length - 1;
        Assert.Equal(1, insertions);

        Assert.Contains(result.Items, i => i.Redline == RedlineOutcome.Applied);
        Assert.Contains(result.Items, i => i.Redline == RedlineOutcome.SupersededOnSpan);
    }

    [Fact]
    public async Task Empty_replacement_wording_is_normalised_to_null()
    {
        var collector = new FindingCollector();
        var tool = collector.AsTool();

        await tool.InvokeAsync(new AIFunctionArguments(new Dictionary<string, object?>
        {
            ["candidateId"] = "c001",
            ["violation"] = true,
            ["rationale"] = "  Spaces are trimmed.  ",
            ["replacementText"] = "   ",
        }));

        var finding = Assert.Single(collector.Findings);
        Assert.Equal("c001", finding.CandidateId);
        Assert.Null(finding.ReplacementText);
        Assert.Equal("Spaces are trimmed.", finding.Rationale);
    }

    [Fact]
    public async Task Model_supplied_text_is_bounded()
    {
        var collector = new FindingCollector();
        var tool = collector.AsTool();

        await tool.InvokeAsync(new AIFunctionArguments(new Dictionary<string, object?>
        {
            ["candidateId"] = "c001",
            ["violation"] = true,
            ["rationale"] = new string('r', 5_000),
            ["replacementText"] = new string('w', 5_000),
        }));

        var finding = Assert.Single(collector.Findings);
        Assert.True(finding.Rationale.Length <= 500);
        Assert.True(finding.ReplacementText!.Length <= 300);
    }

    [Fact]
    public async Task Every_write_capable_office_tool_is_withheld_from_the_model()
    {
        using var harness = await ReviewHarness.CreateAsync();

        var all = new OfficeAgent.AgentFramework.OfficeAgentTools(harness.Client)
            .AsAIFunctions()
            .Select(f => f.Name)
            .ToHashSet(StringComparer.Ordinal);

        var exposed = new ReviewAgent(new NullChatClient(), harness.Client)
            .BuildTools(new FindingCollector())
            .OfType<AIFunction>()
            .Select(f => f.Name)
            .ToHashSet(StringComparer.Ordinal);

        // Anything the toolkit offers that the sample does not expose must be
        // withheld deliberately, and nothing withheld may reappear.
        var withheld = all.Except(exposed).ToList();
        Assert.Contains("apply_plan", withheld);
        Assert.Contains("preview_plan", withheld);

        foreach (var name in exposed.Where(n => n != "report_finding"))
        {
            Assert.Contains(name, ReviewAgent.ReadOnlyToolNames);
        }
    }

    private sealed class NullChatClient : IChatClient
    {
        public Task<ChatResponse> GetResponseAsync(
            IEnumerable<ChatMessage> messages,
            ChatOptions? options = null,
            CancellationToken cancellationToken = default) =>
            throw new NotSupportedException();

        public IAsyncEnumerable<ChatResponseUpdate> GetStreamingResponseAsync(
            IEnumerable<ChatMessage> messages,
            ChatOptions? options = null,
            CancellationToken cancellationToken = default) =>
            throw new NotSupportedException();

        public object? GetService(Type serviceType, object? serviceKey = null) => null;

        public void Dispose()
        {
        }
    }
}
