using OfficeAgent.Abstractions;
using OfficeAgent.Samples.ContractReview;
using Xunit;

namespace ContractReview.Tests;

/// <summary>
/// End-to-end tests for the review pipeline. The model is replaced by a scripted
/// judge, so the whole path runs offline and deterministically.
/// </summary>
public sealed class PipelineTests
{
    private static Func<ScreeningResult, CancellationToken, Task<IReadOnlyList<Finding>>> Judge(
        Func<Candidate, Finding?> decide) =>
        (screening, _) => Task.FromResult<IReadOnlyList<Finding>>(
            screening.Candidates.Select(decide).Where(f => f is not null).Select(f => f!).ToList());

    [Fact]
    public async Task Redline_lands_as_a_tracked_change_with_a_comment()
    {
        using var harness = await ReviewHarness.CreateAsync();
        var playbook = TestPlaybooks.PaymentRedline();

        var result = await new ReviewRunner(harness.Client).RunAsync(
            ReviewHarness.ConnectionId,
            harness.DocumentId,
            playbook,
            Judge(c => new Finding
            {
                CandidateId = c.Id,
                Violation = true,
                Rationale = "Sixty days exceeds the thirty-day house position.",
                ReplacementText = "30 days",
            }),
            "contract.reviewed.docx");

        Assert.Equal(ReviewStatus.Applied, result.Status);
        Assert.True(harness.OutputExists("contract.reviewed.docx"));

        var xml = harness.DocumentXml("contract.reviewed.docx");
        Assert.Contains("<w:ins ", xml, StringComparison.Ordinal);
        Assert.Contains("<w:del ", xml, StringComparison.Ordinal);
        Assert.Contains("30 days", xml, StringComparison.Ordinal);

        var comments = harness.CommentsXml("contract.reviewed.docx");
        Assert.Contains(PlanBuilder.CommentAuthor, comments, StringComparison.Ordinal);
        Assert.Contains("[PAY-01]", comments, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Source_contract_is_never_modified()
    {
        using var harness = await ReviewHarness.CreateAsync();
        var before = await File.ReadAllBytesAsync(harness.ContractPath);

        await new ReviewRunner(harness.Client).RunAsync(
            ReviewHarness.ConnectionId,
            harness.DocumentId,
            TestPlaybooks.PaymentRedline(),
            Judge(c => new Finding
            {
                CandidateId = c.Id,
                Violation = true,
                Rationale = "Too long.",
                ReplacementText = "30 days",
            }),
            "contract.reviewed.docx");

        var after = await File.ReadAllBytesAsync(harness.ContractPath);
        Assert.Equal(before, after);
    }

    [Fact]
    public async Task Comment_only_rule_produces_no_tracked_change()
    {
        using var harness = await ReviewHarness.CreateAsync();

        var result = await new ReviewRunner(harness.Client).RunAsync(
            ReviewHarness.ConnectionId,
            harness.DocumentId,
            TestPlaybooks.CommentOnly(),
            Judge(c => new Finding
            {
                CandidateId = c.Id,
                Violation = true,
                Rationale = "No notice period is stated.",
                ReplacementText = "this should be ignored",
            }),
            "renewal.reviewed.docx");

        Assert.Equal(ReviewStatus.Applied, result.Status);

        var xml = harness.DocumentXml("renewal.reviewed.docx");
        Assert.DoesNotContain("<w:ins ", xml, StringComparison.Ordinal);
        Assert.Contains("[REN-01]", harness.CommentsXml("renewal.reviewed.docx"), StringComparison.Ordinal);
    }

    [Fact]
    public async Task Two_rules_on_one_span_merge_into_one_comment_and_one_redline()
    {
        using var harness = await ReviewHarness.CreateAsync();

        var result = await new ReviewRunner(harness.Client).RunAsync(
            ReviewHarness.ConnectionId,
            harness.DocumentId,
            TestPlaybooks.TwoRulesOneSpan(),
            Judge(c => new Finding
            {
                CandidateId = c.Id,
                Violation = true,
                Rationale = $"Breach reported by {c.RuleId}.",
                ReplacementText = "30 days",
            }),
            "merged.reviewed.docx");

        Assert.Equal(ReviewStatus.Applied, result.Status);

        var comments = harness.CommentsXml("merged.reviewed.docx");
        Assert.Contains("[HIGH-01]", comments, StringComparison.Ordinal);
        Assert.Contains("[LOW-01]", comments, StringComparison.Ordinal);

        // The higher severity carries the redline; the other is recorded as superseded.
        var high = result.Items.Single(i => i.RuleId == "HIGH-01");
        var low = result.Items.Single(i => i.RuleId == "LOW-01");
        Assert.Equal(RedlineOutcome.Applied, high.Redline);
        Assert.Equal(RedlineOutcome.SupersededOnSpan, low.Redline);

        // Exactly one tracked insertion for the span.
        var xml = harness.DocumentXml("merged.reviewed.docx");
        var insertions = xml.Split("<w:ins ", StringSplitOptions.None).Length - 1;
        Assert.Equal(1, insertions);
    }

    [Fact]
    public async Task Redline_without_wording_is_reported_not_silently_dropped()
    {
        using var harness = await ReviewHarness.CreateAsync();

        var result = await new ReviewRunner(harness.Client).RunAsync(
            ReviewHarness.ConnectionId,
            harness.DocumentId,
            TestPlaybooks.PaymentRedline(),
            Judge(c => new Finding
            {
                CandidateId = c.Id,
                Violation = true,
                Rationale = "Too long, but I forgot to draft wording.",
                ReplacementText = null,
            }),
            "nowording.reviewed.docx");

        Assert.Equal(ReviewStatus.Applied, result.Status);
        Assert.Equal(RedlineOutcome.NoWordingProposed, result.Items.Single().Redline);

        var xml = harness.DocumentXml("nowording.reviewed.docx");
        Assert.DoesNotContain("<w:ins ", xml, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Compliant_judgement_writes_nothing_but_is_reported()
    {
        using var harness = await ReviewHarness.CreateAsync();

        var result = await new ReviewRunner(harness.Client).RunAsync(
            ReviewHarness.ConnectionId,
            harness.DocumentId,
            TestPlaybooks.PaymentRedline(),
            Judge(c => new Finding
            {
                CandidateId = c.Id,
                Violation = false,
                Rationale = "The sixty days applies to a disputed invoice only.",
            }),
            "clean.reviewed.docx");

        Assert.Equal(ReviewStatus.NoBreaches, result.Status);
        Assert.False(harness.OutputExists("clean.reviewed.docx"));
        Assert.Single(result.Items);
        Assert.False(result.Items[0].Violation);
    }

    [Fact]
    public async Task Finding_naming_an_unknown_candidate_aborts_before_any_write()
    {
        using var harness = await ReviewHarness.CreateAsync();

        var runner = new ReviewRunner(harness.Client);

        await Assert.ThrowsAsync<InvalidOperationException>(() => runner.RunAsync(
            ReviewHarness.ConnectionId,
            harness.DocumentId,
            TestPlaybooks.PaymentRedline(),
            Judge(_ => new Finding
            {
                CandidateId = "c999",
                Violation = true,
                Rationale = "Hallucinated candidate.",
                ReplacementText = "30 days",
            }),
            "bogus.reviewed.docx"));

        Assert.False(harness.OutputExists("bogus.reviewed.docx"));
    }

    [Fact]
    public async Task Rules_that_match_nothing_are_reported_separately()
    {
        using var harness = await ReviewHarness.CreateAsync();

        var result = await new ReviewRunner(harness.Client).RunAsync(
            ReviewHarness.ConnectionId,
            harness.DocumentId,
            TestPlaybooks.Undetectable(),
            Judge(_ => null),
            "none.reviewed.docx");

        Assert.Equal(ReviewStatus.NoCandidates, result.Status);
        Assert.Single(result.Screening.Undetected);
        Assert.Equal("NONE-01", result.Screening.Undetected[0].RuleId);

        var report = ReviewRunner.Report(TestPlaybooks.Undetectable(), result);
        Assert.Contains("Not detected", report, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Screener_sources_every_anchor_from_the_engine()
    {
        using var harness = await ReviewHarness.CreateAsync();

        var screening = await new Screener(harness.Client).ScreenAsync(
            ReviewHarness.ConnectionId,
            harness.DocumentId,
            TestPlaybooks.PaymentRedline());

        var candidate = Assert.Single(screening.Candidates);
        Assert.Equal("PAY-01", candidate.RuleId);
        Assert.Equal("60 days", candidate.MatchedText);
        Assert.False(string.IsNullOrWhiteSpace(candidate.ParaId));
        Assert.StartsWith("c", candidate.Id, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Regex_rule_matches_and_carries_the_matched_text()
    {
        using var harness = await ReviewHarness.CreateAsync();

        var playbook = Playbook.FromJson(
            """
            {
              "name": "Regex",
              "rules": [
                {
                  "id": "PAY-RX",
                  "title": "Payment days",
                  "severity": "high",
                  "patterns": ["\\b\\d{2}\\s+days\\b"],
                  "regex": true,
                  "rationale": "Payment must be within 30 days.",
                  "action": "redline"
                }
              ]
            }
            """);

        var screening = await new Screener(harness.Client).ScreenAsync(
            ReviewHarness.ConnectionId,
            harness.DocumentId,
            playbook);

        var candidate = Assert.Single(screening.Candidates);
        Assert.Equal("60 days", candidate.MatchedText);
    }

    [Fact]
    public void Plan_puts_the_comment_before_the_tracked_change_on_the_same_span()
    {
        var playbook = TestPlaybooks.PaymentRedline();
        var candidates = new[]
        {
            new Candidate
            {
                Id = "c001",
                RuleId = "PAY-01",
                ParaId = "10000003",
                MatchedText = "60 days",
                Occurrence = 0,
            },
        };

        var build = new PlanBuilder(playbook).Build(
            candidates,
            new[]
            {
                new Finding
                {
                    CandidateId = "c001",
                    Violation = true,
                    Rationale = "Too long.",
                    ReplacementText = "30 days",
                },
            });

        Assert.Collection(
            build.Plan.Operations,
            op => Assert.IsType<CommentOp>(op),
            op => Assert.IsType<ChangeTextOp>(op));
    }
}
