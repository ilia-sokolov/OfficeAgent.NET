using OfficeAgent.Abstractions;
using OfficeAgent.Samples.ContractReview;
using Xunit;

namespace ContractReview.Tests;

/// <summary>
/// Concurrency and determinism: a review must not land on a contract that moved
/// under it, and which clause gets the redline must not depend on the order the
/// model happened to report its findings in.
/// </summary>
public sealed class SnapshotAndSelectionTests
{
    [Fact]
    public async Task A_change_to_an_unrelated_clause_invalidates_the_review()
    {
        // The anchors this review targets still resolve perfectly - only a
        // different paragraph moved. Anchor checking cannot see that; the
        // snapshot can, and the review must be refused rather than applied to a
        // contract the reviewer never saw.
        using var harness = await ReviewHarness.CreateAsync();

        var result = await new ReviewRunner(harness.Client).RunAsync(
            ReviewHarness.ConnectionId,
            harness.DocumentId,
            TestPlaybooks.PaymentRedline(),
            async (screening, ct) =>
            {
                var edited = ContractFixture.Create(governingLaw: "the laws of the Netherlands");
                await File.WriteAllBytesAsync(harness.ContractPath, edited, ct);

                return screening.Candidates.Select(c => new Finding
                {
                    CandidateId = c.Id,
                    Violation = true,
                    Rationale = "Sixty days exceeds the house position.",
                    ReplacementText = "30 days",
                }).ToList();
            },
            "unrelated.reviewed.docx");

        Assert.Equal(ReviewStatus.Rejected, result.Status);
        Assert.Contains(result.Errors, e => e.Contains("stale-snapshot", StringComparison.Ordinal));
        Assert.False(harness.OutputExists("unrelated.reviewed.docx"));
    }

    [Fact]
    public async Task An_untouched_document_still_applies()
    {
        // The snapshot must not make ordinary runs fail.
        using var harness = await ReviewHarness.CreateAsync();

        var result = await new ReviewRunner(harness.Client).RunAsync(
            ReviewHarness.ConnectionId,
            harness.DocumentId,
            TestPlaybooks.PaymentRedline(),
            (screening, _) => Task.FromResult<IReadOnlyList<Finding>>(
                screening.Candidates.Select(c => new Finding
                {
                    CandidateId = c.Id,
                    Violation = true,
                    Rationale = "Too long.",
                    ReplacementText = "30 days",
                }).ToList()),
            "clean.reviewed.docx");

        Assert.Equal(ReviewStatus.Applied, result.Status);
        Assert.True(harness.OutputExists("clean.reviewed.docx"));
    }

    [Fact]
    public async Task A_change_before_a_clean_judgement_invalidates_the_review()
    {
        using var harness = await ReviewHarness.CreateAsync();

        var result = await new ReviewRunner(harness.Client).RunAsync(
            ReviewHarness.ConnectionId,
            harness.DocumentId,
            TestPlaybooks.PaymentRedline(),
            async (screening, ct) =>
            {
                var edited = ContractFixture.Create(governingLaw: "the laws of the Netherlands");
                await File.WriteAllBytesAsync(harness.ContractPath, edited, ct);

                return screening.Candidates.Select(c => new Finding
                {
                    CandidateId = c.Id,
                    Violation = false,
                    Rationale = "Accepted for this agreement.",
                }).ToList();
            },
            "stale-clean.reviewed.docx");

        Assert.Equal(ReviewStatus.Rejected, result.Status);
        Assert.Contains(result.Errors, e => e.Contains("stale-snapshot", StringComparison.Ordinal));
        Assert.False(harness.OutputExists("stale-clean.reviewed.docx"));
    }

    private static Playbook TwoSeverities() => Playbook.FromJson(
        """
        {
          "name": "Severity race",
          "rules": [
            {
              "id": "MED-01", "title": "Medium", "severity": "medium",
              "patterns": ["undisputed"], "regex": false,
              "rationale": "Define undisputed.", "action": "redline"
            },
            {
              "id": "HIGH-01", "title": "High", "severity": "high",
              "patterns": ["60 days"], "regex": false,
              "rationale": "Payment must be within 30 days.", "action": "redline"
            }
          ]
        }
        """);

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void The_most_severe_finding_wins_whatever_order_the_model_reports(bool reverseOrder)
    {
        // Both rules breach the same paragraph. Only one redline can land there,
        // and it must be the high-severity one either way.
        var playbook = TwoSeverities();

        var candidates = new[]
        {
            new Candidate { Id = "c001", RuleId = "MED-01", ParaId = "p1", MatchedText = "undisputed", Occurrence = 0 },
            new Candidate { Id = "c002", RuleId = "HIGH-01", ParaId = "p1", MatchedText = "60 days", Occurrence = 0 },
        };

        var findings = candidates.Select(c => new Finding
        {
            CandidateId = c.Id,
            Violation = true,
            Rationale = $"Breach per {c.RuleId}.",
            ReplacementText = c.RuleId == "HIGH-01" ? "30 days" : "properly disputed",
        }).ToList();

        if (reverseOrder)
        {
            findings.Reverse();
        }

        var build = new PlanBuilder(playbook).Build(candidates, findings);

        var high = build.Items.Single(i => i.RuleId == "HIGH-01");
        var medium = build.Items.Single(i => i.RuleId == "MED-01");

        Assert.Equal(RedlineOutcome.Applied, high.Redline);
        Assert.Equal(RedlineOutcome.SupersededOnSpan, medium.Redline);

        var redline = Assert.Single(build.Plan.Operations.OfType<ChangeTextOp>());
        Assert.Equal("30 days", redline.With);
    }

    [Fact]
    public void The_plan_carries_the_snapshot_it_was_built_against()
    {
        var playbook = TestPlaybooks.PaymentRedline();
        var candidates = new[]
        {
            new Candidate { Id = "c001", RuleId = "PAY-01", ParaId = "p1", MatchedText = "60 days", Occurrence = 0 },
        };

        var build = new PlanBuilder(playbook).Build(
            candidates,
            new[]
            {
                new Finding { CandidateId = "c001", Violation = true, Rationale = "r", ReplacementText = "30 days" },
            },
            new SnapshotToken("etag-123"));

        Assert.NotNull(build.Plan.Snapshot);
        Assert.Equal("etag-123", build.Plan.Snapshot!.ETag);
    }

}
