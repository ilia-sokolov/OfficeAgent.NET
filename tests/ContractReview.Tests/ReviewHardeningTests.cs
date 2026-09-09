using OfficeAgent.Abstractions;
using OfficeAgent.Samples.ContractReview;
using Xunit;

namespace ContractReview.Tests;

/// <summary>
/// Regressions for the hardening pass: overlapping spans in one paragraph, a
/// partially judged run, and playbook patterns that cannot compile.
/// </summary>
public sealed class ReviewHardeningTests
{
    /// <summary>
    /// Two rules whose patterns overlap inside one paragraph - "60 days" sits
    /// inside "within 60 days". The engine keys a conflict on exact anchor
    /// identity, so it does not stop the pair, and a redline emitted before the
    /// other span's comment would strike through text that comment still needs.
    /// Every comment must therefore be emitted before any tracked change.
    /// </summary>
    [Fact]
    public async Task Overlapping_spans_in_one_paragraph_still_commit()
    {
        using var harness = await ReviewHarness.CreateAsync();

        var playbook = Playbook.FromJson(
            """
            {
              "name": "Overlapping spans",
              "rules": [
                {
                  "id": "SHORT-01", "title": "Days", "severity": "high",
                  "patterns": ["60 days"], "regex": false,
                  "rationale": "Payment must be within 30 days.", "action": "redline"
                },
                {
                  "id": "LONG-01", "title": "Whole phrase", "severity": "medium",
                  "patterns": ["within 60 days"], "regex": false,
                  "rationale": "State the trigger for the payment period.", "action": "commentOnly"
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
                    ReplacementText = c.RuleId == "SHORT-01" ? "30 days" : null,
                }).ToList()),
            "overlap.reviewed.docx");

        Assert.Equal(ReviewStatus.Applied, result.Status);
        Assert.True(harness.OutputExists("overlap.reviewed.docx"));

        var comments = harness.CommentsXml("overlap.reviewed.docx");
        Assert.Contains("[SHORT-01]", comments, StringComparison.Ordinal);
        Assert.Contains("[LONG-01]", comments, StringComparison.Ordinal);
        Assert.Contains("<w:ins ", harness.DocumentXml("overlap.reviewed.docx"), StringComparison.Ordinal);
    }

    [Fact]
    public void Every_comment_is_emitted_before_any_redline()
    {
        var playbook = Playbook.FromJson(
            """
            {
              "name": "Phases",
              "rules": [
                {
                  "id": "R1", "title": "One", "severity": "high",
                  "patterns": ["a"], "regex": false, "rationale": "r", "action": "redline"
                },
                {
                  "id": "R2", "title": "Two", "severity": "high",
                  "patterns": ["b"], "regex": false, "rationale": "r", "action": "redline"
                }
              ]
            }
            """);

        var candidates = new[]
        {
            new Candidate { Id = "c001", RuleId = "R1", ParaId = "p1", MatchedText = "alpha", Occurrence = 0 },
            new Candidate { Id = "c002", RuleId = "R2", ParaId = "p2", MatchedText = "beta", Occurrence = 0 },
        };

        var build = new PlanBuilder(playbook).Build(
            candidates,
            candidates.Select(c => new Finding
            {
                CandidateId = c.Id,
                Violation = true,
                Rationale = "breach",
                ReplacementText = "replacement",
            }).ToList());

        var kinds = build.Plan.Operations.Select(o => o is CommentOp ? "c" : "r").ToArray();

        Assert.Equal(new[] { "c", "c", "r", "r" }, kinds);
    }

    [Fact]
    public async Task A_partially_judged_run_writes_nothing()
    {
        using var harness = await ReviewHarness.CreateAsync();

        var playbook = Playbook.FromJson(
            """
            {
              "name": "Two rules",
              "rules": [
                {
                  "id": "PAY-01", "title": "Days", "severity": "high",
                  "patterns": ["60 days"], "regex": false,
                  "rationale": "Too long.", "action": "redline"
                },
                {
                  "id": "REN-01", "title": "Renewal", "severity": "medium",
                  "patterns": ["automatically renew"], "regex": false,
                  "rationale": "Needs notice.", "action": "commentOnly"
                }
              ]
            }
            """);

        var result = await new ReviewRunner(harness.Client).RunAsync(
            ReviewHarness.ConnectionId,
            harness.DocumentId,
            playbook,
            // Judges the first candidate and silently drops the second.
            (screening, _) => Task.FromResult<IReadOnlyList<Finding>>(
                screening.Candidates.Take(1).Select(c => new Finding
                {
                    CandidateId = c.Id,
                    Violation = true,
                    Rationale = "Sixty days is too long.",
                    ReplacementText = "30 days",
                }).ToList()),
            "partial.reviewed.docx");

        Assert.Equal(ReviewStatus.Incomplete, result.Status);
        Assert.Single(result.Unjudged);

        // The point of failing closed: a half-reviewed contract is never written.
        Assert.False(harness.OutputExists("partial.reviewed.docx"));
    }

    [Fact]
    public void An_invalid_regex_fails_at_load_naming_the_rule()
    {
        var json =
            """
            {
              "name": "Broken",
              "rules": [
                {
                  "id": "BAD-01", "title": "Unclosed group", "severity": "high",
                  "patterns": ["(unclosed"], "regex": true,
                  "rationale": "r", "action": "commentOnly"
                }
              ]
            }
            """;

        var ex = Assert.Throws<InvalidDataException>(() => Playbook.FromJson(json));

        Assert.Contains("BAD-01", ex.Message, StringComparison.Ordinal);
        Assert.Contains("(unclosed", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void A_literal_pattern_that_looks_like_a_bad_regex_is_accepted()
    {
        // regex: false means the pattern is literal text; it must not be compiled.
        var playbook = Playbook.FromJson(
            """
            {
              "name": "Literal",
              "rules": [
                {
                  "id": "LIT-01", "title": "Parenthesis", "severity": "low",
                  "patterns": ["(a) the Supplier"], "regex": false,
                  "rationale": "r", "action": "commentOnly"
                }
              ]
            }
            """);

        Assert.Single(playbook.Rules);
    }
}
