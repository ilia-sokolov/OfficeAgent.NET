using OfficeAgent.Abstractions;

namespace OfficeAgent.Samples.ContractReview;

/// <summary>
/// Turns judgements into exactly one <see cref="DocumentPlan"/>. This is the only
/// code that writes to the document, and no model output reaches it except the
/// words inside a tracked insertion and the text of a comment.
/// </summary>
public sealed class PlanBuilder
{
    /// <summary>The author shown on every comment this reviewer writes.</summary>
    public const string CommentAuthor = "Contract Review";

    /// <summary>The initials shown on every comment this reviewer writes.</summary>
    public const string CommentInitials = "CR";

    private readonly Playbook _playbook;

    /// <summary>Initializes a new instance of the <see cref="PlanBuilder"/> class.</summary>
    /// <param name="playbook">The playbook the findings were judged against.</param>
    public PlanBuilder(Playbook playbook) => _playbook = playbook;

    /// <summary>Builds the review plan.</summary>
    /// <param name="candidates">Candidates produced by screening.</param>
    /// <param name="findings">Judgements keyed by candidate id.</param>
    /// <returns>The plan and the report rows describing it.</returns>
    public BuildResult Build(
        IReadOnlyList<Candidate> candidates,
        IReadOnlyList<Finding> findings,
        SnapshotToken? snapshot = null)
    {
        var byId = candidates.ToDictionary(c => c.Id, StringComparer.Ordinal);
        var rules = _playbook.Rules.ToDictionary(r => r.Id, StringComparer.OrdinalIgnoreCase);

        // Only breaches reach the document. Everything judged compliant still
        // reaches the report, because a reviewer needs to know what was checked.
        var breaches = new List<(Candidate Candidate, Finding Finding, PlaybookRule Rule)>();
        var items = new List<ReviewItem>();
        var judged = new HashSet<string>(StringComparer.Ordinal);

        foreach (var finding in findings)
        {
            if (!judged.Add(finding.CandidateId))
            {
                throw new InvalidOperationException(
                    $"Candidate '{finding.CandidateId}' was judged more than once.");
            }

            if (!byId.TryGetValue(finding.CandidateId, out var candidate))
            {
                throw new InvalidOperationException(
                    $"Finding references unknown candidate '{finding.CandidateId}'.");
            }

            if (!rules.TryGetValue(candidate.RuleId, out var rule))
            {
                throw new InvalidOperationException(
                    $"Candidate '{candidate.Id}' references unknown rule '{candidate.RuleId}'.");
            }

            if (finding.Violation)
            {
                breaches.Add((candidate, finding, rule));
            }
            else
            {
                items.Add(new ReviewItem
                {
                    RuleId = rule.Id,
                    Title = rule.Title,
                    Severity = rule.Severity,
                    MatchedText = candidate.MatchedText,
                    Violation = false,
                    Rationale = finding.Rationale,
                    Redline = RedlineOutcome.NotRequested,
                });
            }
        }

        // Two phases, not one pass. Every comment is emitted before any tracked
        // change, because a redline on one span can strike through text that a
        // comment on an *overlapping* span still needs to resolve against. Within
        // a span the same rule applies, so phase order covers both cases.
        var comments = new List<PlanOperation>();
        var redlines = new List<PlanOperation>();

        // Two rules can breach on the same span. The engine keys a conflict on
        // verb plus anchor, so two comments there would be refused; they are
        // merged into one comment instead, and only the most severe breach may
        // carry the tracked change.
        var spans = breaches
            .GroupBy(b => (b.Candidate.ParaId, b.Candidate.Occurrence, b.Candidate.MatchedText))
            .ToList();

        // At most one tracked change per paragraph. Two playbook patterns can match
        // overlapping but textually different wording in one clause; the engine keys a
        // conflict on exact anchor identity, so it would not stop them, and the second
        // replacement would fail on text the first one had already struck through.
        // Which span carries the paragraph's single redline is decided here,
        // from the rules and the document, not from the order the model happened
        // to report its findings in.
        var redlineWinner = breaches
            .Where(b => b.Rule.Action == FindingAction.Redline
                        && !string.IsNullOrWhiteSpace(b.Finding.ReplacementText))
            .GroupBy(b => b.Candidate.ParaId, StringComparer.Ordinal)
            .ToDictionary(
                g => g.Key,
                g => g
                    .OrderByDescending(b => b.Rule.Severity)
                    .ThenBy(b => b.Rule.Id, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(b => b.Candidate.Occurrence)
                    .ThenBy(b => b.Candidate.MatchedText, StringComparer.Ordinal)
                    .First()
                    .Candidate
                    .Id,
                StringComparer.Ordinal);

        foreach (var span in spans)
        {
            var ordered = span
                .OrderByDescending(b => b.Rule.Severity)
                .ThenBy(b => b.Rule.Id, StringComparer.OrdinalIgnoreCase)
                .ToList();

            var anchor = new TextSpanAnchor
            {
                ParaId = span.Key.ParaId,
                Expect = span.Key.MatchedText,
                Occurrence = span.Key.Occurrence,
            };

            // Merged into one comment per span: the engine keys a conflict on verb
            // plus anchor, so two comments on one span would be refused.
            var body = string.Join(
                " ",
                ordered.Select(b => $"[{b.Rule.Id}] {b.Finding.Rationale}".Trim()));

            comments.Add(new CommentOp
            {
                Target = anchor,
                Text = body,
                Author = CommentAuthor,
                Initials = CommentInitials,
                Action = CommentAction.Add,
            });

            foreach (var (candidate, finding, rule) in ordered)
            {
                var outcome = RedlineOutcome.NotRequested;

                if (rule.Action == FindingAction.Redline)
                {
                    if (string.IsNullOrWhiteSpace(finding.ReplacementText))
                    {
                        // The rule allowed a redline and the model proposed no
                        // wording. Say so rather than silently commenting.
                        outcome = RedlineOutcome.NoWordingProposed;
                    }
                    else if (!redlineWinner.TryGetValue(span.Key.ParaId, out var winner)
                             || winner != candidate.Id)
                    {
                        outcome = RedlineOutcome.SupersededOnSpan;
                    }
                    else
                    {
                        redlines.Add(new ChangeTextOp
                        {
                            Target = anchor,
                            With = finding.ReplacementText!,
                            Mode = ChangeMode.Tracked,
                        });


                        outcome = RedlineOutcome.Applied;
                    }
                }

                items.Add(new ReviewItem
                {
                    RuleId = rule.Id,
                    Title = rule.Title,
                    Severity = rule.Severity,
                    MatchedText = candidate.MatchedText,
                    Violation = true,
                    Rationale = finding.Rationale,
                    Redline = outcome,
                });
            }
        }

        // Bind the plan to the document as it was when screening ran. Anchor
        // checks only catch a change to the text being edited; the snapshot
        // catches changes in other snapshotted text hosts, which is just as good
        // a reason to refuse a stale review.
        var plan = new DocumentPlan
        {
            Snapshot = snapshot,
            Operations = comments.Concat(redlines).ToList(),
        };
        return new BuildResult(plan, items);
    }
}

/// <summary>The plan and the rows that describe it.</summary>
/// <param name="Plan">The single plan to preview and apply.</param>
/// <param name="Items">Report rows, including candidates judged compliant.</param>
public sealed record BuildResult(DocumentPlan Plan, IReadOnlyList<ReviewItem> Items);
