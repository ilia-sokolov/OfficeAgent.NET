using OfficeAgent.Abstractions;

namespace OfficeAgent.Samples.RequirementsReview;

/// <summary>What the plan does for one requirement.</summary>
/// <param name="RequirementId">The requirement.</param>
/// <param name="Operations">Human-readable operation summaries.</param>
/// <param name="Note">Why the host did less than the route allowed, if it did.</param>
public sealed record PlannedChange(string RequirementId, IReadOnlyList<string> Operations, string? Note);

/// <summary>The plan and its per-requirement description.</summary>
/// <param name="Plan">The one plan to preview, approve, and commit.</param>
/// <param name="Changes">What each requirement contributed.</param>
public sealed record BuiltPlan(DocumentPlan Plan, IReadOnlyList<PlannedChange> Changes);

/// <summary>
/// Turns routed requirements into one <see cref="DocumentPlan"/>. Anchors come from host
/// evidence; the only model text that reaches the document is a validated comment and a
/// validated replacement inside a tracked change.
/// </summary>
public static class PlanBuilder
{
    /// <summary>Author on every comment and tracked change this workflow writes.</summary>
    public const string Author = "Requirements Review";

    /// <summary>Initials on every comment this workflow writes.</summary>
    public const string Initials = "RR";

    /// <summary>Builds the plan.</summary>
    /// <param name="evidence">The document evidence, including the snapshot.</param>
    /// <param name="items">Routed requirements.</param>
    /// <returns>The plan.</returns>
    public static BuiltPlan Build(DocumentEvidence evidence, IReadOnlyList<ReviewItem> items)
    {
        var comments = new List<PlanOperation>();
        var redlines = new List<PlanOperation>();
        var changes = new List<PlannedChange>();

        // One tracked change per paragraph. When two proposals land in one paragraph the
        // more material one wins, then the lower requirement id, so the outcome does not
        // depend on evaluation order.
        var redlineWinner = items
            .Where(i => i.Routing.Action == WordAction.CommentAndTrackedChange && i.Draft?.Proposal?.ReplacementText is not null)
            .GroupBy(i => i.Draft!.Proposal!.ParaId, StringComparer.Ordinal)
            .ToDictionary(
                g => g.Key,
                g => g.OrderByDescending(i => i.Decision?.Materiality?.Score ?? 0)
                      .ThenBy(i => i.Requirement.Id, StringComparer.Ordinal)
                      .First().Requirement.Id,
                StringComparer.Ordinal);

        foreach (var item in items)
        {
            var ops = new List<string>();
            string? note = null;
            var tag = $"[{item.Requirement.Id} v{item.Requirement.Version}]";

            switch (item.Routing.Action)
            {
                case WordAction.None:
                    break;

                case WordAction.ReviewComment:
                    if (item.Evidence.HeadingParaId is { } headingId && item.Evidence.HeadingText is { } heading)
                    {
                        comments.Add(Comment(headingId, heading, $"{tag} Needs human review. {item.Routing.Note} No change proposed."));
                        ops.Add($"comment on heading \"{heading}\"");
                    }
                    else
                    {
                        note = "no unique section heading to anchor a comment; listed in the report only";
                    }

                    break;

                case WordAction.Comment:
                case WordAction.CommentAndTrackedChange:
                    var proposal = item.Draft?.Proposal;
                    if (proposal is null)
                    {
                        // The violation stands without wording. Say so on the section.
                        var first = item.Evidence.Items.FirstOrDefault();
                        if (first is not null)
                        {
                            comments.Add(Comment(first.ParaId, first.Text,
                                $"{tag} {item.Routing.Note} No valid wording was drafted; please review this section."));
                            ops.Add($"comment on {first.EvidenceId}");
                        }

                        note = "no valid draft; host comment only";
                        break;
                    }

                    comments.Add(Comment(proposal.ParaId, proposal.TargetText, $"{tag} {proposal.CommentText}"));
                    ops.Add($"comment on \"{proposal.TargetText}\" in {proposal.EvidenceId}");

                    if (item.Routing.Action == WordAction.CommentAndTrackedChange && proposal.ReplacementText is { } replacement)
                    {
                        if (redlineWinner.TryGetValue(proposal.ParaId, out var winner) && winner == item.Requirement.Id)
                        {
                            redlines.Add(new ChangeTextOp
                            {
                                Target = Anchor(proposal.ParaId, proposal.TargetText),
                                With = replacement,
                                Mode = ChangeMode.Tracked,
                            });
                            ops.Add($"tracked change \"{proposal.TargetText}\" -> \"{replacement}\"");
                        }
                        else
                        {
                            note = $"tracked change superseded by {winner} in the same paragraph; comment only";
                        }
                    }

                    break;
            }

            changes.Add(new PlannedChange(item.Requirement.Id, ops, note));
        }

        // Every comment before any tracked change: a replacement strikes text through, and
        // a comment anchored on text in the same span must resolve before that happens.
        var plan = new DocumentPlan
        {
            Snapshot = evidence.Snapshot,
            Revision = new RevisionMetadata { Author = Author },
            Operations = comments.Concat(redlines).ToList(),
        };

        return new BuiltPlan(plan, changes);
    }

    private static CommentOp Comment(string paraId, string expect, string text) => new()
    {
        Target = Anchor(paraId, expect),
        Text = text,
        Author = Author,
        Initials = Initials,
        Action = CommentAction.Add,
    };

    private static TextSpanAnchor Anchor(string paraId, string expect) => new()
    {
        ParaId = paraId,
        Expect = expect,
        Occurrence = 0,
    };
}
