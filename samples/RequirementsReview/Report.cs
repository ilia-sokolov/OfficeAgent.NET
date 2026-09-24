using System.Globalization;
using System.Text;

namespace OfficeAgent.Samples.RequirementsReview;

/// <summary>Renders the reviewer's summary. The audit record is the complete version.</summary>
public static class Report
{
    /// <summary>Renders a proposal and, when it ran, the commit.</summary>
    /// <param name="proposal">The proposal.</param>
    /// <param name="commit">The commit outcome, or null.</param>
    /// <returns>Markdown.</returns>
    public static string Render(ReviewProposal proposal, CommitOutcome? commit = null)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"# Requirements review: {proposal.Evidence.Name}");
        sb.AppendLine();
        sb.AppendLine($"- Requirement set: {proposal.Requirements.Id} {proposal.Requirements.Version}");
        sb.AppendLine($"- Policy: {proposal.Policy.Version}");
        sb.AppendLine($"- Document SHA-256: {proposal.Evidence.SourceSha256[..16]}...");
        sb.AppendLine($"- Proposal: {proposal.Status} ({proposal.StatusReason})");
        if (commit is not null)
        {
            sb.AppendLine($"- Commit: {commit.Status} ({commit.Reason})");
            if (commit.Output is { } output)
            {
                sb.AppendLine($"- Reviewed copy: {output.Name}");
            }
        }

        sb.AppendLine();
        sb.AppendLine("| Requirement | Evidence | Answers | Route | Word |");
        sb.AppendLine("| --- | --- | --- | --- | --- |");
        foreach (var item in proposal.Items)
        {
            var evidence = item.Evidence.Items.Count == 0
                ? "none"
                : string.Join(", ", item.Evidence.Items.Select(i => $"{i.EvidenceId} ({i.ParaId})"));
            sb.AppendLine(
                $"| {item.Requirement.Id} v{item.Requirement.Version} | {evidence} | {Answers(item)} | " +
                $"{item.Routing.Route} / {item.Routing.Outcome}: {Escape(item.Routing.Reason)} | {Word(item)} |");
        }

        sb.AppendLine();
        sb.AppendLine("A pass means the supplied evidence was judged to meet the requirement text at the recorded confidence. " +
                      "It is not proof of compliance, and sections not named by a requirement were not read.");
        return sb.ToString();
    }

    private static string Answers(ReviewItem item)
    {
        if (item.Requirement.Kind == RequirementKind.Deterministic)
        {
            return "deterministic: " + string.Join(", ", item.Deterministic.Select(d => $"{d.Check} {(d.Passed ? "match" : "no match")}"));
        }

        if (item.Decision is null)
        {
            return "not evaluated";
        }

        if (item.Decision.Failure is { } failure)
        {
            return $"failed: {failure.Kind}";
        }

        var d = item.Decision;
        return $"evidence_status={d.EvidenceStatus!.Label} ({P(d.EvidenceStatus.Confidence)}), " +
               $"requirement_result={d.Result!.Label} ({P(d.Result.Confidence)}), " +
               $"materiality={P(d.Materiality!.Score)}";
    }

    private static string Word(ReviewItem item)
    {
        if (item.Draft?.Proposal is { } proposal)
        {
            return proposal.ReplacementText is null
                ? $"comment on \"{Escape(proposal.TargetText)}\""
                : $"comment + tracked change \"{Escape(proposal.TargetText)}\" -> \"{Escape(proposal.ReplacementText)}\"";
        }

        return item.Routing.Action switch
        {
            WordAction.None => "none",
            WordAction.ReviewComment => item.Evidence.HeadingParaId is null ? "report only" : "review comment on heading",
            _ => "host comment (no valid draft)",
        };
    }

    private static string P(double value) => value.ToString("0.00", CultureInfo.InvariantCulture);

    private static string Escape(string value) => value.Replace("|", "\\|", StringComparison.Ordinal);
}
