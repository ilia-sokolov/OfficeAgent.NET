using System.Text.RegularExpressions;

namespace OfficeAgent.Samples.RequirementsReview;

/// <summary>The result of one deterministic check.</summary>
/// <param name="Check">What was checked, e.g. <c>pattern[0]</c> or <c>passRequires[0]</c>.</param>
/// <param name="Pattern">The pattern text.</param>
/// <param name="Passed">True when the pattern matched.</param>
/// <param name="EvidenceIds">Evidence items the pattern matched in.</param>
/// <param name="Detail">A short explanation when the check did not pass.</param>
public sealed record DeterministicResult(
    string Check,
    string Pattern,
    bool Passed,
    IReadOnlyList<string> EvidenceIds,
    string? Detail);

/// <summary>
/// Runs registered patterns over anchored evidence. Local, repeatable, and never sent
/// anywhere. A timed-out pattern fails its check.
/// </summary>
public static class HardRuleEvaluator
{
    /// <summary>Runs a deterministic requirement's patterns.</summary>
    /// <param name="requirement">The requirement.</param>
    /// <param name="evidence">Its evidence.</param>
    /// <returns>One result per pattern, plus a section check.</returns>
    public static IReadOnlyList<DeterministicResult> Evaluate(RegisteredRequirement requirement, AnchoredDocumentEvidence evidence) =>
        Run("pattern", requirement.Patterns, evidence);

    /// <summary>Runs a semantic requirement's pass-blocking patterns.</summary>
    /// <param name="requirement">The requirement.</param>
    /// <param name="evidence">Its evidence.</param>
    /// <returns>One result per pattern.</returns>
    public static IReadOnlyList<DeterministicResult> PassSignals(RegisteredRequirement requirement, AnchoredDocumentEvidence evidence) =>
        Run("passRequires", requirement.PassRequires, evidence);

    private static IReadOnlyList<DeterministicResult> Run(string name, IReadOnlyList<Regex> patterns, AnchoredDocumentEvidence evidence)
    {
        var results = new List<DeterministicResult>();
        if (evidence.SectionMatches != 1)
        {
            results.Add(new DeterministicResult(
                "section",
                evidence.Section,
                false,
                Array.Empty<string>(),
                evidence.SectionMatches == 0 ? "section not found" : $"section heading appears {evidence.SectionMatches} times"));
            return results;
        }

        for (var i = 0; i < patterns.Count; i++)
        {
            var pattern = patterns[i];
            var hits = new List<string>();
            string? detail = null;
            try
            {
                hits.AddRange(evidence.Items.Where(item => pattern.IsMatch(item.Text)).Select(item => item.EvidenceId));
            }
            catch (RegexMatchTimeoutException)
            {
                hits.Clear();
                detail = "pattern timed out";
            }

            results.Add(new DeterministicResult(
                $"{name}[{i}]",
                pattern.ToString(),
                hits.Count > 0,
                hits,
                hits.Count > 0 ? null : detail ?? "no match in section"));
        }

        return results;
    }
}
