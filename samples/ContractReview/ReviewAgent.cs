using System.ComponentModel;
using System.Text;
using System.Text.Json;
using Microsoft.Agents.AI;
using Microsoft.Extensions.AI;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;

namespace OfficeAgent.Samples.ContractReview;

/// <summary>
/// The judging half. The agent may read the document and record judgements; it
/// has no tool that writes. The write path lives in <see cref="PlanBuilder"/>.
/// </summary>
public sealed class ReviewAgent
{
    /// <summary>Tools the model is allowed to see from the OfficeAgent surface.</summary>
    public static readonly string[] ReadOnlyToolNames = { "inspect_document", "find_in_document" };

    private readonly IChatClient _chatClient;
    private readonly OfficeAgentClient _officeClient;

    /// <summary>Initializes a new instance of the <see cref="ReviewAgent"/> class.</summary>
    /// <param name="chatClient">A chat client with function invocation enabled.</param>
    /// <param name="officeClient">The OfficeAgent client.</param>
    public ReviewAgent(IChatClient chatClient, OfficeAgentClient officeClient)
    {
        _chatClient = chatClient;
        _officeClient = officeClient;
    }

    /// <summary>
    /// Builds the tool list handed to the agent: the read-only OfficeAgent tools
    /// plus the one tool that records a judgement.
    /// </summary>
    /// <param name="collector">Receives judgements.</param>
    /// <returns>The tools.</returns>
    public IList<AITool> BuildTools(FindingCollector collector)
    {
        // AsAIFunctions() returns the write tools too. They are filtered out here:
        // this single line is what makes "the model cannot write" true.
        var readOnly = new OfficeAgentTools(_officeClient)
            .AsAIFunctions()
            .Where(f => ReadOnlyToolNames.Contains(f.Name, StringComparer.Ordinal))
            .Cast<AITool>()
            .ToList();

        readOnly.Add(collector.AsTool());
        return readOnly;
    }

    /// <summary>Runs the review and returns the judgements the model recorded.</summary>
    /// <param name="connectionId">Connection holding the document.</param>
    /// <param name="documentId">Opaque document id.</param>
    /// <param name="playbook">The playbook.</param>
    /// <param name="screening">The screening result.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    /// <returns>The recorded findings.</returns>
    public async Task<IReadOnlyList<Finding>> JudgeAsync(
        string connectionId,
        string documentId,
        Playbook playbook,
        ScreeningResult screening,
        CancellationToken cancellationToken = default)
    {
        var collector = new FindingCollector();
        var tools = BuildTools(collector);

        var agent = new ChatClientAgent(
            _chatClient,
            instructions: Instructions(connectionId, documentId),
            name: "ContractReview",
            description: "Judges contract clauses against a playbook and records findings.",
            tools: tools);

        var session = await agent.CreateSessionAsync(cancellationToken).ConfigureAwait(false);
        await agent.RunAsync(Brief(playbook, screening), session, cancellationToken: cancellationToken)
            .ConfigureAwait(false);

        return collector.Findings;
    }

    private static string Instructions(string connectionId, string documentId) =>
        $"""
         You are a contract reviewer. You judge clauses against a playbook; you do
         not edit documents. The document is connectionId="{connectionId}",
         documentId="{documentId}".

         You are given a numbered list of candidates. Each one is a place in the
         contract where a playbook pattern matched. For every candidate, in order:

         1. Decide whether the rule is actually breached at that location. A pattern
            match is not a breach; read the surrounding wording and judge.
         2. Call report_finding exactly once for that candidate, quoting its
            candidateId verbatim.
         3. Give a one-sentence rationale a lawyer would accept, naming what the
            wording does and why it departs from the house position.
         4. When the rule permits a redline and you judge a breach, propose the
            smallest replacement wording that satisfies the house position, and
            pass it as replacementText. Replace only the matched text, and keep
            the surrounding sentence grammatical.

         Use inspect_document or find_in_document if you need more context than the
         candidate line gives you. Never invent a candidateId. Report every
         candidate, including the ones you judge compliant.
         """;

    private static string Brief(Playbook playbook, ScreeningResult screening)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"Playbook: {playbook.Name}");
        sb.AppendLine();
        sb.AppendLine("Rules:");
        foreach (var rule in playbook.Rules)
        {
            var redline = rule.Action == FindingAction.Redline ? "redline allowed" : "comment only";
            sb.AppendLine($"- {rule.Id} ({rule.Severity}, {redline}): {rule.Title}");
            sb.AppendLine($"  House position: {rule.Rationale}");
            if (!string.IsNullOrWhiteSpace(rule.PreferredWording))
            {
                sb.AppendLine($"  Preferred wording: {rule.PreferredWording}");
            }
        }

        sb.AppendLine();
        sb.AppendLine("Candidates:");
        foreach (var candidate in screening.Candidates)
        {
            sb.AppendLine(
                $"- candidateId={candidate.Id} rule={candidate.RuleId} " +
                $"matched=\"{candidate.MatchedText}\" context=\"{candidate.Context}\"");
        }

        sb.AppendLine();
        sb.AppendLine("Report a finding for every candidate listed above.");
        return sb.ToString();
    }
}

/// <summary>
/// Collects judgements. Findings are data the host owns; they feed the plan, the
/// report, and the tests without passing back through the model.
/// </summary>
public sealed class FindingCollector
{
    private readonly List<Finding> _findings = new();

    /// <summary>Gets the recorded findings, in the order the model reported them.</summary>
    public IReadOnlyList<Finding> Findings => _findings;

    /// <summary>Projects the collector as the single write-shaped tool the model gets.</summary>
    /// <returns>The tool.</returns>
    public AIFunction AsTool() => AIFunctionFactory.Create(
        Report,
        new AIFunctionFactoryOptions
        {
            Name = "report_finding",
            Description =
                "Record the judgement for one screened candidate. Call exactly once per " +
                "candidate, quoting its candidateId. Returns the recorded judgement.",
        });

    private const int MaxRationale = 500;

    private const int MaxReplacement = 300;

    private static string Clamp(string? value, int max)
    {
        var text = value?.Trim() ?? string.Empty;
        return text.Length <= max ? text : text[..max];
    }

    [Description("Record the judgement for one screened candidate.")]
    private string Report(
        [Description("The candidateId exactly as given in the candidate list.")] string candidateId,
        [Description("True when the playbook rule is actually breached at this location.")] bool violation,
        [Description("One sentence explaining the judgement, written for a lawyer.")] string rationale,
        [Description("Replacement wording when a redline is warranted; otherwise an empty string.")] string replacementText)
    {
        if (string.IsNullOrWhiteSpace(candidateId))
        {
            return JsonSerializer.Serialize(new { recorded = false, error = "candidateId is required" });
        }

        // The model can only put words in two places: a tracked insertion and a
        // comment. Both are bounded here, so a runaway generation cannot land a
        // page of text in a contract.
        var finding = new Finding
        {
            CandidateId = candidateId.Trim(),
            Violation = violation,
            Rationale = Clamp(rationale, MaxRationale),
            ReplacementText = string.IsNullOrWhiteSpace(replacementText)
                ? null
                : Clamp(replacementText, MaxReplacement),
        };

        _findings.Add(finding);

        return JsonSerializer.Serialize(new
        {
            recorded = true,
            candidateId = finding.CandidateId,
            violation = finding.Violation,
        });
    }
}
