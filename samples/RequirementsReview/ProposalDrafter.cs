using System.ComponentModel;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Microsoft.Agents.AI;
using Microsoft.Extensions.AI;

namespace OfficeAgent.Samples.RequirementsReview;

/// <summary>A drafted proposal that passed host validation.</summary>
/// <param name="ProposalId">Content-derived id.</param>
/// <param name="EvidenceId">The evidence item the proposal targets.</param>
/// <param name="ParaId">That item's paragraph id, resolved by the host.</param>
/// <param name="TargetText">Exact text in that paragraph to comment on or replace.</param>
/// <param name="ReplacementText">Proposed wording, or null for a comment only.</param>
/// <param name="CommentText">The rationale shown in the Word comment.</param>
public sealed record Proposal(
    string ProposalId,
    string EvidenceId,
    string ParaId,
    string TargetText,
    string? ReplacementText,
    string CommentText);

/// <summary>What the drafter produced.</summary>
/// <param name="Model">The drafting model id, when the client reports one.</param>
/// <param name="Proposal">The accepted proposal, or null.</param>
/// <param name="Rejections">Submissions the host refused, with reasons.</param>
/// <param name="Error">An error that stopped drafting, or null.</param>
public sealed record DraftResult(string? Model, Proposal? Proposal, IReadOnlyList<string> Rejections, string? Error);

/// <summary>Drafts rationale and wording for a violation the router already decided.</summary>
public interface IProposalDrafter
{
    /// <summary>Drafts one proposal.</summary>
    /// <param name="requirement">The violated requirement.</param>
    /// <param name="evidence">Its evidence.</param>
    /// <param name="allowReplacement">Whether replacement wording may be proposed.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    /// <returns>The draft result. Faults are returned, not thrown.</returns>
    Task<DraftResult> DraftAsync(
        RegisteredRequirement requirement,
        AnchoredDocumentEvidence evidence,
        bool allowReplacement,
        CancellationToken cancellationToken);
}

/// <summary>
/// Microsoft Agent Framework drafter. The agent gets one tool, <c>submit_proposal</c>, which
/// records text for the host to validate. It has no document tools, cannot see thresholds or
/// the policy, and cannot change the requirement, the route, or the plan.
/// </summary>
public sealed class AgentProposalDrafter : IProposalDrafter
{
    /// <summary>The only tool name the agent is given.</summary>
    public const string ToolName = "submit_proposal";

    private readonly IChatClient _chatClient;

    /// <summary>Initializes a new instance of the <see cref="AgentProposalDrafter"/> class.</summary>
    /// <param name="chatClient">The chat client. The agent adds function invocation itself.</param>
    public AgentProposalDrafter(IChatClient chatClient) => _chatClient = chatClient;

    /// <summary>Builds the agent's tool list for a requirement.</summary>
    /// <param name="collector">The collector the tool writes to.</param>
    /// <returns>Exactly one tool.</returns>
    public static IList<AITool> BuildTools(ProposalCollector collector) => new List<AITool> { collector.AsTool() };

    /// <inheritdoc />
    public async Task<DraftResult> DraftAsync(
        RegisteredRequirement requirement,
        AnchoredDocumentEvidence evidence,
        bool allowReplacement,
        CancellationToken cancellationToken)
    {
        var model = _chatClient.GetService<ChatClientMetadata>()?.DefaultModelId;
        var collector = new ProposalCollector(requirement, evidence, allowReplacement);

        try
        {
            var agent = new ChatClientAgent(
                _chatClient,
                instructions: Instructions,
                name: "RequirementsDrafter",
                description: "Drafts a review comment and, when allowed, replacement wording for a violation the host has already decided.",
                tools: BuildTools(collector));

            var session = await agent.CreateSessionAsync(cancellationToken).ConfigureAwait(false);
            await agent.RunAsync(Brief(requirement, evidence, allowReplacement), session, cancellationToken: cancellationToken)
                .ConfigureAwait(false);
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            throw;
        }
        catch (Exception ex)
        {
            // The violation stands whatever happens here; only the wording is lost.
            return new DraftResult(model, collector.Accepted, collector.Rejections, $"drafting failed: {ex.GetType().Name}");
        }

        return new DraftResult(model, collector.Accepted, collector.Rejections, null);
    }

    private const string Instructions =
        """
        You draft review feedback for one requirement in a Word document. The host has
        already decided that the requirement is violated; you do not re-decide it, and
        you cannot change the requirement, its evidence, or what happens to your draft.

        Call submit_proposal exactly once:
        - evidenceId: one of the evidence ids you are given, copied exactly.
        - targetText: the shortest exact phrase from that evidence item that causes the
          violation, copied character for character. It must appear once in that item.
        - replacementText: wording that replaces targetText so the sentence meets the
          requirement and still reads correctly, or an empty string when replacement is
          not allowed.
        - commentText: one or two sentences for the document author explaining what is
          missing against the requirement.

        Do not add facts that are not in the evidence or the requirement. The document
        author and a reviewer will read, accept, or reject your text.
        """;

    private static string Brief(RegisteredRequirement requirement, AnchoredDocumentEvidence evidence, bool allowReplacement)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"Requirement: {requirement.Id} (version {requirement.Version})");
        sb.AppendLine($"Requirement text: {requirement.Text}");
        sb.AppendLine("Decision: violated");
        sb.AppendLine(allowReplacement ? "Replacement wording: allowed" : "Replacement wording: not allowed (comment only)");
        sb.AppendLine();
        sb.AppendLine("Evidence:");
        foreach (var item in evidence.Items)
        {
            sb.AppendLine($"- evidenceId={item.EvidenceId} text=\"{item.Text}\"");
        }

        return sb.ToString();
    }
}

/// <summary>
/// Receives <c>submit_proposal</c> calls. Arguments are untrusted model output: every one is
/// validated against the evidence the host supplied, and only the first valid one is kept.
/// </summary>
public sealed class ProposalCollector
{
    /// <summary>Longest accepted target span.</summary>
    public const int MaxTarget = 200;

    /// <summary>Longest accepted replacement.</summary>
    public const int MaxReplacement = 300;

    /// <summary>Longest accepted comment.</summary>
    public const int MaxComment = 500;

    private readonly RegisteredRequirement _requirement;
    private readonly Dictionary<string, EvidenceItem> _evidence;
    private readonly bool _allowReplacement;
    private readonly List<string> _rejections = new();

    /// <summary>Initializes a new instance of the <see cref="ProposalCollector"/> class.</summary>
    /// <param name="requirement">The requirement.</param>
    /// <param name="evidence">The evidence the model was shown.</param>
    /// <param name="allowReplacement">Whether replacement wording is allowed.</param>
    public ProposalCollector(RegisteredRequirement requirement, AnchoredDocumentEvidence evidence, bool allowReplacement)
    {
        _requirement = requirement;
        _evidence = evidence.Items.ToDictionary(i => i.EvidenceId, StringComparer.Ordinal);
        _allowReplacement = allowReplacement;
    }

    /// <summary>Gets the accepted proposal.</summary>
    public Proposal? Accepted { get; private set; }

    /// <summary>Gets refused submissions.</summary>
    public IReadOnlyList<string> Rejections => _rejections;

    /// <summary>Projects the collector as the agent's only tool.</summary>
    /// <returns>The tool.</returns>
    public AIFunction AsTool() => AIFunctionFactory.Create(
        Submit,
        new AIFunctionFactoryOptions
        {
            Name = AgentProposalDrafter.ToolName,
            Description = "Submit the review comment and, when allowed, replacement wording for one evidence item. Returns whether the host accepted it.",
        });

    [Description("Submit a proposal.")]
    private string Submit(
        [Description("An evidence id exactly as given.")] string evidenceId,
        [Description("The exact phrase from that evidence item to comment on or replace.")] string targetText,
        [Description("Replacement wording, or an empty string when replacement is not allowed.")] string replacementText,
        [Description("One or two sentences explaining the gap.")] string commentText)
    {
        var error = Validate(evidenceId, targetText, replacementText, commentText, out var item);
        if (error is not null)
        {
            _rejections.Add(error);
            return JsonSerializer.Serialize(new { accepted = false, error });
        }

        var replacement = _allowReplacement && !string.IsNullOrWhiteSpace(replacementText) ? replacementText.Trim() : null;
        var comment = commentText.Trim();
        Accepted = new Proposal(
            ProposalId(_requirement.Id, evidenceId, targetText, replacement, comment),
            evidenceId,
            item!.ParaId,
            targetText,
            replacement,
            comment);

        return JsonSerializer.Serialize(new { accepted = true, proposalId = Accepted.ProposalId });
    }

    private string? Validate(string evidenceId, string targetText, string replacementText, string commentText, out EvidenceItem? item)
    {
        item = null;
        if (Accepted is not null)
        {
            return "a proposal was already accepted for this requirement";
        }

        if (evidenceId is null || !_evidence.TryGetValue(evidenceId, out item))
        {
            return "evidenceId is not one of the supplied evidence ids";
        }

        if (string.IsNullOrWhiteSpace(targetText) || targetText.Length > MaxTarget)
        {
            return $"targetText must be 1-{MaxTarget} characters";
        }

        var first = item.Text.IndexOf(targetText, StringComparison.Ordinal);
        if (first < 0)
        {
            return "targetText is not an exact substring of the cited evidence";
        }

        if (item.Text.IndexOf(targetText, first + 1, StringComparison.Ordinal) >= 0)
        {
            return "targetText appears more than once in the cited evidence";
        }

        if (!_allowReplacement && !string.IsNullOrWhiteSpace(replacementText))
        {
            return "replacement wording is not allowed for this requirement";
        }

        if (_allowReplacement && (string.IsNullOrWhiteSpace(replacementText) || replacementText.Length > MaxReplacement))
        {
            return $"replacementText must be 1-{MaxReplacement} characters";
        }

        if (string.IsNullOrWhiteSpace(commentText) || commentText.Length > MaxComment)
        {
            return $"commentText must be 1-{MaxComment} characters";
        }

        return null;
    }

    private static string ProposalId(string requirementId, string evidenceId, string target, string? replacement, string comment)
    {
        var material = string.Join("\u001f", requirementId, evidenceId, target, replacement ?? string.Empty, comment);
        return "prop_" + Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(material)))[..12].ToLowerInvariant();
    }
}
