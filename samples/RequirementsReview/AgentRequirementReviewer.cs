using System.Collections.Concurrent;
using Microsoft.Agents.AI;
using Microsoft.Extensions.AI;

namespace OfficeAgent.Samples.RequirementsReview;

/// <summary>Asks an Agent Framework agent to evaluate each eligible requirement through a Jev-backed tool.</summary>
public interface IRequirementReviewAgent
{
    /// <summary>Returns one decision per eligible requirement; missing or failed tool calls become failed decisions.</summary>
    Task<IReadOnlyDictionary<string, RequirementDecision>> ReviewAsync(
        IReadOnlyList<RegisteredRequirement> requirements,
        IReadOnlyDictionary<string, AnchoredDocumentEvidence> evidence,
        CancellationToken cancellationToken);
}

/// <summary>
/// The agent chooses when to call evaluate_requirement, but the host owns the registered
/// requirement, anchored evidence, Jev questions, and final routing decision.
/// </summary>
public sealed class AgentRequirementReviewer : IRequirementReviewAgent
{
    public const string ToolName = "evaluate_requirement";

    private readonly IChatClient _chatClient;
    private readonly IRequirementEvaluator _evaluator;

    public AgentRequirementReviewer(IChatClient chatClient, IRequirementEvaluator evaluator)
    {
        _chatClient = chatClient;
        _evaluator = evaluator;
    }

    public async Task<IReadOnlyDictionary<string, RequirementDecision>> ReviewAsync(
        IReadOnlyList<RegisteredRequirement> requirements,
        IReadOnlyDictionary<string, AnchoredDocumentEvidence> evidence,
        CancellationToken cancellationToken)
    {
        if (requirements.Count == 0)
        {
            return new Dictionary<string, RequirementDecision>(StringComparer.Ordinal);
        }

        var registered = requirements.ToDictionary(r => r.Id, StringComparer.Ordinal);
        var decisions = new ConcurrentDictionary<string, RequirementDecision>(StringComparer.Ordinal);
        var calls = new ConcurrentDictionary<string, int>(StringComparer.Ordinal);
        var invalidCalls = 0;

        async Task<string> EvaluateRequirementAsync(string requirementId, CancellationToken ct)
        {
            // The agent supplies only an ID. It cannot replace the registered question or
            // choose document text to send to Jev.
            if (!registered.TryGetValue(requirementId, out var requirement) ||
                !evidence.TryGetValue(requirementId, out var passage))
            {
                Interlocked.Increment(ref invalidCalls);
                return "Unknown or ineligible requirement; no evaluation was sent.";
            }

            if (calls.AddOrUpdate(requirementId, 1, (_, count) => count + 1) != 1)
            {
                return "Duplicate evaluation request; human review required.";
            }

            RequirementDecision decision;
            try
            {
                decision = await _evaluator.EvaluateAsync(requirement, passage, ct).ConfigureAwait(false);
            }
            catch (OperationCanceledException) when (ct.IsCancellationRequested)
            {
                throw;
            }
            catch (Exception ex)
            {
                decision = Failed(requirementId, EvaluationFailureKind.EvaluatorError, ex.GetType().Name);
            }

            if (!string.Equals(decision.RequirementId, requirementId, StringComparison.Ordinal))
            {
                decision = Failed(requirementId, EvaluationFailureKind.MalformedResponse, "decision is for a different requirement");
            }

            decisions[requirementId] = decision;
            return decision.IsComplete ? decision.Result!.Label : "needs_review";
        }

        var tool = AIFunctionFactory.Create(
            EvaluateRequirementAsync,
            new AIFunctionFactoryOptions
            {
                Name = ToolName,
                Description = "Evaluate one registered requirement against its fixed Word evidence using TypeSafe Jev.",
            });

        try
        {
            var agent = new ChatClientAgent(
                _chatClient,
                instructions: "Call evaluate_requirement once for every supplied requirement ID. " +
                              "Use the tool results; do not judge compliance yourself. Do not invent IDs.",
                name: "RequirementsReviewer",
                tools: new List<AITool> { tool });

            var session = await agent.CreateSessionAsync(cancellationToken).ConfigureAwait(false);
            await agent.RunAsync(
                "Requirement IDs: " + string.Join(", ", requirements.Select(r => r.Id)),
                session,
                cancellationToken: cancellationToken).ConfigureAwait(false);
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            throw;
        }
        catch (Exception ex)
        {
            // An unfinished agent run is not evidence that any requirement passed.
            return requirements.ToDictionary(
                r => r.Id,
                r => Failed(r.Id, EvaluationFailureKind.EvaluatorError, "agent run failed: " + ex.GetType().Name),
                StringComparer.Ordinal);
        }

        if (invalidCalls > 0)
        {
            return requirements.ToDictionary(
                r => r.Id,
                r => Failed(r.Id, EvaluationFailureKind.MalformedResponse, "agent requested an unknown or ineligible requirement"),
                StringComparer.Ordinal);
        }

        return requirements.ToDictionary(
            r => r.Id,
            r => calls.TryGetValue(r.Id, out var count) && count > 1
                ? Failed(r.Id, EvaluationFailureKind.MalformedResponse, "agent called evaluation more than once")
                : decisions.TryGetValue(r.Id, out var decision)
                    ? decision
                    : Failed(r.Id, EvaluationFailureKind.NotSent, "agent did not call evaluation"),
            StringComparer.Ordinal);
    }

    private static RequirementDecision Failed(string id, EvaluationFailureKind kind, string detail) =>
        RequirementDecision.Failed(id, "agent-tool", new EvaluationFailure(kind, detail));
}
