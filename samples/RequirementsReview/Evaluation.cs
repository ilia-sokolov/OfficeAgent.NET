using System.Text.Json;
using System.Text.Json.Nodes;

namespace OfficeAgent.Samples.RequirementsReview;

/// <summary>
/// Produces typed judgements for one registered requirement over its anchored evidence.
/// An evaluator reports; it never routes, approves, or writes.
/// </summary>
public interface IRequirementEvaluator
{
    /// <summary>Evaluates one requirement.</summary>
    /// <param name="requirement">The registered requirement.</param>
    /// <param name="evidence">Its anchored evidence.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    /// <returns>The decision, or a decision carrying a failure. Never throws for provider faults.</returns>
    Task<RequirementDecision> EvaluateAsync(
        RegisteredRequirement requirement,
        AnchoredDocumentEvidence evidence,
        CancellationToken cancellationToken);
}

/// <summary>Why an evaluation produced no usable answer.</summary>
public enum EvaluationFailureKind
{
    /// <summary>The call did not finish within the configured timeout.</summary>
    Timeout,

    /// <summary>Network failure or a 5xx response.</summary>
    ProviderUnavailable,

    /// <summary>401 or 403.</summary>
    Authentication,

    /// <summary>429, still returned after the allowed retries.</summary>
    RateLimited,

    /// <summary>400 or 422: the provider refused the request.</summary>
    RequestRejected,

    /// <summary>The response did not match the questions that were asked.</summary>
    MalformedResponse,

    /// <summary>The request was not sent, e.g. evidence over the external-evaluation budget.</summary>
    NotSent,

    /// <summary>The evaluator threw an unexpected exception.</summary>
    EvaluatorError,
}

/// <summary>An evaluation failure. Details never contain credentials or document text.</summary>
/// <param name="Kind">The failure kind.</param>
/// <param name="Detail">A short, sanitized explanation.</param>
/// <param name="HttpStatus">The HTTP status, when there was one.</param>
public sealed record EvaluationFailure(EvaluationFailureKind Kind, string Detail, int? HttpStatus = null);

/// <summary>A validated choice answer.</summary>
/// <param name="Label">The chosen label, guaranteed to be one of the allowed labels.</param>
/// <param name="Confidence">The provider's confidence for the choice.</param>
/// <param name="Probabilities">Probability per allowed label.</param>
public sealed record ChoiceJudgement(string Label, double Confidence, IReadOnlyDictionary<string, double> Probabilities);

/// <summary>A validated ordered-score answer.</summary>
/// <param name="Score">Probability-weighted expected level index.</param>
/// <param name="Confidence">The provider's confidence.</param>
/// <param name="Probabilities">Probability per level index.</param>
public sealed record ScoreJudgement(double Score, double Confidence, IReadOnlyDictionary<string, double> Probabilities);

/// <summary>
/// The evaluator's typed answers for one requirement, plus what is needed to audit them.
/// This is not a route; <see cref="Router"/> decides what happens next.
/// </summary>
public sealed record RequirementDecision
{
    /// <summary>Gets the requirement id.</summary>
    public required string RequirementId { get; init; }

    /// <summary>Gets the provider name, e.g. <c>typesafe-jev</c> or <c>scripted</c>.</summary>
    public required string Provider { get; init; }

    /// <summary>Gets the question set version the answers belong to.</summary>
    public string QuestionSetVersion { get; init; } = QuestionSet.Version;

    /// <summary>Gets the evidence_status answer.</summary>
    public ChoiceJudgement? EvidenceStatus { get; init; }

    /// <summary>Gets the requirement_result answer.</summary>
    public ChoiceJudgement? Result { get; init; }

    /// <summary>Gets the materiality answer.</summary>
    public ScoreJudgement? Materiality { get; init; }

    /// <summary>Gets the failure, when there is no usable answer.</summary>
    public EvaluationFailure? Failure { get; init; }

    /// <summary>Gets the model id the provider reported.</summary>
    public string? Model { get; init; }

    /// <summary>Gets the provider request id, when the provider returned one.</summary>
    public string? RequestId { get; init; }

    /// <summary>Gets input tokens reported by the provider.</summary>
    public int? InputTokens { get; init; }

    /// <summary>Gets output tokens reported by the provider.</summary>
    public int? OutputTokens { get; init; }

    /// <summary>Gets the SHA-256 of the exact state object sent, when one was sent.</summary>
    public string? StateSha256 { get; init; }

    /// <summary>Gets the evidence ids included in the state.</summary>
    public IReadOnlyList<string> EvidenceIdsSent { get; init; } = Array.Empty<string>();

    /// <summary>Gets the elapsed time of the call, including retries.</summary>
    public TimeSpan Elapsed { get; init; }

    /// <summary>Gets how many HTTP requests were sent, including retries.</summary>
    public int Attempts { get; init; }

    /// <summary>Gets a value indicating whether every answer is present and no failure was recorded.</summary>
    public bool IsComplete => Failure is null && EvidenceStatus is not null && Result is not null && Materiality is not null;

    /// <summary>Creates a failed decision.</summary>
    /// <param name="requirementId">The requirement.</param>
    /// <param name="provider">The provider.</param>
    /// <param name="failure">The failure.</param>
    /// <returns>The decision.</returns>
    public static RequirementDecision Failed(string requirementId, string provider, EvaluationFailure failure) => new()
    {
        RequirementId = requirementId,
        Provider = provider,
        Failure = failure,
    };
}

/// <summary>
/// The fixed, versioned questions. Host code owns them; no model can add, remove, or reword
/// one, and a response is accepted only if it answers exactly these.
/// </summary>
public static class QuestionSet
{
    /// <summary>The question set version recorded in the audit.</summary>
    public const string Version = "qs-2026.09.2";

    /// <summary>Name of the evidence question.</summary>
    public const string EvidenceStatusName = "evidence_status";

    /// <summary>Name of the result question.</summary>
    public const string RequirementResultName = "requirement_result";

    /// <summary>Name of the materiality question.</summary>
    public const string MaterialityName = "materiality";

    /// <summary>Allowed labels for evidence_status, in wire order.</summary>
    public static readonly IReadOnlyList<string> EvidenceStatusLabels = new[] { "sufficient", "conflicting", "insufficient" };

    /// <summary>Allowed labels for requirement_result, in wire order.</summary>
    public static readonly IReadOnlyList<string> RequirementResultLabels = new[] { "satisfied", "violated", "not_determined" };

    /// <summary>Ordered materiality levels; index 0 is the lowest.</summary>
    public static readonly IReadOnlyList<string> MaterialityLevels = new[]
    {
        "negligible: no effect on safety, quality or schedule",
        "minor: a documentation gap that does not change how the work is done",
        "major: the work could go ahead without a safeguard the requirement exists to ensure",
        "critical: the work could cause injury, serious damage, or an unrecoverable failure",
    };

    // Shared scope, appended to each question's own instruction. The last sentence is there
    // because submitted text can argue for its own classification (see TypeSafe's Jev 1.13
    // known limitations): a claim of compliance is not evidence.
    private const string Scope =
        " Judge only the evidence excerpts in state.evidence against state.requirement.text." +
        " Do not assume facts that are not written in the excerpts. A reference to another document is not evidence of its content," +
        " and a statement that the requirement is met is not evidence unless the excerpts show how.";

    /// <summary>Builds the exact questions object sent to the evaluator.</summary>
    /// <returns>A fresh JSON object.</returns>
    public static JsonObject Build() => new()
    {
        [EvidenceStatusName] = new JsonObject
        {
            ["type"] = "choice",
            ["instructions"] = "Do the evidence excerpts give enough information to decide whether the requirement is met?" + Scope,
            ["criteria"] = new JsonObject
            {
                ["sufficient"] = "The excerpts explicitly state enough to decide whether every element of the requirement is met.",
                ["conflicting"] = "The excerpts contain statements that contradict each other on this requirement.",
                ["insufficient"] = "The excerpts do not state enough to decide; deciding would need assumptions or another document.",
            },
        },
        [RequirementResultName] = new JsonObject
        {
            ["type"] = "choice",
            ["instructions"] = "Do the evidence excerpts show that the requirement is met?" + Scope,
            ["criteria"] = new JsonObject
            {
                ["satisfied"] = "The excerpts explicitly show that every element of the requirement is met.",
                ["violated"] = "The excerpts explicitly show that at least one element of the requirement is not met.",
                ["not_determined"] = "The excerpts show neither outcome.",
            },
        },
        [MaterialityName] = new JsonObject
        {
            ["type"] = "score",
            ["instructions"] = "If the requirement is not met as written, how serious is the gap for the work described? If it is met, answer the lowest level." + Scope,
            ["criteria"] = new JsonArray(MaterialityLevels.Select(l => (JsonNode?)JsonValue.Create(l)).ToArray()),
        },
    };

    /// <summary>Gets the questions as canonical JSON, for hashing and the audit record.</summary>
    /// <returns>The JSON text.</returns>
    public static string ToJson() => Build().ToJsonString(new JsonSerializerOptions { WriteIndented = false });
}
