using System.Globalization;
using System.Net;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using Microsoft.Extensions.AI;

namespace OfficeAgent.Samples.RequirementsReview;

/// <summary>
/// An offline evaluator that returns decisions from a script. For workflow tests that need a
/// specific decision or fault without going through HTTP.
/// </summary>
public sealed class ScriptedRequirementEvaluator : IRequirementEvaluator
{
    private readonly Func<RegisteredRequirement, AnchoredDocumentEvidence, CancellationToken, Task<RequirementDecision>> _script;
    private readonly List<string> _calls = new();

    /// <summary>Initializes a new instance of the <see cref="ScriptedRequirementEvaluator"/> class.</summary>
    /// <param name="script">Produces a decision per call.</param>
    public ScriptedRequirementEvaluator(Func<RegisteredRequirement, AnchoredDocumentEvidence, CancellationToken, Task<RequirementDecision>> script) => _script = script;

    /// <summary>Gets the requirement ids evaluated, in order.</summary>
    public IReadOnlyList<string> Calls => _calls;

    /// <inheritdoc />
    public Task<RequirementDecision> EvaluateAsync(RegisteredRequirement requirement, AnchoredDocumentEvidence evidence, CancellationToken cancellationToken)
    {
        _calls.Add(requirement.Id);
        return _script(requirement, evidence, cancellationToken);
    }
}

/// <summary>A scripted answer in Jev's terms.</summary>
/// <param name="Evidence">evidence_status label.</param>
/// <param name="EvidenceProbability">Probability of that label; the rest is spread evenly.</param>
/// <param name="Result">requirement_result label.</param>
/// <param name="ResultProbability">Probability of that label; the rest is spread evenly.</param>
/// <param name="Materiality">Probability per materiality level, lowest first.</param>
public sealed record ScriptedAnswer(string Evidence, double EvidenceProbability, string Result, double ResultProbability, double[] Materiality)
{
    /// <summary>Scripted answers for the sample method statement. Invented for the offline demo, not Jev output.</summary>
    public static IReadOnlyDictionary<string, ScriptedAnswer> Demo { get; } = new Dictionary<string, ScriptedAnswer>(StringComparer.Ordinal)
    {
        ["REQ-INS-01"] = new("sufficient", 0.93, "satisfied", 0.91, new[] { 0.86, 0.10, 0.03, 0.01 }),
        ["REQ-REC-01"] = new("sufficient", 0.95, "violated", 0.92, new[] { 0.01, 0.04, 0.80, 0.15 }),
        ["REQ-RSP-01"] = new("insufficient", 0.88, "not_determined", 0.83, new[] { 0.10, 0.35, 0.45, 0.10 }),
    };

    /// <summary>Invented answers for the fictional supplier privacy proposal.</summary>
    public static IReadOnlyDictionary<string, ScriptedAnswer> PrivacyDemo { get; } = new Dictionary<string, ScriptedAnswer>(StringComparer.Ordinal)
    {
        ["REQ-PRIV-ACCESS"] = new("sufficient", 0.95, "violated", 0.94, new[] { 0.01, 0.04, 0.80, 0.15 }),
        ["REQ-PRIV-STORAGE"] = new("sufficient", 0.93, "violated", 0.91, new[] { 0.02, 0.06, 0.78, 0.14 }),
        ["REQ-PRIV-RETENTION"] = new("sufficient", 0.95, "violated", 0.94, new[] { 0.01, 0.04, 0.80, 0.15 }),
    };

    /// <summary>Renders the answer as a Jev <c>/v1/systemone</c> response body.</summary>
    /// <param name="model">The model name to report.</param>
    /// <returns>The JSON body.</returns>
    public JsonObject ToResponse(string model)
    {
        var score = Materiality.Select((p, i) => p * i).Sum();
        return new JsonObject
        {
            ["model"] = model,
            ["answers"] = new JsonObject
            {
                [QuestionSet.EvidenceStatusName] = Choice(Evidence, EvidenceProbability, QuestionSet.EvidenceStatusLabels),
                [QuestionSet.RequirementResultName] = Choice(Result, ResultProbability, QuestionSet.RequirementResultLabels),
                [QuestionSet.MaterialityName] = new JsonObject
                {
                    ["type"] = "score",
                    ["score"] = Math.Round(score, 4),
                    ["confidence"] = Math.Round(Confidence(Materiality), 4),
                    ["legend"] = new JsonObject(QuestionSet.MaterialityLevels.Select((l, i) =>
                        new KeyValuePair<string, JsonNode?>(i.ToString(CultureInfo.InvariantCulture), l))),
                    ["probabilities"] = new JsonObject(Materiality.Select((p, i) =>
                        new KeyValuePair<string, JsonNode?>(i.ToString(CultureInfo.InvariantCulture), p))),
                },
            },
            ["usage"] = new JsonObject { ["input_tokens"] = 0, ["output_tokens"] = 0 },
        };
    }

    /// <summary>
    /// Confidence as TypeSafe documents it: the peak probability rescaled so a uniform
    /// distribution is 0 and a certain one is 1, <c>(n * peak - 1) / (n - 1)</c>.
    /// </summary>
    /// <param name="probabilities">The distribution.</param>
    /// <returns>The confidence.</returns>
    public static double Confidence(IReadOnlyCollection<double> probabilities) =>
        Math.Clamp((probabilities.Count * probabilities.Max() - 1) / (probabilities.Count - 1), 0, 1);

    private static JsonObject Choice(string label, double probability, IReadOnlyList<string> labels)
    {
        var rest = (1 - probability) / (labels.Count - 1);
        var distribution = labels.Select(l => Math.Round(l == label ? probability : rest, 4)).ToList();
        return new JsonObject
        {
            ["type"] = "choice",
            ["choice"] = label,
            ["confidence"] = Math.Round(Confidence(distribution), 4),
            ["probabilities"] = new JsonObject(labels.Select((l, i) =>
                new KeyValuePair<string, JsonNode?>(l, distribution[i]))),
        };
    }
}

/// <summary>
/// Serves scripted responses in the Jev wire format, so offline runs exercise the real HTTP
/// adapter: request building, response validation, and the request-id header.
/// </summary>
public sealed class ScriptedJevHandler : HttpMessageHandler
{
    /// <summary>The model name reported by scripted responses.</summary>
    public const string ScriptedModel = "scripted-jev-wire";

    /// <summary>The provider label for audits of runs that use this handler.</summary>
    public const string ProviderLabel = "scripted-jev-wire (offline)";

    private readonly Func<string, HttpResponseMessage> _respond;
    private int _count;

    /// <summary>Initializes a new instance of the <see cref="ScriptedJevHandler"/> class with the demo answers.</summary>
    public ScriptedJevHandler()
        : this(id => ScriptedAnswer.Demo.TryGetValue(id, out var answer)
            ? Json(HttpStatusCode.OK, answer.ToResponse(ScriptedModel).ToJsonString())
            : Json(HttpStatusCode.InternalServerError, "{\"detail\":\"no scripted answer\"}"))
    {
    }

    /// <summary>Initializes a new instance of the <see cref="ScriptedJevHandler"/> class.</summary>
    /// <param name="respond">Maps a requirement id to a response.</param>
    public ScriptedJevHandler(Func<string, HttpResponseMessage> respond) => _respond = respond;

    /// <summary>Gets request bodies received, in order.</summary>
    public List<string> Bodies { get; } = new();

    /// <summary>Gets the authorization scheme of each request.</summary>
    public List<string?> Schemes { get; } = new();

    /// <summary>Creates a JSON response.</summary>
    /// <param name="status">The status.</param>
    /// <param name="body">The body.</param>
    /// <returns>The response.</returns>
    public static HttpResponseMessage Json(HttpStatusCode status, string body) => new(status)
    {
        Content = new StringContent(body, Encoding.UTF8, "application/json"),
    };

    /// <inheritdoc />
    protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
    {
        var body = request.Content is null ? string.Empty : await request.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
        Bodies.Add(body);
        Schemes.Add(request.Headers.Authorization?.Scheme);

        var id = (JsonNode.Parse(body)?["state"]?["requirement"]?["id"]?.GetValue<string>()) ?? string.Empty;
        var response = _respond(id);
        response.Headers.TryAddWithoutValidation(JevRequirementEvaluator.RequestIdHeader, $"req_scripted_{Interlocked.Increment(ref _count):D4}");
        return response;
    }
}

/// <summary>A scripted tool call.</summary>
/// <param name="ToolName">The tool the model "calls".</param>
/// <param name="Arguments">Its arguments.</param>
public sealed record ScriptedCall(string ToolName, IDictionary<string, object?> Arguments);

/// <summary>
/// An offline <see cref="IChatClient"/> for the drafter. On the first turn it returns the
/// scripted tool call for the requirement in the brief; after a tool result it returns text.
/// </summary>
public sealed class ScriptedChatClient : IChatClient
{
    /// <summary>The model id reported through <see cref="ChatClientMetadata"/>.</summary>
    public const string ScriptedModel = "scripted-drafter";

    private static readonly Regex RequirementLine = new(@"^Requirement: (\S+)", RegexOptions.Multiline, TimeSpan.FromMilliseconds(100));

    private readonly Func<string, ScriptedCall?> _script;
    private readonly ChatClientMetadata _metadata = new("scripted", null, ScriptedModel);
    private int _calls;

    /// <summary>Initializes a new instance of the <see cref="ScriptedChatClient"/> class with the demo script.</summary>
    public ScriptedChatClient()
        : this(DemoScript)
    {
    }

    /// <summary>Initializes a new instance of the <see cref="ScriptedChatClient"/> class.</summary>
    /// <param name="script">Maps a requirement id to a tool call, or null for none.</param>
    public ScriptedChatClient(Func<string, ScriptedCall?> script) => _script = script;

    /// <summary>Gets the tool names offered on each turn.</summary>
    public List<IReadOnlyList<string>> OfferedTools { get; } = new();

    /// <summary>Gets every message text the client received.</summary>
    public List<string> Transcript { get; } = new();

    /// <summary>Invented comments for the legacy and privacy offline demos.</summary>
    /// <param name="requirementId">The requirement.</param>
    /// <returns>The call, or null.</returns>
    public static ScriptedCall? DemoScript(string requirementId) => requirementId switch
    {
        "REQ-REC-01" => new ScriptedCall(AgentProposalDrafter.ToolName, new Dictionary<string, object?>
        {
            ["evidenceId"] = "REQ-REC-01:1",
            ["targetText"] = "the team will assess the situation and agree next steps",
            ["replacementText"] = "",
            ["commentText"] = "Please list the recovery actions if installation fails, for example making the unit safe and reinstating the existing unit, and name the role responsible for each action.",
        }),
        "REQ-PRIV-ACCESS" => new ScriptedCall(AgentProposalDrafter.ToolName, new Dictionary<string, object?>
        {
            ["evidenceId"] = "REQ-PRIV-ACCESS:1",
            ["targetText"] = "for all project members",
            ["replacementText"] = "",
            ["commentText"] = "Please name the project roles allowed to access these logs instead of sharing them with all project members.",
        }),
        "REQ-PRIV-STORAGE" => new ScriptedCall(AgentProposalDrafter.ToolName, new Dictionary<string, object?>
        {
            ["evidenceId"] = "REQ-PRIV-STORAGE:1",
            ["targetText"] = "shared project folder",
            ["replacementText"] = "",
            ["commentText"] = "Please identify the restricted location where these logs will be stored.",
        }),
        "REQ-PRIV-RETENTION" => new ScriptedCall(AgentProposalDrafter.ToolName, new Dictionary<string, object?>
        {
            ["evidenceId"] = "REQ-PRIV-RETENTION:1",
            ["targetText"] = "kept after handover in case they are needed",
            ["replacementText"] = "",
            ["commentText"] = "Please state when these logs will be deleted after handover.",
        }),
        _ => null,
    };

    /// <inheritdoc />
    public Task<ChatResponse> GetResponseAsync(IEnumerable<ChatMessage> messages, ChatOptions? options = null, CancellationToken cancellationToken = default)
    {
        var list = messages.ToList();
        OfferedTools.Add((options?.Tools ?? new List<AITool>()).Select(t => t.Name).ToList());
        Transcript.AddRange(list.Select(m => m.Text));

        if (list.Any(m => m.Contents.OfType<FunctionResultContent>().Any()))
        {
            return Task.FromResult(new ChatResponse(new ChatMessage(ChatRole.Assistant, "Submitted.")) { ModelId = ScriptedModel });
        }

        var brief = list.LastOrDefault(m => m.Role == ChatRole.User)?.Text ?? string.Empty;
        var match = RequirementLine.Match(brief);
        var call = match.Success ? _script(match.Groups[1].Value) : null;
        if (call is null)
        {
            return Task.FromResult(new ChatResponse(new ChatMessage(ChatRole.Assistant, "No proposal.")) { ModelId = ScriptedModel });
        }

        var content = new FunctionCallContent($"call_{Interlocked.Increment(ref _calls)}", call.ToolName, call.Arguments);
        return Task.FromResult(new ChatResponse(new ChatMessage(ChatRole.Assistant, new List<AIContent> { content })) { ModelId = ScriptedModel });
    }

    /// <inheritdoc />
    public async IAsyncEnumerable<ChatResponseUpdate> GetStreamingResponseAsync(
        IEnumerable<ChatMessage> messages,
        ChatOptions? options = null,
        [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        var response = await GetResponseAsync(messages, options, cancellationToken).ConfigureAwait(false);
        foreach (var update in response.ToChatResponseUpdates())
        {
            yield return update;
        }
    }

    /// <inheritdoc />
    public object? GetService(Type serviceType, object? serviceKey = null) =>
        serviceKey is not null ? null
        : serviceType == typeof(ChatClientMetadata) ? _metadata
        : serviceType.IsInstanceOfType(this) ? this
        : null;

    /// <inheritdoc />
    public void Dispose()
    {
    }
}
