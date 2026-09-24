using System.Diagnostics;
using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;

namespace OfficeAgent.Samples.RequirementsReview;

/// <summary>Settings for the TypeSafe Jev adapter. The API key is not one of them.</summary>
public sealed record JevOptions
{
    /// <summary>The documented endpoint (OpenAPI 0.2.0, checked 2026-09-23).</summary>
    public static readonly Uri DefaultEndpoint = new("https://api.typesafe.ai/v1/systemone");

    /// <summary>Gets the endpoint.</summary>
    public Uri Endpoint { get; init; } = DefaultEndpoint;

    /// <summary>Gets the provider label recorded in the audit. Scripted transports must say so here.</summary>
    public string ProviderName { get; init; } = JevRequirementEvaluator.ProviderName;

    /// <summary>Gets the model to request. The model the provider reports is audited separately.</summary>
    public required string Model { get; init; }

    /// <summary>Gets the per-call timeout.</summary>
    public TimeSpan Timeout { get; init; } = TimeSpan.FromSeconds(30);

    /// <summary>Gets the largest evidence payload, in characters, that may leave the host.</summary>
    public int MaxEvidenceCharacters { get; init; } = 4000;

    /// <summary>Gets the largest response body accepted, in bytes.</summary>
    public int MaxResponseBytes { get; init; } = 64 * 1024;

    /// <summary>Gets how many times a 429 or 529 response is retried before it counts as a failure.</summary>
    public int MaxRetries { get; init; } = 2;

    /// <summary>Gets the first backoff delay; it doubles on each retry unless the response sends <c>Retry-After</c>.</summary>
    public TimeSpan RetryBaseDelay { get; init; } = TimeSpan.FromSeconds(1);

    /// <summary>Gets the longest single backoff delay. A longer <c>Retry-After</c> ends the retries.</summary>
    public TimeSpan MaxRetryDelay { get; init; } = TimeSpan.FromSeconds(8);
}

/// <summary>
/// Calls <c>POST /v1/systemone</c> with the fixed <see cref="QuestionSet"/> and validates the
/// answer against it. Anything unexpected becomes a failure; nothing unexpected becomes a pass.
/// </summary>
public sealed class JevRequirementEvaluator : IRequirementEvaluator
{
    /// <summary>The provider name recorded in the audit.</summary>
    public const string ProviderName = "typesafe-jev";

    /// <summary>The response header that carries the request id.</summary>
    public const string RequestIdHeader = "x-typesafe-request-id";

    private const double ProbabilityTolerance = 0.02;

    private readonly HttpClient _http;
    private readonly JevOptions _options;
    private readonly string _apiKey;
    private readonly ILogger _logger;

    /// <summary>Initializes a new instance of the <see cref="JevRequirementEvaluator"/> class.</summary>
    /// <param name="http">The HTTP client. Its own timeout is not relied on.</param>
    /// <param name="options">Adapter settings.</param>
    /// <param name="apiKey">The API key, read by the host from its secret store. Never logged or audited.</param>
    /// <param name="logger">Optional logger.</param>
    public JevRequirementEvaluator(HttpClient http, JevOptions options, string apiKey, ILogger? logger = null)
    {
        if (string.IsNullOrWhiteSpace(apiKey))
        {
            throw new ArgumentException("An API key is required for the live evaluator.", nameof(apiKey));
        }

        _http = http;
        _options = options;
        _apiKey = apiKey;
        _logger = logger ?? NullLogger.Instance;
    }

    /// <summary>Builds the state object: the requirement and its evidence text, nothing else.</summary>
    /// <param name="requirement">The requirement.</param>
    /// <param name="evidence">The evidence.</param>
    /// <returns>The state.</returns>
    public static JsonObject BuildState(RegisteredRequirement requirement, AnchoredDocumentEvidence evidence) => new()
    {
        // No document name, paragraph ids, authors, comments, or other sections. The
        // evidence ids are host-assigned labels that carry no document content.
        ["requirement"] = new JsonObject
        {
            ["id"] = requirement.Id,
            ["version"] = requirement.Version,
            ["text"] = requirement.Text,
        },
        ["evidence"] = new JsonArray(evidence.Items
            .Select(i => (JsonNode?)new JsonObject { ["id"] = i.EvidenceId, ["text"] = i.Text })
            .ToArray()),
    };

    /// <inheritdoc />
    public async Task<RequirementDecision> EvaluateAsync(
        RegisteredRequirement requirement,
        AnchoredDocumentEvidence evidence,
        CancellationToken cancellationToken)
    {
        if (evidence.Items.Count == 0 || evidence.Characters > _options.MaxEvidenceCharacters)
        {
            // The workflow checks this first. The adapter is the data boundary, so it
            // refuses as well rather than trusting its caller.
            return RequirementDecision.Failed(requirement.Id, _options.ProviderName, new EvaluationFailure(
                EvaluationFailureKind.NotSent,
                evidence.Items.Count == 0 ? "no evidence to send" : "evidence exceeds the external-evaluation budget"));
        }

        var state = BuildState(requirement, evidence);
        var stateJson = state.ToJsonString();
        var body = new JsonObject
        {
            ["state"] = state,
            ["model"] = _options.Model,
            ["questions"] = QuestionSet.Build(),
        };

        var sent = new RequirementDecision
        {
            RequirementId = requirement.Id,
            Provider = _options.ProviderName,
            StateSha256 = RequirementSet.Hash(stateJson),
            EvidenceIdsSent = evidence.Items.Select(i => i.EvidenceId).ToList(),
        };

        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeout.CancelAfter(_options.Timeout);
        var clock = Stopwatch.StartNew();

        var payload = body.ToJsonString();
        HttpResponseMessage response;
        while (true)
        {
            sent = sent with { Attempts = sent.Attempts + 1 };
            using var request = new HttpRequestMessage(HttpMethod.Post, _options.Endpoint)
            {
                Content = new StringContent(payload, Encoding.UTF8, "application/json"),
            };
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _apiKey);
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            try
            {
                response = await _http.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, timeout.Token).ConfigureAwait(false);
            }
            catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
            {
                return Fail(sent, clock, null, new EvaluationFailure(EvaluationFailureKind.Timeout, $"no response within {_options.Timeout.TotalSeconds:0} s"));
            }
            catch (HttpRequestException ex)
            {
                return Fail(sent, clock, null, new EvaluationFailure(EvaluationFailureKind.ProviderUnavailable, $"transport error: {ex.GetType().Name}"));
            }

            // TypeSafe asks clients to back off and retry 429 (rate limited) and 529
            // (overloaded). Retries stay inside the call's timeout; when they run out the
            // response is handled below like any other error, which routes to a person.
            if (RetryDelay(response, sent.Attempts, _options.Timeout - clock.Elapsed) is not { } delay)
            {
                break;
            }

            _logger.LogWarning(
                "Jev returned {Status} for {RequirementId}; retry {Retry} in {Delay} ms",
                (int)response.StatusCode, requirement.Id, sent.Attempts, (long)delay.TotalMilliseconds);
            response.Dispose();

            try
            {
                await Task.Delay(delay, timeout.Token).ConfigureAwait(false);
            }
            catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
            {
                return Fail(sent, clock, null, new EvaluationFailure(EvaluationFailureKind.Timeout, $"no answer within {_options.Timeout.TotalSeconds:0} s"));
            }
        }

        using (response)
        {
            var requestId = response.Headers.TryGetValues(RequestIdHeader, out var ids) ? ids.FirstOrDefault() : null;
            var status = (int)response.StatusCode;

            string text;
            try
            {
                text = await ReadBoundedAsync(response.Content, _options.MaxResponseBytes, timeout.Token).ConfigureAwait(false);
            }
            catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
            {
                return Fail(sent, clock, requestId, new EvaluationFailure(EvaluationFailureKind.Timeout, "response body not received in time", status));
            }
            catch (InvalidDataException ex)
            {
                return Fail(sent, clock, requestId, new EvaluationFailure(EvaluationFailureKind.MalformedResponse, ex.Message, status));
            }
            catch (HttpRequestException ex)
            {
                return Fail(sent, clock, requestId, new EvaluationFailure(EvaluationFailureKind.ProviderUnavailable, $"transport error: {ex.GetType().Name}", status));
            }

            if (!response.IsSuccessStatusCode)
            {
                return Fail(sent, clock, requestId, new EvaluationFailure(KindFor(response.StatusCode), ErrorSummary(status, text), status));
            }

            var decision = Parse(sent, text, out var problem);
            if (problem is not null)
            {
                return Fail(sent, clock, requestId, new EvaluationFailure(EvaluationFailureKind.MalformedResponse, problem, status));
            }

            _logger.LogInformation(
                "Jev evaluated {RequirementId}: request {RequestId}, model {Model}, {Elapsed} ms",
                requirement.Id, requestId ?? "(none)", decision!.Model, clock.ElapsedMilliseconds);

            return decision with { RequestId = requestId, Elapsed = clock.Elapsed };
        }
    }

    /// <summary>Validates a response body against the fixed questions.</summary>
    /// <param name="sent">The decision shell describing what was sent.</param>
    /// <param name="text">The response body.</param>
    /// <param name="problem">Why the body was refused, or null.</param>
    /// <returns>The decision, or null when refused.</returns>
    public static RequirementDecision? Parse(RequirementDecision sent, string text, out string? problem)
    {
        JsonNode? root;
        try
        {
            root = JsonNode.Parse(text);
        }
        catch (JsonException)
        {
            problem = "response is not JSON";
            return null;
        }

        if (root is not JsonObject obj)
        {
            problem = "response is not a JSON object";
            return null;
        }

        if (obj["model"] is not JsonValue modelValue || !modelValue.TryGetValue<string>(out var model) || string.IsNullOrWhiteSpace(model))
        {
            problem = "response has no model";
            return null;
        }

        if (obj["answers"] is not JsonObject answers)
        {
            problem = "response has no answers object";
            return null;
        }

        var expected = new[] { QuestionSet.EvidenceStatusName, QuestionSet.RequirementResultName, QuestionSet.MaterialityName };
        var extra = answers.Select(a => a.Key).Except(expected, StringComparer.Ordinal).ToList();
        if (extra.Count > 0)
        {
            problem = $"response answers questions that were not asked: {string.Join(", ", extra)}";
            return null;
        }

        var evidence = ParseChoice(answers, QuestionSet.EvidenceStatusName, QuestionSet.EvidenceStatusLabels, out problem);
        if (evidence is null)
        {
            return null;
        }

        var result = ParseChoice(answers, QuestionSet.RequirementResultName, QuestionSet.RequirementResultLabels, out problem);
        if (result is null)
        {
            return null;
        }

        var materiality = ParseScore(answers, QuestionSet.MaterialityName, QuestionSet.MaterialityLevels.Count, out problem);
        if (materiality is null)
        {
            return null;
        }

        int? inputTokens = null, outputTokens = null;
        if (obj["usage"] is JsonObject usage)
        {
            inputTokens = Int(usage["input_tokens"]);
            outputTokens = Int(usage["output_tokens"]);
        }

        problem = null;
        return sent with
        {
            Model = model,
            EvidenceStatus = evidence,
            Result = result,
            Materiality = materiality,
            InputTokens = inputTokens,
            OutputTokens = outputTokens,
        };
    }

    private static ChoiceJudgement? ParseChoice(JsonObject answers, string name, IReadOnlyList<string> labels, out string? problem)
    {
        if (answers[name] is not JsonObject answer)
        {
            problem = $"{name}: missing";
            return null;
        }

        if (Str(answer["type"]) != "choice")
        {
            problem = $"{name}: expected a choice answer";
            return null;
        }

        var choice = Str(answer["choice"]);
        if (choice is null || !labels.Contains(choice, StringComparer.Ordinal))
        {
            problem = $"{name}: choice is not one of the allowed labels";
            return null;
        }

        if (Unit(answer["confidence"]) is not { } confidence)
        {
            problem = $"{name}: confidence missing or outside [0,1]";
            return null;
        }

        var probabilities = Probabilities(answer["probabilities"], labels, out problem);
        if (probabilities is null)
        {
            problem = $"{name}: {problem}";
            return null;
        }

        problem = null;
        return new ChoiceJudgement(choice, confidence, probabilities);
    }

    private static ScoreJudgement? ParseScore(JsonObject answers, string name, int levels, out string? problem)
    {
        if (answers[name] is not JsonObject answer)
        {
            problem = $"{name}: missing";
            return null;
        }

        if (Str(answer["type"]) != "score")
        {
            problem = $"{name}: expected a score answer";
            return null;
        }

        if (Num(answer["score"]) is not { } score || score < 0 || score > levels - 1)
        {
            problem = $"{name}: score missing or outside [0,{levels - 1}]";
            return null;
        }

        if (Unit(answer["confidence"]) is not { } confidence)
        {
            problem = $"{name}: confidence missing or outside [0,1]";
            return null;
        }

        var keys = Enumerable.Range(0, levels).Select(i => i.ToString(System.Globalization.CultureInfo.InvariantCulture)).ToList();
        var probabilities = Probabilities(answer["probabilities"], keys, out problem);
        if (probabilities is null)
        {
            problem = $"{name}: {problem}";
            return null;
        }

        problem = null;
        return new ScoreJudgement(score, confidence, probabilities);
    }

    private static IReadOnlyDictionary<string, double>? Probabilities(JsonNode? node, IReadOnlyList<string> keys, out string? problem)
    {
        if (node is not JsonObject map)
        {
            problem = "probabilities missing";
            return null;
        }

        var actual = map.Select(p => p.Key).ToHashSet(StringComparer.Ordinal);
        if (!actual.SetEquals(keys))
        {
            problem = "probabilities do not cover exactly the allowed outcomes";
            return null;
        }

        var values = new Dictionary<string, double>(StringComparer.Ordinal);
        foreach (var key in keys)
        {
            if (Unit(map[key]) is not { } p)
            {
                problem = "a probability is missing or outside [0,1]";
                return null;
            }

            values[key] = p;
        }

        if (Math.Abs(values.Values.Sum() - 1.0) > ProbabilityTolerance)
        {
            problem = "probabilities do not sum to 1";
            return null;
        }

        problem = null;
        return values;
    }

    private static string? Str(JsonNode? node) =>
        node is JsonValue v && v.TryGetValue<string>(out var s) ? s : null;

    private static double? Num(JsonNode? node)
    {
        if (node is not JsonValue v || v.GetValueKind() != JsonValueKind.Number)
        {
            return null;
        }

        var d = v.GetValue<double>();
        return double.IsFinite(d) ? d : null;
    }

    private static double? Unit(JsonNode? node) => Num(node) is { } d && d >= 0 && d <= 1 ? d : null;

    private static int? Int(JsonNode? node) =>
        node is JsonValue v && v.GetValueKind() == JsonValueKind.Number && v.TryGetValue<int>(out var i) ? i : null;

    /// <summary>The delay before retrying, or null when the response is final.</summary>
    private TimeSpan? RetryDelay(HttpResponseMessage response, int attempts, TimeSpan remaining)
    {
        var status = (int)response.StatusCode;
        if (status is not (429 or 529) || attempts > _options.MaxRetries)
        {
            return null;
        }

        TimeSpan delay;
        if (response.Headers.RetryAfter is { } retryAfter)
        {
            // The server's wait is honoured, never shortened. A wait longer than the cap is
            // not worth holding the review for: the requirement goes to a person instead.
            delay = retryAfter.Delta ?? (retryAfter.Date is { } date ? date - DateTimeOffset.UtcNow : TimeSpan.Zero);
            if (delay > _options.MaxRetryDelay)
            {
                return null;
            }
        }
        else
        {
            delay = _options.RetryBaseDelay * Math.Pow(2, attempts - 1);
            if (delay > _options.MaxRetryDelay)
            {
                delay = _options.MaxRetryDelay;
            }
        }

        if (delay < TimeSpan.Zero)
        {
            delay = TimeSpan.Zero;
        }

        return delay < remaining ? delay : null;
    }

    private static EvaluationFailureKind KindFor(HttpStatusCode status) => (int)status switch
    {
        401 or 403 => EvaluationFailureKind.Authentication,
        400 or 422 => EvaluationFailureKind.RequestRejected,
        429 => EvaluationFailureKind.RateLimited,
        _ => EvaluationFailureKind.ProviderUnavailable,
    };

    /// <summary>
    /// A short error summary. Only the provider's error type is kept; a 422 detail can echo
    /// the submitted input, which is document content, so it is not copied into the audit.
    /// </summary>
    private static string ErrorSummary(int status, string text)
    {
        string? errorType = null;
        try
        {
            if (JsonNode.Parse(text) is JsonObject obj && obj["detail"] is JsonObject detail)
            {
                errorType = Str(detail["error_type"]);
            }
        }
        catch (JsonException)
        {
            // Not JSON; the status is enough.
        }

        return errorType is { Length: > 0 and <= 64 } ? $"HTTP {status} ({errorType})" : $"HTTP {status}";
    }

    private static async Task<string> ReadBoundedAsync(HttpContent content, int maxBytes, CancellationToken cancellationToken)
    {
        await using var stream = await content.ReadAsStreamAsync(cancellationToken).ConfigureAwait(false);
        using var buffer = new MemoryStream();
        var chunk = new byte[8192];
        int read;
        while ((read = await stream.ReadAsync(chunk, cancellationToken).ConfigureAwait(false)) > 0)
        {
            if (buffer.Length + read > maxBytes)
            {
                throw new InvalidDataException($"response exceeds {maxBytes} bytes");
            }

            buffer.Write(chunk, 0, read);
        }

        return Encoding.UTF8.GetString(buffer.GetBuffer(), 0, (int)buffer.Length);
    }

    private RequirementDecision Fail(RequirementDecision sent, Stopwatch clock, string? requestId, EvaluationFailure failure)
    {
        _logger.LogWarning(
            "Jev evaluation failed for {RequirementId}: {Kind} {Detail} (request {RequestId})",
            sent.RequirementId, failure.Kind, failure.Detail, requestId ?? "(none)");

        return sent with { Failure = failure, RequestId = requestId, Elapsed = clock.Elapsed };
    }
}
