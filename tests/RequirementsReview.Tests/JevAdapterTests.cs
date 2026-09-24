using System.Net;
using System.Text.Json.Nodes;
using OfficeAgent.Samples.RequirementsReview;
using Xunit;

namespace RequirementsReview.Tests;

/// <summary>
/// The Jev adapter against a scripted transport: the request it builds, the answers it
/// accepts, and every way a response can fail to be an answer.
/// </summary>
public sealed class JevAdapterTests
{
    private static async Task<(RequirementDecision Decision, ScriptedJevHandler Handler, Harness Harness)> EvaluateAsync(
        Func<string, HttpResponseMessage> respond,
        string requirementId = "REQ-REC-01",
        TimeSpan? timeout = null)
    {
        var harness = await Harness.CreateAsync();
        var handler = new ScriptedJevHandler(respond);
        var evidence = await EvidenceCollector.CollectAsync(harness.Client, Harness.ConnectionId, harness.DocumentId, harness.Requirements, default);
        var requirement = harness.Requirements.Requirements.Single(r => r.Id == requirementId);
        var decision = await harness.Jev(handler, timeout).EvaluateAsync(requirement, evidence.ByRequirement[requirementId], default);
        return (decision, handler, harness);
    }

    [Fact]
    public async Task Request_carries_only_the_requirement_its_evidence_and_the_fixed_questions()
    {
        var (_, handler, harness) = await EvaluateAsync(id => Script.Ok(Script.DemoBody(id)));
        using var _h = harness;

        var body = Assert.IsType<JsonObject>(JsonNode.Parse(Assert.Single(handler.Bodies)));
        Assert.Equal(new[] { "model", "questions", "state" }, body.Select(p => p.Key).OrderBy(k => k));
        Assert.Equal("jev-latest", body["model"]!.GetValue<string>());
        Assert.Equal("Bearer", Assert.Single(handler.Schemes));

        var questions = body["questions"]!.AsObject();
        Assert.Equal(new[] { "evidence_status", "materiality", "requirement_result" }, questions.Select(p => p.Key).OrderBy(k => k));
        Assert.Equal(QuestionSet.EvidenceStatusLabels, questions["evidence_status"]!["criteria"]!.AsObject().Select(p => p.Key));
        Assert.Equal(QuestionSet.RequirementResultLabels, questions["requirement_result"]!["criteria"]!.AsObject().Select(p => p.Key));
        Assert.Equal("score", questions["materiality"]!["type"]!.GetValue<string>());

        var state = body["state"]!.AsObject();
        Assert.Equal(new[] { "evidence", "requirement" }, state.Select(p => p.Key).OrderBy(k => k));
        Assert.Equal(new[] { "id", "text", "version" }, state["requirement"]!.AsObject().Select(p => p.Key).OrderBy(k => k));
        Assert.All(state["evidence"]!.AsArray(), e => Assert.Equal(new[] { "id", "text" }, e!.AsObject().Select(p => p.Key).OrderBy(k => k)));

        // Data minimization: no paragraph ids, no document name, no other sections.
        var raw = handler.Bodies[0];
        Assert.DoesNotContain("w14:", raw);
        Assert.DoesNotContain(MethodStatementFixture.FileName, raw, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("AHU-3", raw);
        Assert.DoesNotContain("Hold point", raw);
        Assert.DoesNotContain("Priya Raman", raw);
    }

    [Fact]
    public async Task A_valid_answer_is_parsed_with_its_request_id_and_model()
    {
        var (decision, _, harness) = await EvaluateAsync(id => Script.Ok(Script.DemoBody(id)));
        using var _h = harness;

        Assert.True(decision.IsComplete);
        Assert.Equal("req_scripted_0001", decision.RequestId);
        Assert.Equal("jev-test", decision.Model);
        Assert.Equal(("sufficient", "violated"), (decision.EvidenceStatus!.Label, decision.Result!.Label));
        Assert.Equal(2.09, decision.Materiality!.Score, 3);
        Assert.Equal(0.925, decision.EvidenceStatus.Confidence, 3);
        Assert.Equal(new[] { "REQ-REC-01:1", "REQ-REC-01:2" }, decision.EvidenceIdsSent);
        Assert.NotNull(decision.StateSha256);
    }

    [Fact]
    public void Scripted_confidence_follows_the_documented_definition()
    {
        // The quick-start example in TypeSafe's docs: probabilities 0.85 / 0.15 / 0.0 over
        // three options come back with confidence 0.78, that is 0.775 rounded.
        Assert.Equal(0.775, ScriptedAnswer.Confidence(new[] { 0.85, 0.15, 0.0 }), 6);
        Assert.Equal(0.0, ScriptedAnswer.Confidence(new[] { 0.25, 0.25, 0.25, 0.25 }), 6);
        Assert.Equal(1.0, ScriptedAnswer.Confidence(new[] { 0.0, 1.0, 0.0 }), 6);
    }

    [Theory]
    [InlineData(500, EvaluationFailureKind.ProviderUnavailable)]
    [InlineData(503, EvaluationFailureKind.ProviderUnavailable)]
    [InlineData(401, EvaluationFailureKind.Authentication)]
    [InlineData(403, EvaluationFailureKind.Authentication)]
    [InlineData(400, EvaluationFailureKind.RequestRejected)]
    [InlineData(422, EvaluationFailureKind.RequestRejected)]
    [InlineData(429, EvaluationFailureKind.RateLimited)]
    public async Task Http_errors_become_failures(int status, EvaluationFailureKind kind)
    {
        var (decision, _, harness) = await EvaluateAsync(_ => ScriptedJevHandler.Json((HttpStatusCode)status, "{\"detail\":{\"error_type\":\"x\"}}"));
        using var _h = harness;

        Assert.False(decision.IsComplete);
        Assert.Equal(kind, decision.Failure!.Kind);
        Assert.Equal(status, decision.Failure.HttpStatus);
        Assert.StartsWith("req_scripted_", decision.RequestId);
    }

    [Theory]
    [InlineData(429)]
    [InlineData(529)]
    public async Task Throttling_is_retried_with_backoff_and_then_answered(int status)
    {
        var calls = 0;
        var (decision, handler, harness) = await EvaluateAsync(id => ++calls == 1
            ? ScriptedJevHandler.Json((HttpStatusCode)status, "{}")
            : Script.Ok(Script.DemoBody(id)));
        using var _h = harness;

        Assert.True(decision.IsComplete);
        Assert.Equal(2, decision.Attempts);
        Assert.Equal(2, handler.Bodies.Count);
        Assert.Equal(handler.Bodies[0], handler.Bodies[1]);
    }

    [Theory]
    [InlineData(429, EvaluationFailureKind.RateLimited)]
    [InlineData(529, EvaluationFailureKind.ProviderUnavailable)]
    public async Task Persistent_throttling_goes_to_a_person_after_the_allowed_retries(int status, EvaluationFailureKind kind)
    {
        var (decision, handler, harness) = await EvaluateAsync(_ => ScriptedJevHandler.Json((HttpStatusCode)status, "{}"));
        using var _h = harness;

        Assert.Equal(kind, decision.Failure!.Kind);
        Assert.Equal(3, decision.Attempts);
        Assert.Equal(3, handler.Bodies.Count);
    }

    [Fact]
    public async Task A_short_retry_after_is_honoured()
    {
        var calls = 0;
        var (decision, handler, harness) = await EvaluateAsync(id =>
        {
            if (++calls > 1)
            {
                return Script.Ok(Script.DemoBody(id));
            }

            var throttled = ScriptedJevHandler.Json((HttpStatusCode)429, "{}");
            throttled.Headers.RetryAfter = new System.Net.Http.Headers.RetryConditionHeaderValue(TimeSpan.Zero);
            return throttled;
        });
        using var _h = harness;

        Assert.True(decision.IsComplete);
        Assert.Equal(2, handler.Bodies.Count);
    }

    [Fact]
    public async Task A_retry_after_longer_than_the_cap_goes_to_a_person_without_retrying_early()
    {
        var (decision, handler, harness) = await EvaluateAsync(_ =>
        {
            var throttled = ScriptedJevHandler.Json((HttpStatusCode)429, "{}");
            throttled.Headers.RetryAfter = new System.Net.Http.Headers.RetryConditionHeaderValue(TimeSpan.FromSeconds(60));
            return throttled;
        });
        using var _h = harness;

        Assert.Equal(EvaluationFailureKind.RateLimited, decision.Failure!.Kind);
        Assert.Single(handler.Bodies);
    }

    [Fact]
    public async Task Other_errors_are_not_retried()
    {
        var (decision, handler, harness) = await EvaluateAsync(_ => ScriptedJevHandler.Json(HttpStatusCode.InternalServerError, "{}"));
        using var _h = harness;

        Assert.Equal(EvaluationFailureKind.ProviderUnavailable, decision.Failure!.Kind);
        Assert.Single(handler.Bodies);
    }

    [Fact]
    public async Task A_slow_provider_times_out()
    {
        var harness = await Harness.CreateAsync();
        using var _h = harness;
        var handler = new DelegatingStub(async (_, ct) =>
        {
            await Task.Delay(Timeout.Infinite, ct);
            return new HttpResponseMessage(HttpStatusCode.OK);
        });

        var decision = await Evaluate(harness, handler, TimeSpan.FromMilliseconds(200));

        Assert.Equal(EvaluationFailureKind.Timeout, decision.Failure!.Kind);
    }

    [Fact]
    public async Task A_network_failure_is_provider_unavailable()
    {
        var harness = await Harness.CreateAsync();
        using var _h = harness;
        var handler = new DelegatingStub((_, _) => throw new HttpRequestException("connection refused"));

        var decision = await Evaluate(harness, handler);

        Assert.Equal(EvaluationFailureKind.ProviderUnavailable, decision.Failure!.Kind);
    }

    [Fact]
    public async Task Caller_cancellation_is_not_turned_into_an_answer()
    {
        var harness = await Harness.CreateAsync();
        using var _h = harness;
        var handler = new DelegatingStub(async (_, ct) =>
        {
            await Task.Delay(Timeout.Infinite, ct);
            return new HttpResponseMessage(HttpStatusCode.OK);
        });

        using var cts = new CancellationTokenSource(TimeSpan.FromMilliseconds(100));
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => Evaluate(harness, handler, cancellationToken: cts.Token));
    }

    public static TheoryData<string, string> MalformedBodies()
    {
        JsonObject Demo() => ScriptedAnswer.Demo["REQ-REC-01"].ToResponse("jev-test");
        string With(Action<JsonObject> mutate)
        {
            var body = Demo();
            mutate(body);
            return body.ToJsonString();
        }

        return new TheoryData<string, string>
        {
            { "not json", "response is not JSON" },
            { "[]", "response is not a JSON object" },
            { With(b => b.Remove("model")), "response has no model" },
            { With(b => b.Remove("answers")), "response has no answers object" },
            { With(b => b["answers"]!.AsObject().Remove("materiality")), "materiality: missing" },
            { With(b => b["answers"]!["evidence_status"]!["type"] = "score"), "evidence_status: expected a choice answer" },
            { With(b => b["answers"]!["requirement_result"]!["choice"] = "compliant"), "requirement_result: choice is not one of the allowed labels" },
            { With(b => b["answers"]!["requirement_result"]!["probabilities"]!.AsObject().Remove("not_determined")), "requirement_result: probabilities do not cover exactly the allowed outcomes" },
            { With(b => b["answers"]!["requirement_result"]!["probabilities"]!["compliant"] = 0.0), "requirement_result: probabilities do not cover exactly the allowed outcomes" },
            { With(b => b["answers"]!["requirement_result"]!["probabilities"]!["violated"] = 0.5), "requirement_result: probabilities do not sum to 1" },
            { With(b => b["answers"]!["evidence_status"]!["confidence"] = 1.2), "evidence_status: confidence missing or outside [0,1]" },
            { With(b => b["answers"]!["evidence_status"]!["confidence"] = "high"), "evidence_status: confidence missing or outside [0,1]" },
            { With(b => b["answers"]!["is_compliant"] = new JsonObject { ["type"] = "noul", ["noul"] = 0.9 }), "response answers questions that were not asked: is_compliant" },
            { With(b => b["answers"]!["materiality"]!["score"] = 7), "materiality: score missing or outside [0,3]" },
            { With(b => b["answers"]!["materiality"]!["probabilities"] = new JsonObject { ["1"] = 0.25, ["2"] = 0.25, ["3"] = 0.25, ["4"] = 0.25 }), "materiality: probabilities do not cover exactly the allowed outcomes" },
        };
    }

    [Theory]
    [MemberData(nameof(MalformedBodies))]
    public async Task Malformed_responses_are_refused(string body, string problem)
    {
        var (decision, _, harness) = await EvaluateAsync(_ => Script.Ok(body));
        using var _h = harness;

        Assert.False(decision.IsComplete);
        Assert.Equal(EvaluationFailureKind.MalformedResponse, decision.Failure!.Kind);
        Assert.Equal(problem, decision.Failure.Detail);
        Assert.Null(decision.Result);
    }

    [Fact]
    public async Task An_oversized_response_is_refused()
    {
        var (decision, _, harness) = await EvaluateAsync(_ => Script.Ok("{\"pad\":\"" + new string('x', 70_000) + "\"}"));
        using var _h = harness;

        Assert.Equal(EvaluationFailureKind.MalformedResponse, decision.Failure!.Kind);
    }

    [Fact]
    public async Task Error_bodies_are_not_copied_into_the_failure()
    {
        // A 422 detail can echo the submitted input. Only the status and error type are kept.
        var echo = $"{{\"detail\":[{{\"msg\":\"bad\",\"input\":\"{MethodStatementFixture.RecoveryText} Bearer {Harness.FakeKey}\"}}]}}";
        var (decision, _, harness) = await EvaluateAsync(_ => ScriptedJevHandler.Json(HttpStatusCode.UnprocessableEntity, echo));
        using var _h = harness;

        Assert.Equal("HTTP 422", decision.Failure!.Detail);
        Assert.DoesNotContain(harness.Log.Lines, l => l.Contains(Harness.FakeKey, StringComparison.Ordinal) || l.Contains("the team will assess the situation and agree next steps", StringComparison.Ordinal));
    }

    [Fact]
    public async Task Evidence_over_the_budget_is_never_sent()
    {
        var harness = await Harness.CreateAsync();
        using var _h = harness;
        var handler = new ScriptedJevHandler();
        var evidence = await EvidenceCollector.CollectAsync(harness.Client, Harness.ConnectionId, harness.DocumentId, harness.Requirements, default);
        var evaluator = new JevRequirementEvaluator(new HttpClient(handler), new JevOptions { Model = "jev-latest", MaxEvidenceCharacters = 20 }, Harness.FakeKey);

        var decision = await evaluator.EvaluateAsync(harness.Requirements.Requirements.Single(r => r.Id == "REQ-REC-01"), evidence.ByRequirement["REQ-REC-01"], default);

        Assert.Equal(EvaluationFailureKind.NotSent, decision.Failure!.Kind);
        Assert.Empty(handler.Bodies);
    }

    [Fact]
    public void The_live_adapter_refuses_to_start_without_a_key()
    {
        Assert.Throws<ArgumentException>(() => new JevRequirementEvaluator(new HttpClient(), new JevOptions { Model = "jev-latest" }, " "));
    }

    private static async Task<RequirementDecision> Evaluate(Harness harness, HttpMessageHandler handler, TimeSpan? timeout = null, CancellationToken cancellationToken = default)
    {
        var evidence = await EvidenceCollector.CollectAsync(harness.Client, Harness.ConnectionId, harness.DocumentId, harness.Requirements, default);
        var evaluator = new JevRequirementEvaluator(
            new HttpClient(handler),
            new JevOptions { Model = "jev-latest", Timeout = timeout ?? TimeSpan.FromSeconds(30) },
            Harness.FakeKey);
        return await evaluator.EvaluateAsync(harness.Requirements.Requirements.Single(r => r.Id == "REQ-REC-01"), evidence.ByRequirement["REQ-REC-01"], cancellationToken);
    }
}

/// <summary>An HTTP handler driven by a delegate.</summary>
internal sealed class DelegatingStub : HttpMessageHandler
{
    private readonly Func<HttpRequestMessage, CancellationToken, Task<HttpResponseMessage>> _send;

    public DelegatingStub(Func<HttpRequestMessage, CancellationToken, Task<HttpResponseMessage>> send) => _send = send;

    protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) => _send(request, cancellationToken);
}
