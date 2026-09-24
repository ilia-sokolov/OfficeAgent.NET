using System.Runtime.CompilerServices;
using Microsoft.Extensions.AI;

namespace OfficeAgent.Samples.RequirementsReview;

/// <summary>
/// Offline chat client that calls evaluate_requirement once per supplied ID. The tool still
/// runs the real Jev HTTP adapter over the scripted transport; no answer comes from this model.
/// </summary>
public sealed class ScriptedReviewChatClient : IChatClient
{
    public const string ScriptedModel = "scripted-review-agent";

    private readonly ChatClientMetadata _metadata = new("scripted", null, ScriptedModel);
    private readonly Func<IReadOnlyList<string>, IReadOnlyList<string>> _sequence;
    private int _calls;

    public ScriptedReviewChatClient(Func<IReadOnlyList<string>, IReadOnlyList<string>>? sequence = null) =>
        _sequence = sequence ?? (ids => ids);

    public List<IReadOnlyList<string>> OfferedTools { get; } = new();

    public Task<ChatResponse> GetResponseAsync(
        IEnumerable<ChatMessage> messages,
        ChatOptions? options = null,
        CancellationToken cancellationToken = default)
    {
        var list = messages.ToList();
        OfferedTools.Add((options?.Tools ?? new List<AITool>()).Select(t => t.Name).ToList());

        const string prefix = "Requirement IDs: ";
        var prompt = list.FirstOrDefault(m => m.Role == ChatRole.User)?.Text ?? string.Empty;
        if (!prompt.StartsWith(prefix, StringComparison.Ordinal))
        {
            return Task.FromResult(Text("No registered requirements were supplied."));
        }

        var ids = prompt[prefix.Length..].Split(", ", StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        var scheduled = _sequence(ids);
        var completed = list.Sum(m => m.Contents.OfType<FunctionResultContent>().Count());
        if (completed >= scheduled.Count)
        {
            return Task.FromResult(Text("All requested tool calls have returned."));
        }

        var arguments = new Dictionary<string, object?> { ["requirementId"] = scheduled[completed] };
        var call = new FunctionCallContent($"review_call_{Interlocked.Increment(ref _calls)}", AgentRequirementReviewer.ToolName, arguments);
        return Task.FromResult(new ChatResponse(new ChatMessage(ChatRole.Assistant, new List<AIContent> { call }))
        {
            ModelId = ScriptedModel,
        });
    }

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

    public object? GetService(Type serviceType, object? serviceKey = null) =>
        serviceKey is not null ? null
        : serviceType == typeof(ChatClientMetadata) ? _metadata
        : serviceType.IsInstanceOfType(this) ? this
        : null;

    public void Dispose()
    {
    }

    private static ChatResponse Text(string value) =>
        new(new ChatMessage(ChatRole.Assistant, value)) { ModelId = ScriptedModel };
}
