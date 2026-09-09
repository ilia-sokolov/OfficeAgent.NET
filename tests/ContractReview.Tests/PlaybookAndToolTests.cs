using Microsoft.Extensions.AI;
using OfficeAgent.Samples.ContractReview;
using Xunit;

namespace ContractReview.Tests;

/// <summary>Playbook parsing, and the guarantee that the model gets no write tool.</summary>
public sealed class PlaybookAndToolTests
{
    [Fact]
    public void Playbook_parses_severity_and_action_names()
    {
        var playbook = TestPlaybooks.PaymentRedline();
        var rule = Assert.Single(playbook.Rules);

        Assert.Equal("PAY-01", rule.Id);
        Assert.Equal(Severity.High, rule.Severity);
        Assert.Equal(FindingAction.Redline, rule.Action);
    }

    [Theory]
    [InlineData("{ \"name\": \"x\", \"rules\": [] }", "no rules")]
    [InlineData("{ \"name\": \"x\", \"rules\": [ { \"id\": \"\", \"patterns\": [\"a\"], \"rationale\": \"r\" } ] }", "id")]
    [InlineData("{ \"name\": \"x\", \"rules\": [ { \"id\": \"A\", \"patterns\": [], \"rationale\": \"r\" } ] }", "pattern")]
    [InlineData("{ \"name\": \"x\", \"rules\": [ { \"id\": \"A\", \"patterns\": [\"a\"], \"rationale\": \"\" } ] }", "rationale")]
    [InlineData("{ \"name\": \"x\", \"rules\": [ { \"id\": \"A\", \"patterns\": [\"a\"], \"rationale\": \"r\" }, { \"id\": \"A\", \"patterns\": [\"b\"], \"rationale\": \"r\" } ] }", "Duplicate")]
    public void Malformed_playbooks_are_rejected_with_a_reason(string json, string expected)
    {
        var ex = Assert.Throws<InvalidDataException>(() => Playbook.FromJson(json));
        Assert.Contains(expected, ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Unknown_severity_is_rejected_rather_than_defaulted()
    {
        var json = "{ \"name\": \"x\", \"rules\": [ { \"id\": \"A\", \"severity\": \"catastrophic\", \"patterns\": [\"a\"], \"rationale\": \"r\" } ] }";
        Assert.Throws<InvalidDataException>(() => Playbook.FromJson(json));
    }

    [Fact]
    public async Task The_agent_is_given_no_tool_that_can_write()
    {
        using var harness = await ReviewHarness.CreateAsync();

        var agent = new ReviewAgent(new UnusedChatClient(), harness.Client);
        var tools = agent.BuildTools(new FindingCollector());

        var names = tools.OfType<AIFunction>().Select(f => f.Name).OrderBy(n => n, StringComparer.Ordinal).ToArray();

        Assert.Equal(new[] { "find_in_document", "inspect_document", "report_finding" }, names);
        Assert.DoesNotContain("apply_plan", names);
        Assert.DoesNotContain("preview_plan", names);
        Assert.DoesNotContain("edit_document", names);
    }

    private sealed class UnusedChatClient : IChatClient
    {
        public Task<ChatResponse> GetResponseAsync(
            IEnumerable<ChatMessage> messages,
            ChatOptions? options = null,
            CancellationToken cancellationToken = default) =>
            throw new NotSupportedException("This test never calls the model.");

        public IAsyncEnumerable<ChatResponseUpdate> GetStreamingResponseAsync(
            IEnumerable<ChatMessage> messages,
            ChatOptions? options = null,
            CancellationToken cancellationToken = default) =>
            throw new NotSupportedException("This test never calls the model.");

        public object? GetService(Type serviceType, object? serviceKey = null) => null;

        public void Dispose()
        {
        }
    }
}
