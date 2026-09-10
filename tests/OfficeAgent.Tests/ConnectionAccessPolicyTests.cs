using System.Security.Claims;
using System.Text.Json;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

public sealed class ConnectionAccessPolicyTests
{
    [Fact]
    public async Task Every_connection_bearing_tool_fails_closed_before_provider_lookup()
    {
        var policy = new RecordingPolicy(_ => false);
        var tools = Tools(policy);
        var calls = new Func<Task<string>>[]
        {
            () => tools.InspectDocument("private", "document"),
            () => tools.FindInDocument("private", "document", "text"),
            () => tools.PreviewPlan("private", "document", "[]"),
            () => tools.ApplyPlan("private", "document", "[]"),
            () => tools.RegisterDocument("private", "document.docx"),
            () => tools.CreateDocument("private", "document.docx"),
            () => tools.OpenDocument("private", "document.docx"),
            () => tools.EditDocument("private", "document.docx", "[]"),
            () => tools.ImportDocumentContent("private", "document.docx", "AA=="),
            () => tools.ExportDocumentContent("private", "document"),
            () => tools.PopulateTemplateBatch("private", "document", "{\"items\":[]}"),
            () => tools.CompareDocuments("private", "document", "other", "revised"),
            () => tools.PreviewDocumentMerge("{\"sources\":[{\"connectionId\":\"private\",\"documentId\":\"one\"},{\"connectionId\":\"other\",\"documentId\":\"two\"}]}"),
            () => tools.MergeDocuments("{\"inputs\":[{\"document\":{\"connectionId\":\"private\",\"itemId\":\"one\"}}]}", "out", "packet.docx"),
            () => tools.RemoveDocument("private", "document")
        };

        foreach (var call in calls) AssertError(await call(), "connection-forbidden");

        Assert.Contains(ConnectionCapability.Read, policy.Requests);
        Assert.Contains(ConnectionCapability.Register, policy.Requests);
        Assert.Contains(ConnectionCapability.Create, policy.Requests);
        Assert.Contains(ConnectionCapability.Delete, policy.Requests);
    }

    [Theory]
    [InlineData("Replace", ConnectionCapability.Edit)]
    [InlineData("NewVersion", ConnectionCapability.Create)]
    [InlineData("NewDocument", ConnectionCapability.Create)]
    public async Task Apply_requires_read_and_the_save_capability(
        string saveMode,
        ConnectionCapability saveCapability)
    {
        var policy = new RecordingPolicy(capability => capability == ConnectionCapability.Read);

        var result = await Tools(policy).ApplyPlan("private", "document", "[]", saveMode);

        AssertError(result, "connection-forbidden");
        Assert.Equal(new[] { ConnectionCapability.Read, saveCapability }, policy.Requests);
    }

    [Fact]
    public async Task Compare_requires_read_access_to_both_connections()
    {
        var policy = new PerConnectionPolicy(connectionId => connectionId == "original");
        var result = await new OfficeAgentTools(
            new OfficeAgentClient(new WordModule()), policy, new FixedPrincipalAccessor())
            .CompareDocuments("original", "one", "revised", "two");

        AssertError(result, "connection-forbidden");
        Assert.Equal(new[] { "original", "revised" }, policy.Connections);
    }

    private static OfficeAgentTools Tools(IConnectionAccessPolicy policy) => new(
        new OfficeAgentClient(new WordModule()),
        policy,
        new FixedPrincipalAccessor());

    [Fact]
    public async Task Merge_checks_all_sources_and_destination_before_provider_lookup()
    {
        const string request = "{\"sources\":[{\"connectionId\":\"allowed\",\"documentId\":\"one\"},{\"connectionId\":\"denied\",\"documentId\":\"two\"}]}";
        var policy = new PerConnectionPolicy(id => id == "allowed");
        AssertError(await Tools(policy).PreviewDocumentMerge(request), "connection-forbidden");
        Assert.Equal(new[] { "allowed", "denied" }, policy.Connections);
        var createPolicy = new RecordingPolicy(capability => capability == ConnectionCapability.Read);
        const string plan = "{\"inputs\":[{\"document\":{\"connectionId\":\"one\",\"itemId\":\"a\"}},{\"document\":{\"connectionId\":\"two\",\"itemId\":\"b\"}}]}";
        AssertError(await Tools(createPolicy).MergeDocuments(plan, "destination", "packet.docx"), "connection-forbidden");
        Assert.Equal(new[] { ConnectionCapability.Read, ConnectionCapability.Read, ConnectionCapability.Create }, createPolicy.Requests);
    }

    private static void AssertError(string json, string expectedCode)
    {
        using var document = JsonDocument.Parse(json);
        Assert.Contains(document.RootElement.GetProperty("errors").EnumerateArray(),
            error => error.GetProperty("Code").GetString() == expectedCode);
    }

    private sealed class RecordingPolicy(Func<ConnectionCapability, bool> decision)
        : IConnectionAccessPolicy
    {
        public List<ConnectionCapability> Requests { get; } = new();

        public ValueTask<bool> IsAllowedAsync(
            ClaimsPrincipal principal,
            string connectionId,
            ConnectionCapability capability,
            CancellationToken cancellationToken = default)
        {
            Assert.Equal("trusted-user", principal.FindFirstValue(ClaimTypes.NameIdentifier));
            Requests.Add(capability);
            return new ValueTask<bool>(decision(capability));
        }
    }

    private sealed class FixedPrincipalAccessor : ITrustedPrincipalAccessor
    {
        public ClaimsPrincipal Principal { get; } = new(new ClaimsIdentity(new[]
        {
            new Claim(ClaimTypes.NameIdentifier, "trusted-user")
        }, "test"));
    }

    private sealed class PerConnectionPolicy(Func<string, bool> decision) : IConnectionAccessPolicy
    {
        public List<string> Connections { get; } = new();

        public ValueTask<bool> IsAllowedAsync(
            ClaimsPrincipal principal,
            string connectionId,
            ConnectionCapability capability,
            CancellationToken cancellationToken = default)
        {
            Connections.Add(connectionId);
            return new ValueTask<bool>(decision(connectionId));
        }
    }
}
