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

    private static OfficeAgentTools Tools(IConnectionAccessPolicy policy) => new(
        new OfficeAgentClient(new WordModule()),
        policy,
        new FixedPrincipalAccessor());

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
}
