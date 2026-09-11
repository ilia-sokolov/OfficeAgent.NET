using System.Collections.Concurrent;
using System.Net;
using System.Text.Json;
using Microsoft.AspNetCore.Hosting.Server;
using Microsoft.AspNetCore.Hosting.Server.Features;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using OfficeAgent.Samples.HostedGateway;

namespace OfficeAgent.Tests;

public sealed class HostedGatewayTests
{
    [Fact]
    public async Task Authenticated_callers_are_isolated_by_connection_and_sensitive_values_are_not_logged()
    {
        using var roots = new GatewayRoots();
        var logs = new CapturingLoggerProvider();
        await using var app = HostedGatewayApp.Build(
            new[] { "--urls", "http://127.0.0.1:0" },
            builder =>
            {
                builder.Configuration.AddInMemoryCollection(new Dictionary<string, string?>
                {
                    ["HostedGateway:ConnectionARoot"] = roots.A,
                    ["HostedGateway:ConnectionBRoot"] = roots.B
                });
                builder.Logging.AddProvider(logs);
            });

        await app.StartAsync();
        var address = app.Services.GetRequiredService<IServer>()
            .Features.Get<IServerAddressesFeature>()!.Addresses.Single();

        using (var anonymous = new HttpClient())
        {
            var response = await anonymous.GetAsync(address);
            Assert.Equal(HttpStatusCode.Unauthorized, response.StatusCode);
        }

        await using var alice = await ClientAsync(new Uri(address), "alice-token");
        await using var bob = await ClientAsync(new Uri(address), "bob-token");

        Assert.Equal("connection-a", OnlyConnection(await CallAsync(alice, "list_connections")));
        Assert.Equal("connection-b", OnlyConnection(await CallAsync(bob, "list_connections")));

        const string sensitiveDocumentText = "HOSTED-DOCUMENT-CONTENT-MUST-NOT-BE-LOGGED";
        var aliceCreate = await CallAsync(alice, "create_document", new Dictionary<string, object?>
        {
            ["connectionId"] = "connection-a",
            ["name"] = "alice.docx",
            ["planJson"] = Plan(sensitiveDocumentText)
        });
        var aliceId = CreatedId(aliceCreate);
        using (var creation = JsonDocument.Parse(aliceCreate))
        {
            var actor = creation.RootElement.GetProperty("receipt").GetProperty("actor");
            Assert.Equal("alice", actor.GetProperty("subject").GetString());
            Assert.Equal("demo", actor.GetProperty("issuer").GetString());
        }
        var bobId = CreatedId(await CallAsync(bob, "create_document", new Dictionary<string, object?>
        {
            ["connectionId"] = "connection-b",
            ["name"] = "bob.docx"
        }));

        Assert.Equal("Word", Property(await CallAsync(alice, "inspect_document", new Dictionary<string, object?>
        {
            ["connectionId"] = "connection-a", ["documentId"] = aliceId
        }), "format"));
        Assert.Equal("Word", Property(await CallAsync(bob, "inspect_document", new Dictionary<string, object?>
        {
            ["connectionId"] = "connection-b", ["documentId"] = bobId
        }), "format"));

        AssertError(await CallAsync(alice, "inspect_document", new Dictionary<string, object?>
        {
            ["connectionId"] = "connection-b", ["documentId"] = bobId
        }), "connection-forbidden");
        AssertError(await CallAsync(bob, "inspect_document", new Dictionary<string, object?>
        {
            ["connectionId"] = "connection-a", ["documentId"] = aliceId
        }), "connection-forbidden");

        // A document id issued by one connection has no authority in another connection.
        AssertError(await CallAsync(alice, "inspect_document", new Dictionary<string, object?>
        {
            ["connectionId"] = "connection-a", ["documentId"] = bobId
        }), "not-found");

        await app.StopAsync();
        var applicationLogs = string.Join(Environment.NewLine, logs.Messages);
        Assert.DoesNotContain("alice-token", applicationLogs, StringComparison.Ordinal);
        Assert.DoesNotContain("bob-token", applicationLogs, StringComparison.Ordinal);
        Assert.DoesNotContain(sensitiveDocumentText, applicationLogs, StringComparison.Ordinal);
    }

    private static async Task<McpClient> ClientAsync(Uri endpoint, string token)
    {
        var transport = new HttpClientTransport(
            new HttpClientTransportOptions
            {
                Endpoint = endpoint,
                TransportMode = HttpTransportMode.StreamableHttp,
                AdditionalHeaders = new Dictionary<string, string>
                {
                    ["Authorization"] = "Bearer " + token
                }
            },
            NullLoggerFactory.Instance);
        return await McpClient.CreateAsync(transport);
    }

    private static async Task<string> CallAsync(
        McpClient client,
        string name,
        IReadOnlyDictionary<string, object?>? arguments = null)
    {
        var result = await client.CallToolAsync(name, arguments ?? new Dictionary<string, object?>());
        var text = Assert.Single(result.Content.OfType<TextContentBlock>()).Text;
        using var document = JsonDocument.Parse(text);
        return document.RootElement.ValueKind == JsonValueKind.String
            ? document.RootElement.GetString()!
            : text;
    }

    private static string OnlyConnection(string json)
    {
        using var document = JsonDocument.Parse(json);
        return Assert.Single(document.RootElement.EnumerateArray())
            .GetProperty("connectionId").GetString()!;
    }

    private static string CreatedId(string json) => Property(json, "outputDocumentId");

    private static string Property(string json, string name)
    {
        using var document = JsonDocument.Parse(json);
        return document.RootElement.GetProperty(name).GetString()!;
    }

    private static void AssertError(string json, string expectedCode)
    {
        using var document = JsonDocument.Parse(json);
        Assert.Contains(document.RootElement.GetProperty("errors").EnumerateArray(),
            error => error.GetProperty("Code").GetString() == expectedCode);
    }

    private static string Plan(string text) => $$"""
        { "operations": [ { "op": "insert",
            "target": { "paraId": "auto-0000", "expect": "" },
            "position": "Before", "text": "{{text}}" } ] }
        """;

    private sealed class GatewayRoots : IDisposable
    {
        private readonly string _root = Path.Combine(
            Path.GetTempPath(), "officeagent-hosted-tests-" + Guid.NewGuid().ToString("N"));

        public GatewayRoots()
        {
            A = Path.Combine(_root, "a");
            B = Path.Combine(_root, "b");
        }

        public string A { get; }
        public string B { get; }

        public void Dispose()
        {
            if (Directory.Exists(_root)) Directory.Delete(_root, recursive: true);
        }
    }

    private sealed class CapturingLoggerProvider : ILoggerProvider
    {
        public ConcurrentQueue<string> Messages { get; } = new();

        public ILogger CreateLogger(string categoryName) => new CapturingLogger(Messages);
        public void Dispose() { }

        private sealed class CapturingLogger(ConcurrentQueue<string> messages) : ILogger
        {
            public IDisposable? BeginScope<TState>(TState state) where TState : notnull
            {
                messages.Enqueue("SCOPE: " + state);
                return null;
            }
            public bool IsEnabled(LogLevel logLevel) => true;
            public void Log<TState>(LogLevel logLevel, EventId eventId, TState state,
                Exception? exception, Func<TState, Exception?, string> formatter) =>
                messages.Enqueue(formatter(state, exception));
        }
    }
}
