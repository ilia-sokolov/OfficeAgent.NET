using System.Text.Json;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting.Server;
using Microsoft.AspNetCore.Hosting.Server.Features;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging.Abstractions;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using OfficeAgent.Abstractions;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Mcp;

namespace OfficeAgent.Tests;

/// <summary>
/// The session connection on an HTTP server: one process serves every caller, so each MCP
/// session must see only the documents it added.
/// </summary>
public sealed class SessionPartitionTests
{
    [Fact]
    public async Task Two_mcp_sessions_on_one_http_server_cannot_see_each_others_documents()
    {
        var options = new OfficeAgentMcpOptions { EphemeralConnectionId = "session", AllowCreation = true };
        var builder = WebApplication.CreateBuilder(new[] { "--urls", "http://127.0.0.1:0" });
        builder.Services.AddHttpContextAccessor();
        builder.Services.AddMcpServer().WithHttpTransport().WithTools(OfficeAgentMcpServer.BuildToolset(options));
        await using var app = builder.Build();
        app.MapMcp();
        await app.StartAsync();
        var address = new Uri(app.Services.GetRequiredService<IServer>()
            .Features.Get<IServerAddressesFeature>()!.Addresses.Single());

        await using var alice = await ClientAsync(address);
        await using var bob = await ClientAsync(address);

        var imported = await CallAsync(alice, "import_document_content", new Dictionary<string, object?>
        {
            ["connectionId"] = "session",
            ["name"] = "alice.docx",
            ["contentBase64"] = Convert.ToBase64String(DocxFactory.Contract())
        });
        var aliceId = Property(imported, "documentId");

        // Alice exports her document.
        Assert.Contains("contentBase64", await CallAsync(alice, "export_document_content", new Dictionary<string, object?>
        {
            ["connectionId"] = "session", ["documentId"] = aliceId
        }));

        // Bob neither exports it by id nor learns that its name exists.
        Assert.Contains("not-found", await CallAsync(bob, "export_document_content", new Dictionary<string, object?>
        {
            ["connectionId"] = "session", ["documentId"] = aliceId
        }));
        var created = new Dictionary<string, object?>
        {
            ["connectionId"] = "session", ["name"] = "draft.docx", ["planJson"] = ""
        };
        Assert.True(Committed(await CallAsync(alice, "create_document", created)));
        var again = await CallAsync(alice, "create_document", created);
        Assert.False(Committed(again));
        Assert.Contains("already-exists", again);
        Assert.True(Committed(await CallAsync(bob, "create_document", created)));
    }

    private static bool Committed(string json)
    {
        using var document = JsonDocument.Parse(json);
        return document.RootElement.GetProperty("committed").GetBoolean();
    }

    [Fact]
    public async Task A_partitioned_memory_connection_hides_documents_from_other_partitions()
    {
        var partition = "alice";
        var provider = new MemoryDocumentProvider("session", new MemoryDocumentProviderOptions { Partition = () => partition });
        var reference = provider.Add("contract.docx", DocxFactory.Contract());

        partition = "bob";
        Assert.Empty(provider.List());
        Assert.Empty((await provider.ListAsync(new DocumentListQuery())).Items);
        Assert.Equal(ProviderErrorCode.NotFound,
            Assert.Throws<DocumentProviderException>(() => provider.Read(reference.ItemId)).Code);
        Assert.Equal(ProviderErrorCode.NotFound,
            (await Assert.ThrowsAsync<DocumentProviderException>(() => provider.RemoveAsync(reference))).Code);
        Assert.Equal(ProviderErrorCode.InvalidArgument,
            (await Assert.ThrowsAsync<DocumentProviderException>(() => provider.RegisterAsync(reference.ItemId))).Code);
        using (var content = new MemoryStream(DocxFactory.Contract()))
            await provider.CreateAsync("contract.docx", content); // the name is free in this partition

        partition = "alice";
        Assert.Single(provider.List());
        Assert.NotEmpty(provider.Read(reference.ItemId));
    }

    [Fact]
    public void An_http_request_without_a_session_id_gets_a_partition_of_its_own()
    {
        var first = new DefaultHttpContext { TraceIdentifier = "one" };
        var second = new DefaultHttpContext { TraceIdentifier = "two" };
        var withSession = new DefaultHttpContext();
        withSession.Request.Headers["Mcp-Session-Id"] = "abc";

        Assert.Null(OfficeAgentMcpServer.SessionPartition(null));
        Assert.NotEqual(OfficeAgentMcpServer.SessionPartition(first), OfficeAgentMcpServer.SessionPartition(second));
        Assert.Equal("session:abc", OfficeAgentMcpServer.SessionPartition(withSession));
    }

    private static async Task<McpClient> ClientAsync(Uri endpoint) =>
        await McpClient.CreateAsync(new HttpClientTransport(
            new HttpClientTransportOptions { Endpoint = endpoint, TransportMode = HttpTransportMode.StreamableHttp },
            NullLoggerFactory.Instance));

    private static async Task<string> CallAsync(McpClient client, string name, IReadOnlyDictionary<string, object?> arguments)
    {
        var result = await client.CallToolAsync(name, arguments);
        var text = Assert.Single(result.Content.OfType<TextContentBlock>()).Text;
        using var document = JsonDocument.Parse(text);
        return document.RootElement.ValueKind == JsonValueKind.String ? document.RootElement.GetString()! : text;
    }

    private static string Property(string json, string name)
    {
        using var document = JsonDocument.Parse(json);
        return document.RootElement.GetProperty(name).GetString()!;
    }
}
