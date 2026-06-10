using System.Text.Json;
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

public class AiToolsTests
{
    [Fact]
    public void AsAIFunctions_exposes_only_the_read_and_edit_tools()
    {
        using var workspace = new ToolsWorkspace();
        var tools = new OfficeAgentTools(workspace.Client);
        var functions = tools.AsAIFunctions();

        // The agent never registers, deletes, or otherwise manages storage. The host
        // pre-registers documents and threads the id into the system prompt; the agent
        // only inspects, searches, previews, and applies.
        var names = functions.Select(f => f.Name).ToArray();
        Assert.Equal(4, functions.Length);
        Assert.Contains("inspect_document", names);
        Assert.Contains("find_in_document", names);
        Assert.Contains("preview_plan", names);
        Assert.Contains("apply_plan", names);
        Assert.DoesNotContain("add_document", names);
        Assert.DoesNotContain("remove_document", names);
    }

    [Fact]
    public async Task ApplyPlan_round_trips_a_polymorphic_plan_through_storage()
    {
        using var workspace = new ToolsWorkspace();
        var tools = new OfficeAgentTools(workspace.Client);
        var added = await workspace.Client.RegisterBytesAsync("workspace", workspace.Root, DocxFactory.Contract(), "contract.docx");

        var planJson =
            "{ \"operations\": [ { \"op\": \"setProperty\", " +
            "\"target\": { \"$anchor\": \"node\", \"kind\": \"docProperty\", \"path\": \"core/subject\" }, " +
            "\"value\": \"Adapted for Globex\" } ] }";

        var report = await tools.ApplyPlan("workspace", added.ItemId, planJson);
        using var parsed = JsonDocument.Parse(report);
        Assert.True(parsed.RootElement.GetProperty("committed").GetBoolean());

        var outputId = parsed.RootElement.GetProperty("outputDocumentId").GetString()!;
        Assert.Equal("contract.v2.docx", parsed.RootElement.GetProperty("outputName").GetString());
        Assert.Equal(
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            parsed.RootElement.GetProperty("outputContentType").GetString());
        using var content = await workspace.Client.OpenReadAsync(OfficeAgent.Abstractions.DocumentReference.ForFileSystem("workspace", outputId));
        using var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(content.Stream, false);
        Assert.Equal("Adapted for Globex", doc.PackageProperties.Subject);
    }

    [Fact]
    public async Task Tools_inspect_and_find_round_trip_through_storage()
    {
        using var workspace = new ToolsWorkspace();
        var tools = new OfficeAgentTools(workspace.Client);
        var added = await workspace.Client.RegisterBytesAsync("workspace", workspace.Root, DocxFactory.Contract(), "contract.docx");

        using (var inspect = JsonDocument.Parse(await tools.InspectDocument("workspace", added.ItemId)))
        {
            Assert.Equal("Word", inspect.RootElement.GetProperty("format").GetString());
            Assert.Equal(4, inspect.RootElement.GetProperty("paragraphs").GetArrayLength());
        }

        var firstHit = JsonDocument.Parse(await tools.FindInDocument("workspace", added.ItemId, "Acme Corp"))
            .RootElement[0];
        Assert.Equal("Acme Corp", firstHit.GetProperty("expect").GetString());
    }

    [Fact]
    public async Task PreviewPlan_returns_structured_error_for_invalid_json()
    {
        using var workspace = new ToolsWorkspace();
        var tools = new OfficeAgentTools(workspace.Client);
        var added = await workspace.Client.RegisterBytesAsync("workspace", workspace.Root, DocxFactory.Contract(), "contract.docx");

        var report = await tools.PreviewPlan("workspace", added.ItemId, "{ not json");
        using var parsed = JsonDocument.Parse(report);

        Assert.False(parsed.RootElement.GetProperty("isValid").GetBoolean());
        Assert.Equal("invalid-json", parsed.RootElement.GetProperty("errors")[0].GetProperty("Code").GetString());
    }

    [Fact]
    public async Task InspectDocument_honours_fidelity_and_paging()
    {
        using var workspace = new ToolsWorkspace();
        var tools = new OfficeAgentTools(workspace.Client);
        var added = await workspace.Client.RegisterBytesAsync("workspace", workspace.Root, DocxFactory.Contract(), "contract.docx");

        using var outline = JsonDocument.Parse(await tools.InspectDocument("workspace", added.ItemId, fidelity: "outline"));
        Assert.Equal(0, outline.RootElement.GetProperty("paragraphs").GetArrayLength());

        using var page = JsonDocument.Parse(await tools.InspectDocument("workspace", added.ItemId, paragraphOffset: 1, paragraphLimit: 2));
        Assert.Equal(2, page.RootElement.GetProperty("paragraphs").GetArrayLength());
        Assert.Equal(4, page.RootElement.GetProperty("paragraphsTotal").GetInt32());
    }

    [Fact]
    public async Task ApplyPlan_unknown_documentId_returns_not_found()
    {
        using var workspace = new ToolsWorkspace();
        var tools = new OfficeAgentTools(workspace.Client);

        var planJson = "{ \"operations\": [] }";
        var report = await tools.ApplyPlan("workspace", "ffffffffffffffff", planJson);
        using var parsed = JsonDocument.Parse(report);

        Assert.False(parsed.RootElement.GetProperty("isValid").GetBoolean());
        Assert.Equal("not-found", parsed.RootElement.GetProperty("errors")[0].GetProperty("Code").GetString());
    }

    private sealed class ToolsWorkspace : IDisposable
    {
        private readonly ServiceProvider _serviceProvider;

        public ToolsWorkspace()
        {
            Root = Path.Combine(Path.GetTempPath(), $"officeagent-tools-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Root);

            var services = new ServiceCollection();
            services.AddWordFormat();
            services.AddFileSystemDocumentProvider("workspace", Root);
            services.AddOfficeAgent();
            _serviceProvider = services.BuildServiceProvider();
            Client = _serviceProvider.GetRequiredService<OfficeAgentClient>();
        }

        public string Root { get; }
        public OfficeAgentClient Client { get; }

        public void Dispose()
        {
            _serviceProvider.Dispose();
            if (Directory.Exists(Root)) Directory.Delete(Root, recursive: true);
        }
    }
}
