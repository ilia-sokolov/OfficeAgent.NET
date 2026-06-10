using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

public class AnchorJsonInferenceTests
{
    [Fact]
    public async Task PreviewPlan_accepts_format_target_without_anchor_discriminator()
    {
        // The shape an LLM tends to produce: omits "$anchor" inside target.
        using var workspace = new ToolsWorkspace();
        var tools = new OfficeAgentTools(workspace.Client);
        var added = await workspace.Client.RegisterBytesAsync("workspace", workspace.Root, DocxFactory.Contract(), "contract.docx");

        var paraId = JsonDocument.Parse(await tools.FindInDocument("workspace", added.ItemId, "Acme Corp"))
            .RootElement[0].GetProperty("paraId").GetString()!;

        var planJson = $$"""
            {
              "operations": [
                {
                  "op": "format",
                  "target": { "paraId": "{{paraId}}", "expect": "Acme Corp", "occurrence": 0 },
                  "highlight": "yellow",
                  "color": "FF0000"
                }
              ]
            }
            """;

        var report = await tools.PreviewPlan("workspace", added.ItemId, planJson);
        using var parsed = JsonDocument.Parse(report);

        Assert.True(parsed.RootElement.GetProperty("isValid").GetBoolean(),
            $"Expected isValid=true. Got: {report}");
        Assert.Equal(1, parsed.RootElement.GetProperty("changes").GetArrayLength());
    }

    [Fact]
    public async Task PreviewPlan_still_accepts_explicit_anchor_discriminator()
    {
        using var workspace = new ToolsWorkspace();
        var tools = new OfficeAgentTools(workspace.Client);
        var added = await workspace.Client.RegisterBytesAsync("workspace", workspace.Root, DocxFactory.Contract(), "contract.docx");

        var paraId = JsonDocument.Parse(await tools.FindInDocument("workspace", added.ItemId, "Acme Corp"))
            .RootElement[0].GetProperty("paraId").GetString()!;

        var planJson = $$"""
            {
              "operations": [
                {
                  "op": "changeText",
                  "target": { "$anchor": "textSpan", "paraId": "{{paraId}}", "expect": "Acme Corp", "occurrence": 0 },
                  "with": "Globex",
                  "mode": "Direct"
                }
              ]
            }
            """;

        var report = await tools.PreviewPlan("workspace", added.ItemId, planJson);
        using var parsed = JsonDocument.Parse(report);

        Assert.True(parsed.RootElement.GetProperty("isValid").GetBoolean());
    }

    [Fact]
    public async Task PreviewPlan_infers_node_anchor_for_insertTableRows()
    {
        using var workspace = new ToolsWorkspace();
        var tools = new OfficeAgentTools(workspace.Client);
        var added = await workspace.Client.RegisterBytesAsync("workspace", workspace.Root, BuildDocWithTable(), "table.docx");

        var planJson = """
            {
              "operations": [
                {
                  "op": "insertTableRows",
                  "target": { "kind": "table", "path": "table#0" },
                  "rows": [ ["new", "row"] ]
                }
              ]
            }
            """;

        var report = await tools.PreviewPlan("workspace", added.ItemId, planJson);
        using var parsed = JsonDocument.Parse(report);
        Assert.True(parsed.RootElement.GetProperty("isValid").GetBoolean(),
            $"Expected isValid=true. Got: {report}");
    }

    [Fact]
    public async Task PreviewPlan_infers_structural_anchor_for_fill()
    {
        using var workspace = new ToolsWorkspace();
        var tools = new OfficeAgentTools(workspace.Client);
        var added = await workspace.Client.RegisterBytesAsync("workspace", workspace.Root, DocxFactory.Contract(), "contract.docx");

        var planJson = $$"""
            {
              "operations": [
                {
                  "op": "fill",
                  "target": { "tag": "{{DocxFactory.ClientControlTag}}" },
                  "value": "Globex"
                }
              ]
            }
            """;

        var report = await tools.PreviewPlan("workspace", added.ItemId, planJson);
        using var parsed = JsonDocument.Parse(report);
        Assert.True(parsed.RootElement.GetProperty("isValid").GetBoolean(),
            $"Expected isValid=true. Got: {report}");
    }

    [Fact]
    public async Task User_reported_insertTableRows_payload_previews_cleanly()
    {
        // Locks in the LLM-ergonomics fixes:
        //  - No more "polymorphic discriminator required" internal-error (AnchorJsonConverter).
        //  - LLM-supplied contractVersion is tolerated (strict version check relaxed pre-1.0).
        //  - target.kind="table"+path="table#0" infers NodeAnchor with kind="table".
        using var workspace = new ToolsWorkspace();
        var tools = new OfficeAgentTools(workspace.Client);
        var added = await workspace.Client.RegisterBytesAsync("workspace", workspace.Root, BuildDocWithTable(), "table.docx");

        var planJson = """
            {
              "contractVersion": "1.0",
              "operations": [
                {
                  "op": "insertTableRows",
                  "target": { "kind": "table", "path": "table#0" },
                  "rows": [["ru","",""], ["us","",""], ["uk","",""]]
                }
              ]
            }
            """;

        var report = await tools.PreviewPlan("workspace", added.ItemId, planJson);
        using var parsed = JsonDocument.Parse(report);

        Assert.True(parsed.RootElement.GetProperty("isValid").GetBoolean(),
            $"Expected isValid=true. Got: {report}");
        Assert.Equal(1, parsed.RootElement.GetProperty("changes").GetArrayLength());
    }

    private static byte[] BuildDocWithTable()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var t = new Table(
                new TableRow(
                    new TableCell(new Paragraph(new Run(new Text("h1")))),
                    new TableCell(new Paragraph(new Run(new Text("h2"))))));
            main.Document = new Document(new Body(t));
            main.Document.Save();
        }
        return ms.ToArray();
    }

    private sealed class ToolsWorkspace : IDisposable
    {
        private readonly ServiceProvider _serviceProvider;

        public ToolsWorkspace()
        {
            Root = Path.Combine(Path.GetTempPath(), $"officeagent-anchor-{Guid.NewGuid():N}");
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
