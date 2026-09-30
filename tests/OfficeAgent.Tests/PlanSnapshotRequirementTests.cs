using System.Security.Cryptography;
using System.Text.Json;
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.Abstractions;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Excel;
using OfficeAgent.Mcp;

namespace OfficeAgent.Tests;

/// <summary>
/// A plan bound to its snapshot cannot be applied twice: after the first save the same plan
/// is stale. That is what makes a retried call safe, and it holds only if the plan carries
/// the snapshot, which the engine does not add. With <see cref="OfficeAgentTools.RequirePlanSnapshot"/>
/// on, a plan without one is refused before anything is read or written, so the guarantee no
/// longer depends on the model remembering to copy it.
/// </summary>
public class PlanSnapshotRequirementTests
{
    [Fact]
    public async Task By_default_a_plan_without_a_snapshot_is_still_applied()
    {
        using var workspace = new Workspace(requireSnapshot: false);
        var documentId = await workspace.Register();

        var applied = Parse(await workspace.Tools.ApplyPlan("registers", documentId, AppendRow(snapshot: null)));

        Assert.True(applied.GetProperty("committed").GetBoolean(), applied.GetRawText());
    }

    [Fact]
    public async Task Preview_refuses_a_plan_without_a_snapshot()
    {
        using var workspace = new Workspace(requireSnapshot: true);
        var documentId = await workspace.Register();

        var preview = Parse(await workspace.Tools.PreviewPlan("registers", documentId, AppendRow(snapshot: null)));

        AssertSnapshotRequired(preview);
        Assert.Equal(documentId, preview.GetProperty("sourceDocumentId").GetString());
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public async Task Apply_refuses_a_plan_without_a_snapshot_and_writes_nothing(string? snapshot)
    {
        using var workspace = new Workspace(requireSnapshot: true);
        var documentId = await workspace.Register();
        var before = workspace.FileHash();

        var applied = Parse(await workspace.Tools.ApplyPlan("registers", documentId, AppendRow(snapshot)));

        AssertSnapshotRequired(applied);
        Assert.Equal(before, workspace.FileHash());
    }

    [Fact]
    public async Task A_bare_operations_array_cannot_carry_a_snapshot_and_is_refused()
    {
        using var workspace = new Workspace(requireSnapshot: true);
        var documentId = await workspace.Register();
        var operations = JsonDocument.Parse(AppendRow(snapshot: null)).RootElement.GetProperty("operations").GetRawText();

        var applied = Parse(await workspace.Tools.ApplyPlan("registers", documentId, operations));

        AssertSnapshotRequired(applied);
    }

    [Fact]
    public async Task A_retried_plan_cannot_append_the_row_twice()
    {
        using var workspace = new Workspace(requireSnapshot: true);
        var documentId = await workspace.Register();
        var snapshot = await workspace.Snapshot(documentId);
        var plan = AppendRow(snapshot);

        var first = Parse(await workspace.Tools.ApplyPlan("registers", documentId, plan));
        var retry = Parse(await workspace.Tools.ApplyPlan("registers", documentId, plan));

        Assert.True(first.GetProperty("committed").GetBoolean(), first.GetRawText());
        Assert.False(retry.GetProperty("committed").GetBoolean());
        Assert.Equal("notWritten", retry.GetProperty("writeOutcome").GetString());
        Assert.Contains(retry.GetProperty("errors").EnumerateArray(),
            error => error.GetProperty("code").GetString() == ValidationErrorCodes.StaleSnapshot);
        Assert.Equal(1, await workspace.RowsWith(documentId, "APAC"));
    }

    [Fact]
    public async Task Edit_document_refuses_before_registering_anything()
    {
        using var workspace = new Workspace(requireSnapshot: true);
        var source = workspace.Stage();

        var edited = Parse(await workspace.Tools.EditDocument("registers", source, AppendRow(snapshot: null)));

        AssertSnapshotRequired(edited);
        Assert.Equal(JsonValueKind.Null, edited.GetProperty("sourceDocumentId").ValueKind);
    }

    [Fact]
    public async Task Edit_document_reads_the_snapshot_whatever_its_casing()
    {
        using var workspace = new Workspace(requireSnapshot: true);
        var source = workspace.Stage();
        var inspected = await workspace.Client.RegisterAsync("registers", source);
        var snapshot = await workspace.Snapshot(inspected.ItemId);
        var plan = AppendRow(snapshot).Replace("\"snapshot\"", "\"Snapshot\"").Replace("\"eTag\"", "\"ETag\"");

        var edited = Parse(await workspace.Tools.EditDocument("registers", source, plan));

        Assert.True(edited.GetProperty("committed").GetBoolean(), edited.GetRawText());
    }

    [Fact]
    public void The_mcp_server_passes_the_setting_to_the_tools_and_tells_the_agent()
    {
        using var workspace = new Workspace(requireSnapshot: false);
        var options = new OfficeAgentMcpOptions
        {
            RequirePlanSnapshot = true,
            FileSystemConnections =
            {
                new FileSystemConnectionOptions { ConnectionId = "registers", RootPath = workspace.Root }
            }
        };
        using var services = new ServiceCollection().AddOfficeAgentMcp(options).BuildServiceProvider();

        Assert.True(services.GetRequiredService<OfficeAgentTools>().RequirePlanSnapshot);
        Assert.Contains(OfficeAgentTools.SnapshotRequiredPromptGuidance, OfficeAgentMcpServer.InstructionsFor(options));
        Assert.DoesNotContain(
            OfficeAgentTools.SnapshotRequiredPromptGuidance,
            OfficeAgentMcpServer.InstructionsFor(new OfficeAgentMcpOptions { FileSystemConnections = options.FileSystemConnections }));
    }

    private static void AssertSnapshotRequired(JsonElement response)
    {
        Assert.False(response.GetProperty("isValid").GetBoolean());
        Assert.False(response.GetProperty("committed").GetBoolean());
        Assert.Equal("notWritten", response.GetProperty("writeOutcome").GetString());
        var error = Assert.Single(response.GetProperty("errors").EnumerateArray());
        Assert.Equal(ToolErrorCodes.SnapshotRequired, error.GetProperty("code").GetString());
        Assert.Contains("snapshot.eTag", error.GetProperty("message").GetString());
    }

    private static string AppendRow(string? snapshot) => JsonSerializer.Serialize(new
    {
        snapshot = new { eTag = snapshot },
        operations = new object[]
        {
            new
            {
                op = "appendTableRows",
                target = new { kind = "spreadsheetTable", path = "table#7/Sales" },
                rows = new[] { new[] { "APAC", "15" } }
            }
        }
    });

    private static JsonElement Parse(string json) => JsonDocument.Parse(json).RootElement.Clone();

    private sealed class Workspace : IDisposable
    {
        private readonly ServiceProvider _services;
        private string? _path;

        public Workspace(bool requireSnapshot)
        {
            Root = Path.Combine(Path.GetTempPath(), $"officeagent-snapshot-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Root);
            var services = new ServiceCollection();
            services.AddExcelFormat();
            services.AddFileSystemDocumentProvider("registers", Root, o => o.AllowedExtensions = new[] { ".xlsx" });
            services.AddOfficeAgent();
            _services = services.BuildServiceProvider();
            Client = _services.GetRequiredService<OfficeAgentClient>();
            Tools = new OfficeAgentTools(Client) { RequirePlanSnapshot = requireSnapshot };
        }

        public string Root { get; }
        public OfficeAgentClient Client { get; }
        public OfficeAgentTools Tools { get; }

        public string Stage() =>
            _path = TestRegistrationExtensions.StageFixture(Root, XlsxFactory.WorkbookWithTable(), "register.xlsx");

        public async Task<string> Register() => (await Client.RegisterAsync("registers", Stage())).ItemId;

        public async Task<string> Snapshot(string documentId) =>
            (await Client.InspectAsync("registers", documentId)).Snapshot.ETag;

        public async Task<int> RowsWith(string documentId, string value) =>
            (await Client.InspectAsync("registers", documentId)).Cells.Count(cell => cell.DisplayValue == value);

        public string FileHash() => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(_path!)));

        public void Dispose()
        {
            _services.Dispose();
            if (Directory.Exists(Root)) Directory.Delete(Root, recursive: true);
        }
    }
}
