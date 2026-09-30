using System.Security.Cryptography;
using System.Text.Json;
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.Abstractions;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Excel;

namespace OfficeAgent.Tests;

/// <summary>
/// An agent told only that "the document provider rejected an argument" cannot fix its
/// call. The provider's own message stays hidden, because it can carry paths or upstream
/// text, so the tool names the argument itself: it refuses a rename in Replace mode before
/// the provider sees it, and otherwise lists the arguments of the call that reached storage.
/// </summary>
public class ArgumentRefusalTests
{
    [Theory]
    [InlineData("apply_plan")]
    [InlineData("edit_document")]
    public async Task A_name_with_replace_is_refused_naming_new_name_and_nothing_is_written(string tool)
    {
        using var workspace = new Workspace();
        var source = workspace.Stage();
        var documentId = (await workspace.Client.RegisterAsync("registers", source)).ItemId;
        var before = workspace.FileHash();

        var json = tool == "apply_plan"
            ? await workspace.Tools.ApplyPlan("registers", documentId, AppendRow, saveMode: "Replace", newName: "register.xlsx")
            : await workspace.Tools.EditDocument("registers", source, AppendRow, saveMode: "Replace", newName: "register.xlsx");
        var response = JsonDocument.Parse(json).RootElement;

        var error = Assert.Single(response.GetProperty("errors").EnumerateArray());
        Assert.Equal(ToolErrorCodes.InvalidArgument, error.GetProperty("code").GetString());
        var message = error.GetProperty("message").GetString()!;
        Assert.Contains("newName", message);
        Assert.Contains("NewVersion", message);
        Assert.Equal("notWritten", response.GetProperty("writeOutcome").GetString());
        Assert.Equal(before, workspace.FileHash());
    }

    [Fact]
    public async Task A_blank_name_with_replace_means_no_name()
    {
        using var workspace = new Workspace();
        var documentId = (await workspace.Client.RegisterAsync("registers", workspace.Stage())).ItemId;

        var response = JsonDocument.Parse(
            await workspace.Tools.ApplyPlan("registers", documentId, AppendRow, saveMode: "Replace", newName: " ")).RootElement;

        Assert.True(response.GetProperty("committed").GetBoolean(), response.GetRawText());
        Assert.Equal(documentId, response.GetProperty("outputDocumentId").GetString());
    }

    [Fact]
    public async Task A_name_still_applies_to_a_new_version()
    {
        using var workspace = new Workspace();
        var documentId = (await workspace.Client.RegisterAsync("registers", workspace.Stage())).ItemId;

        var response = JsonDocument.Parse(
            await workspace.Tools.ApplyPlan("registers", documentId, AppendRow, saveMode: "NewVersion", newName: "register-q4.xlsx")).RootElement;

        Assert.True(response.GetProperty("committed").GetBoolean(), response.GetRawText());
        Assert.Equal("register-q4.xlsx", response.GetProperty("outputName").GetString());
    }

    [Fact]
    public async Task A_provider_refusal_names_the_arguments_that_reached_storage()
    {
        using var workspace = new Workspace();

        var response = JsonDocument.Parse(await workspace.Tools.RegisterDocument("registers", "")).RootElement;

        var error = Assert.Single(response.GetProperty("errors").EnumerateArray());
        Assert.Equal(ToolErrorCodes.InvalidArgument, error.GetProperty("code").GetString());
        var message = error.GetProperty("message").GetString()!;
        Assert.StartsWith("The document provider rejected an argument.", message);
        Assert.Contains("connectionId or source", message);
        // The provider's own wording is still withheld.
        Assert.DoesNotContain("filesystem registration", message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task A_refused_output_name_is_attributed_to_new_name()
    {
        using var workspace = new Workspace();
        var documentId = (await workspace.Client.RegisterAsync("registers", workspace.Stage())).ItemId;

        var response = JsonDocument.Parse(
            await workspace.Tools.ApplyPlan("registers", documentId, AppendRow, saveMode: "NewVersion", newName: "../outside.xlsx")).RootElement;

        Assert.False(response.GetProperty("committed").GetBoolean());
        var error = Assert.Single(response.GetProperty("errors").EnumerateArray());
        Assert.Equal(ToolErrorCodes.InvalidArgument, error.GetProperty("code").GetString());
        var message = error.GetProperty("message").GetString()!;
        Assert.Contains("newName", message);
    }

    private static readonly string AppendRow = JsonSerializer.Serialize(new
    {
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

    private sealed class Workspace : IDisposable
    {
        private readonly ServiceProvider _services;
        private string? _path;

        public Workspace()
        {
            Root = Path.Combine(Path.GetTempPath(), $"officeagent-arguments-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Root);
            var services = new ServiceCollection();
            services.AddExcelFormat();
            services.AddFileSystemDocumentProvider("registers", Root, o => o.AllowedExtensions = new[] { ".xlsx" });
            services.AddOfficeAgent();
            _services = services.BuildServiceProvider();
            Client = _services.GetRequiredService<OfficeAgentClient>();
            Tools = new OfficeAgentTools(Client);
        }

        public string Root { get; }
        public OfficeAgentClient Client { get; }
        public OfficeAgentTools Tools { get; }

        public string Stage() =>
            _path = TestRegistrationExtensions.StageFixture(Root, XlsxFactory.WorkbookWithTable(), "register.xlsx");

        public string FileHash() => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(_path!)));

        public void Dispose()
        {
            _services.Dispose();
            if (Directory.Exists(Root)) Directory.Delete(Root, recursive: true);
        }
    }
}
