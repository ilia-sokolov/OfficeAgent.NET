using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.Abstractions;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// Editing a document changes that document. The alternative - leaving the source untouched
/// beside a <c>contract.v2.docx</c> the caller then has to find and reconcile - is available
/// but opt-in. These tests pin which way round that is, at the contract, the client, and
/// the tool surface, because getting it wrong is silent: the caller reads the source back,
/// sees the old content, and concludes the edit did not apply.
/// </summary>
public class SaveModeDefaultTests
{
    [Fact]
    public void The_declared_default_is_replace()
    {
        Assert.Equal(SaveMode.Replace, new SaveDocumentOptions().Mode);
    }

    [Fact]
    public async Task A_commit_without_options_overwrites_the_source()
    {
        using var workspace = new ToolsWorkspace();
        var source = await workspace.Client.RegisterBytesAsync(
            "workspace", workspace.Root, DocxFactory.Contract(), "contract.docx");

        var result = await workspace.Client.CommitAsync(source, TitlePlan("Updated in place"));

        Assert.True(result.Committed);
        Assert.Equal(source.ItemId, result.Document!.ItemId);
        Assert.Equal("contract.docx", result.Document.Name);
        // Reading the id the caller already held must show the edit.
        Assert.Equal("Updated in place", await TitleOf(workspace, source.ItemId));
        Assert.Single(Directory.GetFiles(workspace.Root, "*.docx", SearchOption.AllDirectories));
    }

    [Fact]
    public async Task Asking_for_a_new_version_still_preserves_the_source()
    {
        using var workspace = new ToolsWorkspace();
        var source = await workspace.Client.RegisterBytesAsync(
            "workspace", workspace.Root, DocxFactory.Contract(), "contract.docx");

        var result = await workspace.Client.CommitAsync(
            source, TitlePlan("Kept aside"), new SaveDocumentOptions { Mode = SaveMode.NewVersion });

        Assert.True(result.Committed);
        Assert.NotEqual(source.ItemId, result.Document!.ItemId);
        Assert.Equal("contract.v2.docx", result.Document.Name);
        Assert.Equal("Kept aside", await TitleOf(workspace, result.Document.ItemId));
        Assert.NotEqual("Kept aside", await TitleOf(workspace, source.ItemId));
    }

    [Fact]
    public async Task The_tool_defaults_to_replace_and_still_honours_an_explicit_mode()
    {
        using var workspace = new ToolsWorkspace();
        var tools = new OfficeAgentTools(workspace.Client);
        var source = await workspace.Client.RegisterBytesAsync(
            "workspace", workspace.Root, DocxFactory.Contract(), "contract.docx");

        var byDefault = await ApplyTitle(tools, source.ItemId, "In place", saveMode: null);
        Assert.Equal(source.ItemId, byDefault.GetProperty("outputDocumentId").GetString());

        var explicitly = await ApplyTitle(tools, source.ItemId, "Aside", saveMode: "NewVersion");
        Assert.NotEqual(source.ItemId, explicitly.GetProperty("outputDocumentId").GetString());
        Assert.Equal("contract.v2.docx", explicitly.GetProperty("outputName").GetString());
    }

    [Fact]
    public async Task A_misspelled_save_mode_is_refused_rather_than_defaulted()
    {
        using var workspace = new ToolsWorkspace();
        var tools = new OfficeAgentTools(workspace.Client);
        var source = await workspace.Client.RegisterBytesAsync(
            "workspace", workspace.Root, DocxFactory.Contract(), "contract.docx");

        // Now that the default writes over the source, falling back to it on a typo would
        // destroy the document the caller was asking to keep.
        var report = await ApplyTitle(tools, source.ItemId, "Typo", saveMode: "NewVerison");

        var error = report.GetProperty("errors")[0];
        Assert.Equal("invalid-argument", error.GetProperty("Code").GetString());
        Assert.Contains("NewVersion", error.GetProperty("Message").GetString());
        Assert.False(report.GetProperty("committed").GetBoolean());
        // Nothing was written under either name.
        Assert.NotEqual("Typo", await TitleOf(workspace, source.ItemId));
        Assert.Single(Directory.GetFiles(workspace.Root, "*.docx", SearchOption.AllDirectories));
    }

    // ── helpers ───────────────────────────────────────────────────────────────

    private static DocumentPlan TitlePlan(string title) => new()
    {
        Operations = new PlanOperation[]
        {
            new SetPropertyOp
            {
                Target = new NodeAnchor { Kind = "docProperty", Path = "core/title" },
                Value = title
            }
        }
    };

    private static async Task<JsonElement> ApplyTitle(
        OfficeAgentTools tools, string documentId, string title, string? saveMode)
    {
        var planJson = JsonSerializer.Serialize(new
        {
            operations = new[]
            {
                new
                {
                    op = "setProperty",
                    target = new { kind = "docProperty", path = "core/title" },
                    value = title
                }
            }
        });

        var report = saveMode is null
            ? await tools.ApplyPlan("workspace", documentId, planJson)
            : await tools.ApplyPlan("workspace", documentId, planJson, saveMode);

        return JsonDocument.Parse(report).RootElement.Clone();
    }

    private static async Task<string?> TitleOf(ToolsWorkspace workspace, string documentId)
    {
        using var content = await workspace.Client.OpenReadAsync(
            DocumentReference.ForFileSystem("workspace", documentId));
        using var document = WordprocessingDocument.Open(content.Stream, isEditable: false);
        return document.PackageProperties.Title;
    }

    private sealed class ToolsWorkspace : IDisposable
    {
        private readonly ServiceProvider _serviceProvider;

        public ToolsWorkspace()
        {
            Root = Path.Combine(Path.GetTempPath(), $"officeagent-savemode-{Guid.NewGuid():N}");
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
