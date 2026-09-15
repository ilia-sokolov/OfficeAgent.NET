using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeAgent.Abstractions;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// The template preflight contract as an agent reaches it: through the adapter, in JSON.
/// </summary>
/// <remarks>
/// V09-09 built discovery, batch preflight and a hash-bound commit, and V09-10 added typed
/// media bindings. All of it landed on the direct .NET client, and until V09-14 neither the
/// Agent Framework nor the MCP surface exposed any of it. An agent could commit a batch and
/// could not discover a template's slots, could not preview a batch, and could not bind a
/// commit to a preview it had reviewed, which is the guarantee the preflight exists to give.
/// These tests hold the adapter to the same contract the direct client keeps.
/// </remarks>
public sealed class AdapterTemplateParityTests
{
    /// <summary>Discovery reports the template's slots and writes nothing.</summary>
    [Fact]
    public async Task Discover_template_reports_slots_and_writes_nothing()
    {
        using var workspace = new TemplateToolsWorkspace();
        var tools = new OfficeAgentTools(workspace.Client);
        var template = await workspace.RegisterAsync("quote.docx", QuoteTemplate());
        var before = Directory.GetFiles(workspace.Root).Length;

        using var discovery = JsonDocument.Parse(
            await tools.DiscoverTemplate("workspace", template.ItemId));

        var slots = discovery.RootElement.GetProperty("slots").EnumerateArray()
            .Select(slot => slot.GetProperty("name").GetString())
            .ToArray();
        Assert.Contains("CustomerName", slots);

        var rows = discovery.RootElement.GetProperty("repeatingRows").EnumerateArray().ToArray();
        Assert.NotEmpty(rows);

        Assert.False(string.IsNullOrWhiteSpace(
            discovery.RootElement.GetProperty("templateSha256").GetString()));
        Assert.Equal(before, Directory.GetFiles(workspace.Root).Length);
    }

    /// <summary>The preflight validates the whole batch, writes nothing, and issues a token.</summary>
    [Fact]
    public async Task Preview_template_batch_issues_a_token_and_writes_nothing()
    {
        using var workspace = new TemplateToolsWorkspace();
        var tools = new OfficeAgentTools(workspace.Client);
        var template = await workspace.RegisterAsync("quote.docx", QuoteTemplate());
        var before = Directory.GetFiles(workspace.Root).Length;

        using var preview = JsonDocument.Parse(
            await tools.PreviewTemplateBatch("workspace", template.ItemId, Batch("a.docx", "Fabrikam")));

        Assert.True(preview.RootElement.GetProperty("isValid").GetBoolean(), preview.RootElement.ToString());
        Assert.Single(preview.RootElement.GetProperty("items").EnumerateArray());

        var token = preview.RootElement.GetProperty("token");
        Assert.False(string.IsNullOrWhiteSpace(token.GetProperty("templateSha256").GetString()));
        Assert.False(string.IsNullOrWhiteSpace(token.GetProperty("batchSha256").GetString()));

        Assert.Equal(before, Directory.GetFiles(workspace.Root).Length);
    }

    /// <summary>
    /// The full loop an agent is meant to run: preview, review, commit bound to that review.
    /// </summary>
    [Fact]
    public async Task A_previewed_batch_commits_when_the_token_still_matches()
    {
        using var workspace = new TemplateToolsWorkspace();
        var tools = new OfficeAgentTools(workspace.Client);
        var template = await workspace.RegisterAsync("quote.docx", QuoteTemplate());
        var batch = Batch("a.docx", "Fabrikam");

        using var preview = JsonDocument.Parse(
            await tools.PreviewTemplateBatch("workspace", template.ItemId, batch));
        var token = preview.RootElement.GetProperty("token").GetRawText();

        using var result = JsonDocument.Parse(
            await tools.PopulateTemplateBatch("workspace", template.ItemId, batch, token));

        Assert.True(result.RootElement.GetProperty("committed").GetBoolean(), result.RootElement.ToString());
        Assert.Contains(Directory.GetFiles(workspace.Root), file => Path.GetFileName(file) == "a.docx");
    }

    /// <summary>
    /// The repair loop. A commit whose batch no longer matches the reviewed preview is
    /// refused through the adapter exactly as it is directly, writes nothing, and names the
    /// reason; previewing the new batch produces a token that then commits.
    /// </summary>
    [Fact]
    public async Task A_commit_whose_batch_changed_after_review_is_refused_then_repairable()
    {
        using var workspace = new TemplateToolsWorkspace();
        var tools = new OfficeAgentTools(workspace.Client);
        var template = await workspace.RegisterAsync("quote.docx", QuoteTemplate());

        using var preview = JsonDocument.Parse(
            await tools.PreviewTemplateBatch("workspace", template.ItemId, Batch("a.docx", "Fabrikam")));
        var staleToken = preview.RootElement.GetProperty("token").GetRawText();

        int before = Directory.GetFiles(workspace.Root).Length;
        using var refused = JsonDocument.Parse(await tools.PopulateTemplateBatch(
            "workspace", template.ItemId, Batch("a.docx", "Someone Else"), staleToken));

        Assert.False(refused.RootElement.GetProperty("committed").GetBoolean());
        Assert.Contains("stale-batch-preview", refused.RootElement.ToString(), StringComparison.Ordinal);
        Assert.Equal(before, Directory.GetFiles(workspace.Root).Length);

        // Repair: preview what is actually being committed, then commit against that token.
        using var repreview = JsonDocument.Parse(await tools.PreviewTemplateBatch(
            "workspace", template.ItemId, Batch("a.docx", "Someone Else")));
        using var committed = JsonDocument.Parse(await tools.PopulateTemplateBatch(
            "workspace", template.ItemId, Batch("a.docx", "Someone Else"),
            repreview.RootElement.GetProperty("token").GetRawText()));

        Assert.True(committed.RootElement.GetProperty("committed").GetBoolean(), committed.RootElement.ToString());
        Assert.Contains(Directory.GetFiles(workspace.Root), file => Path.GetFileName(file) == "a.docx");
    }

    /// <summary>
    /// Omitting the token keeps the previous behavior rather than becoming a new failure
    /// mode: the batch is still preflighted internally and still commits.
    /// </summary>
    [Fact]
    public async Task Committing_without_a_token_still_works()
    {
        using var workspace = new TemplateToolsWorkspace();
        var tools = new OfficeAgentTools(workspace.Client);
        var template = await workspace.RegisterAsync("quote.docx", QuoteTemplate());

        using var result = JsonDocument.Parse(
            await tools.PopulateTemplateBatch("workspace", template.ItemId, Batch("a.docx", "Fabrikam")));

        Assert.True(result.RootElement.GetProperty("committed").GetBoolean(), result.RootElement.ToString());
    }

    private static string Batch(string outputName, string customer) =>
        "{ \"items\": [ { \"outputName\": \"" + outputName + "\", \"binding\": { " +
        "\"values\": { \"CustomerName\": \"" + customer + "\" }, " +
        "\"missingValueBehavior\": \"Ignore\" } } ] }";

    private static SdtRun Control(string tag, string text) => new(
        new SdtProperties(new Tag { Val = tag }, new SdtId { Val = Math.Abs(tag.GetHashCode() % 10000) + 10 }),
        new SdtContentRun(new Run(new Text(text))));

    private static TableRow Row(params string[] values) => new(values.Select(value =>
        new TableCell(new Paragraph(new Run(new Text(value))))));

    private static byte[] QuoteTemplate()
    {
        using var ms = new MemoryStream();
        using (var document = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Quote for ")), Control("CustomerName", "CUSTOMER")),
                new Table(
                    new TableProperties(new TableStyle { Val = "TableGrid" }),
                    new TableGrid(
                        new GridColumn { Width = "3600" },
                        new GridColumn { Width = "1200" },
                        new GridColumn { Width = "1800" }),
                    Row("Description", "Quantity", "Price"),
                    Row("{{Description}}", "{{Quantity}}", "{{Price}}")),
                new Paragraph()));
            main.Document.Save();
        }

        return ms.ToArray();
    }

    private sealed class TemplateToolsWorkspace : IDisposable
    {
        public TemplateToolsWorkspace()
        {
            Root = Path.Combine(Path.GetTempPath(), $"officeagent-adapter-template-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Root);

            Client = new OfficeAgentClient(
                new DocumentProviderRegistry(new[]
                {
                    new FileSystemDocumentProvider(new FileSystemDocumentProviderOptions
                    {
                        ConnectionId = "workspace",
                        RootPath = Root,
                        DefaultChangeMode = ChangeMode.Direct
                    })
                }),
                new WordModule());
        }

        public string Root { get; }

        public OfficeAgentClient Client { get; }

        public async Task<DocumentReference> RegisterAsync(string name, byte[] content)
        {
            var path = Path.Combine(Root, name);
            await File.WriteAllBytesAsync(path, content);
            return await Client.RegisterAsync("workspace", path);
        }

        public void Dispose()
        {
            if (Directory.Exists(Root)) Directory.Delete(Root, recursive: true);
        }
    }
}
