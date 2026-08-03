using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.Abstractions;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.PowerPoint;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// A connection carries its own review policy. Tracked stays the global default - a
/// contract nobody asked to redline should still arrive redlined - but a connection over
/// generated documents, or one serving decks that refuse Tracked outright, can say so
/// once in configuration rather than relying on the agent to remember every call.
/// </summary>
public class ConnectionChangeModeTests
{
    [Fact]
    public void Tracked_remains_the_default_when_a_connection_says_nothing()
    {
        using var workspace = new ModeWorkspace(defaultChangeMode: null);

        Assert.Equal(ChangeMode.Tracked, workspace.Client.DefaultChangeModeFor("workspace"));
    }

    [Fact]
    public void A_connection_configured_for_direct_reports_direct()
    {
        using var workspace = new ModeWorkspace(ChangeMode.Direct);

        Assert.Equal(ChangeMode.Direct, workspace.Client.DefaultChangeModeFor("workspace"));
    }

    [Fact]
    public void An_unknown_connection_falls_back_to_the_global_default()
    {
        using var workspace = new ModeWorkspace(ChangeMode.Direct);

        // Reporting the wrong connection's policy would be worse than reporting none;
        // the caller's next step resolves the connection and fails properly.
        Assert.Equal(ChangeMode.Tracked, workspace.Client.DefaultChangeModeFor("no-such-connection"));
        Assert.Equal(ChangeMode.Tracked, workspace.Client.DefaultChangeModeFor(""));
    }

    [Fact]
    public async Task An_operation_without_a_mode_follows_the_connection()
    {
        using var tracked = new ModeWorkspace(ChangeMode.Tracked);
        using var direct = new ModeWorkspace(ChangeMode.Direct);

        Assert.True(await HasRevisions(tracked, ModeWorkspace.PlanWithoutMode));
        Assert.False(await HasRevisions(direct, ModeWorkspace.PlanWithoutMode));
    }

    [Fact]
    public async Task An_operation_that_names_its_mode_is_left_alone()
    {
        using var direct = new ModeWorkspace(ChangeMode.Direct);
        using var tracked = new ModeWorkspace(ChangeMode.Tracked);

        // The connection default fills a gap; it never overrides an explicit request.
        Assert.True(await HasRevisions(direct, ModeWorkspace.PlanWithMode("Tracked")));
        Assert.False(await HasRevisions(tracked, ModeWorkspace.PlanWithMode("Direct")));
    }

    [Fact]
    public async Task Preview_resolves_the_mode_the_same_way_apply_does()
    {
        using var tracked = new ModeWorkspace(ChangeMode.Tracked, deck: true);
        using var direct = new ModeWorkspace(ChangeMode.Direct, deck: true);

        // A deck refuses Tracked outright, which makes it the sharpest probe of which mode
        // preview actually resolved. A preview that validated while apply refused - or the
        // reverse - would send the agent round a loop it cannot get out of.
        Assert.False(await PreviewDeckEdit(tracked));
        Assert.True(await PreviewDeckEdit(direct));
    }

    [Fact]
    public async Task A_deck_connection_set_to_direct_stops_refusing_every_edit()
    {
        using var workspace = new ModeWorkspace(ChangeMode.Direct, deck: true);
        var tools = new OfficeAgentTools(workspace.Client);
        var deck = await workspace.Client.RegisterBytesAsync(
            "workspace", workspace.Root, PptxFactory.Deck(), "review.pptx");

        // PresentationML has no redline vocabulary, so under the global Tracked default
        // every deck edit that omits a mode is refused. This is the case the setting exists
        // for: the host states the policy once instead of the agent restating it per call.
        var report = await tools.ApplyPlan("workspace", deck.ItemId, DeckPlan);
        using var parsed = JsonDocument.Parse(report);

        Assert.True(parsed.RootElement.GetProperty("committed").GetBoolean(),
            parsed.RootElement.GetProperty("errors").ToString());
    }

    // ── helpers ───────────────────────────────────────────────────────────────

    private const string DeckPlan = """
        { "operations": [ { "op": "changeText",
            "target": { "paraId": "slide256/shape2/p0", "expect": "Quarterly Review" },
            "with": "Annual Review" } ] }
        """;

    private static async Task<bool> PreviewDeckEdit(ModeWorkspace workspace)
    {
        var tools = new OfficeAgentTools(workspace.Client);
        var deck = await workspace.Client.RegisterBytesAsync(
            "workspace", workspace.Root, PptxFactory.Deck(), "review.pptx");

        var report = await tools.PreviewPlan("workspace", deck.ItemId, DeckPlan);
        using var parsed = JsonDocument.Parse(report);
        return parsed.RootElement.GetProperty("isValid").GetBoolean();
    }

    private static async Task<bool> HasRevisions(ModeWorkspace workspace, string planJson)
    {
        var tools = new OfficeAgentTools(workspace.Client);
        var document = await workspace.Register();

        var report = await tools.ApplyPlan("workspace", document.ItemId, planJson);
        using var parsed = JsonDocument.Parse(report);
        Assert.True(parsed.RootElement.GetProperty("committed").GetBoolean(),
            parsed.RootElement.GetProperty("errors").ToString());

        using var content = await workspace.Client.OpenReadAsync(
            DocumentReference.ForFileSystem("workspace", document.ItemId));
        using var opened = WordprocessingDocument.Open(content.Stream, isEditable: false);
        return opened.MainDocumentPart!.Document.Descendants<InsertedRun>().Any();
    }

    private sealed class ModeWorkspace : IDisposable
    {
        public const string PlanWithoutMode = """
            { "operations": [ { "op": "changeText",
                "target": { "paraId": "w14:00000002", "expect": "Acme Corp" },
                "with": "Globex Inc." } ] }
            """;

        private readonly ServiceProvider _services;

        public ModeWorkspace(ChangeMode? defaultChangeMode, bool deck = false)
        {
            Root = Path.Combine(Path.GetTempPath(), $"officeagent-mode-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Root);

            var services = new ServiceCollection();
            services.AddWordFormat();
            services.AddPowerPointFormat();
            services.AddFileSystemDocumentProvider("workspace", Root, o =>
            {
                if (deck) o.AllowedExtensions = new[] { ".docx", ".pptx" };
                // Leaving it unset is the case that must still mean Tracked.
                if (defaultChangeMode is { } mode) o.DefaultChangeMode = mode;
            });
            services.AddOfficeAgent();
            _services = services.BuildServiceProvider();
            Client = _services.GetRequiredService<OfficeAgentClient>();
        }

        public string Root { get; }
        public OfficeAgentClient Client { get; }

        public static string PlanWithMode(string mode) => $$"""
            { "operations": [ { "op": "changeText",
                "target": { "paraId": "w14:00000002", "expect": "Acme Corp" },
                "with": "Globex Inc.", "mode": "{{mode}}" } ] }
            """;

        public Task<DocumentReference> Register() =>
            Client.RegisterBytesAsync("workspace", Root, DocxFactory.Contract(), "contract.docx");

        public void Dispose()
        {
            _services.Dispose();
            if (Directory.Exists(Root)) Directory.Delete(Root, recursive: true);
        }
    }
}
