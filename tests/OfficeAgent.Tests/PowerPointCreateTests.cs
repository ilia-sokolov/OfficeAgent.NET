using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Mcp;
using OfficeAgent.PowerPoint;
using OfficeAgent.Word;
using DocFormat = OfficeAgent.Abstractions.DocumentFormat;

namespace OfficeAgent.Tests;

/// <summary>
/// Creating decks: the PowerPoint module mints a blank .pptx, the extension of the
/// requested name picks the format, and the MCP capability report follows the modules
/// actually registered rather than a hard-coded format.
/// </summary>
public class PowerPointCreateTests
{
    [Fact]
    public void Blank_deck_is_schema_valid_and_carries_one_anchor()
    {
        var module = new PowerPointModule();
        Assert.Equal(".pptx", module.Extension);

        var bytes = module.CreateBlank();

        using var stream = new MemoryStream(bytes);
        using var document = PresentationDocument.Open(stream, isEditable: false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(document).ToList();
        Assert.True(errors.Count == 0,
            string.Join("; ", errors.Select(e => $"{e.Path?.XPath}: {e.Description}")));

        // A deck needs a master, a layout, and a theme to open at all.
        var presentation = document.PresentationPart!;
        var master = Assert.Single(presentation.SlideMasterParts);
        Assert.NotNull(master.ThemePart);
        Assert.Single(presentation.SlideParts);

        // The layouts insertSlide names have to exist in the deck it creates, or a
        // generated slide would silently fall back to whichever layout happened to be
        // first and land with the wrong geometry.
        var types = master.SlideLayoutParts.Select(p => p.SlideLayout?.Type?.Value).ToList();
        Assert.Contains(SlideLayoutValues.Title, types);
        Assert.Contains(SlideLayoutValues.Object, types);
        Assert.Contains(SlideLayoutValues.SectionHeader, types);
        Assert.Contains(SlideLayoutValues.TitleOnly, types);
        Assert.Contains(SlideLayoutValues.Blank, types);
        // Every layout must point back at its master, or PowerPoint offers to repair.
        Assert.All(master.SlideLayoutParts, p => Assert.NotNull(p.SlideMasterPart));
    }

    [Fact]
    public void A_new_deck_exposes_the_anchor_an_initial_plan_targets()
    {
        var client = new OfficeAgentClient(new PowerPointModule());

        var inspection = client.Inspect(new PowerPointModule().CreateBlank());

        var paragraph = Assert.Single(inspection.Paragraphs);
        Assert.Equal("slide256/shape2/p0", paragraph.ParaId);
        Assert.Equal(string.Empty, paragraph.Text);
    }

    [Fact]
    public async Task Create_document_picks_the_format_from_the_name()
    {
        using var workspace = new DeckWorkspace();

        var deck = await workspace.Client.CreateAsync("workspace", "review.pptx");
        var document = await workspace.Client.CreateAsync("workspace", "notes.docx");

        Assert.True(deck.Committed);
        Assert.True(document.Committed);

        // One connection, two formats, chosen by extension alone.
        Assert.Equal(DocFormat.PowerPoint,
            (await workspace.Client.InspectAsync("workspace", deck.Document!.ItemId)).Format);
        Assert.Equal(DocFormat.Word,
            (await workspace.Client.InspectAsync("workspace", document.Document!.ItemId)).Format);
    }

    [Fact]
    public async Task A_new_deck_accepts_an_initial_plan()
    {
        using var workspace = new DeckWorkspace();

        var created = await workspace.Client.CreateAsync("workspace", "review.pptx", new DocumentPlan
        {
            Format = DocFormat.PowerPoint,
            Operations = new PlanOperation[]
            {
                new InsertTableOp
                {
                    Target = new NodeAnchor { Kind = "slide", Path = "slide#256" },
                    Table = new TableData
                    {
                        Headers = new[] { "Region", "Q1" },
                        Rows = new[] { new[] { "EMEA", "41850" } }
                    }
                }
            }
        });

        Assert.True(created.Committed,
            string.Join("; ", created.Report.Errors.Select(e => $"{e.Code}: {e.Message}")));

        var inspection = await workspace.Client.InspectAsync("workspace", created.Document!.ItemId);
        Assert.Single(inspection.Nodes, n => n.Kind == "table");
    }

    [Fact]
    public void Mcp_reports_creation_against_the_formats_it_actually_registers()
    {
        using var root = new TemporaryDeckRoot();

        // A connection that allows only .pptx is creatable now that a PowerPoint module
        // ships; before it existed the same configuration could create nothing.
        var options = new OfficeAgentMcpOptions
        {
            AllowCreation = true,
            FileSystemConnections =
            {
                new FileSystemConnectionOptions
                {
                    ConnectionId = "decks",
                    RootPath = root.Path,
                    AllowedExtensions = new List<string> { ".pptx" }
                },
                new FileSystemConnectionOptions
                {
                    ConnectionId = "sheets",
                    RootPath = root.Path,
                    AllowedExtensions = new List<string> { ".xlsx" }
                }
            }
        };

        var payload = OfficeAgentMcpServer.ConnectionsPayload(options);

        Assert.Contains("\"connectionId\":\"decks\",\"provider\":\"filesystem\",\"canCreateDocuments\":true", payload);
        // No module mints .xlsx, so that connection must still report false.
        Assert.Contains("\"connectionId\":\"sheets\",\"provider\":\"filesystem\",\"canCreateDocuments\":false", payload);
    }

    private sealed class DeckWorkspace : IDisposable
    {
        private readonly ServiceProvider _services;

        public DeckWorkspace()
        {
            Root = Path.Combine(Path.GetTempPath(), $"officeagent-deck-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Root);

            var services = new ServiceCollection();
            services.AddWordFormat();
            services.AddPowerPointFormat();
            services.AddFileSystemDocumentProvider("workspace", Root, o =>
                o.AllowedExtensions = new[] { ".docx", ".pptx" });
            services.AddOfficeAgent();
            _services = services.BuildServiceProvider();
            Client = _services.GetRequiredService<OfficeAgentClient>();
        }

        public string Root { get; }
        public OfficeAgentClient Client { get; }

        public void Dispose()
        {
            _services.Dispose();
            if (Directory.Exists(Root)) Directory.Delete(Root, recursive: true);
        }
    }

    private sealed class TemporaryDeckRoot : IDisposable
    {
        public TemporaryDeckRoot()
        {
            Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"officeagent-mcpdeck-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Path);
        }

        public string Path { get; }

        public void Dispose()
        {
            if (Directory.Exists(Path)) Directory.Delete(Path, recursive: true);
        }
    }
}
