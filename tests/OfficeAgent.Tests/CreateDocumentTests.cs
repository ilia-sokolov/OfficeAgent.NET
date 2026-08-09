using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.Abstractions;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Word;
using DocFormat = OfficeAgent.Abstractions.DocumentFormat;

namespace OfficeAgent.Tests;

public class CreateDocumentTests
{
    private const string BlankAnchor = "auto-0000";

    [Fact]
    public void A_created_document_defines_the_styles_its_own_verbs_name()
    {
        var blank = new WordModule().CreateBlank();

        using var stream = new MemoryStream(blank);
        using var document = WordprocessingDocument.Open(stream, isEditable: false);
        var styles = document.MainDocumentPart!.StyleDefinitionsPart?.Styles;

        // styleId writes a w:pStyle *reference*. With no definition to resolve, Word renders
        // the paragraph as Normal - the plan commits, the reference is in the file, and the
        // document looks nothing like what was asked for.
        Assert.NotNull(styles);
        var ids = styles!.Elements<Style>().Select(s => s.StyleId?.Value).ToList();
        Assert.Contains("Normal", ids);
        Assert.Contains("Heading1", ids);
        Assert.Contains("Heading2", ids);
        Assert.Contains("TableGrid", ids);

        // A heading must carry its outline level, or Inspect builds no outline from it.
        var heading = styles.Elements<Style>().Single(s => s.StyleId?.Value == "Heading1");
        Assert.Equal(0, heading.StyleParagraphProperties?.OutlineLevel?.Val?.Value);
    }

    [Fact]
    public void A_heading_inserted_into_a_created_document_resolves_to_a_real_style()
    {
        var client = new OfficeAgentClient(new WordModule());

        using var applied = client.Commit(
            new StreamHandle(new MemoryStream(new WordModule().CreateBlank())),
            new DocumentPlan
            {
                Operations = new PlanOperation[]
                {
                    new InsertOp
                    {
                        Target = new TextSpanAnchor { ParaId = BlankAnchor, Expect = string.Empty },
                        Position = InsertPosition.Before,
                        Text = "Statement of Work",
                        StyleId = "Heading1"
                    }
                }
            });

        Assert.True(applied.Committed,
            string.Join("; ", applied.Report.Errors.Select(e => $"{e.Code}: {e.Message}")));

        // Inspect resolves the style and builds the outline from it - both of which fail
        // silently when the style is only a dangling reference.
        var inspection = client.Inspect(applied.ToBytes());
        Assert.Equal("Heading1", inspection.Paragraphs.First().StyleId);
        Assert.Equal("Statement of Work", Assert.Single(inspection.Outline).Text);
    }

    [Fact]
    public void Blank_word_document_is_minimal_and_schema_valid()
    {
        var module = new WordModule();
        Assert.Equal(".docx", module.Extension);
        var bytes = module.CreateBlank();

        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, isEditable: false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(document).ToList();

        Assert.True(errors.Count == 0,
            string.Join("; ", errors.Select(e => $"{e.Path?.XPath}: {e.Description}")));
        var body = document.MainDocumentPart!.Document!.Body!;
        Assert.Single(body.Elements<Paragraph>());
        Assert.Single(document.Parts);
    }

    [Fact]
    public async Task Create_writes_and_registers_a_document()
    {
        using var workspace = new CreateWorkspace();

        var result = await workspace.Client.CreateAsync("workspace", "report.docx");

        Assert.True(result.Committed);
        Assert.Equal("report.docx", result.Document!.Name);
        Assert.True(File.Exists(Path.Combine(workspace.Root, "report.docx")));

        var inspect = await workspace.Client.InspectAsync("workspace", result.Document.ItemId);
        Assert.Equal(DocFormat.Word, inspect.Format);
        var paragraph = Assert.Single(inspect.Paragraphs);
        Assert.Equal(BlankAnchor, paragraph.ParaId);
        Assert.Equal(string.Empty, paragraph.Text);
    }

    [Fact]
    public async Task Create_applies_the_initial_plan_before_writing()
    {
        using var workspace = new CreateWorkspace();

        var result = await workspace.Client.CreateAsync(
            "workspace", "report.docx", InsertPlan("Quarterly Report"));

        Assert.True(result.Committed);
        var inspect = await workspace.Client.InspectAsync("workspace", result.Document!.ItemId);
        Assert.Equal("Quarterly Report", inspect.Paragraphs[0].Text);
    }

    [Fact]
    public async Task Invalid_initial_plan_writes_nothing()
    {
        using var workspace = new CreateWorkspace();
        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new InsertOp
                {
                    Target = new TextSpanAnchor { ParaId = "w14:DEADBEEF", Expect = "" },
                    Text = "Never written"
                }
            }
        };

        var result = await workspace.Client.CreateAsync("workspace", "report.docx", plan);

        Assert.False(result.Committed);
        Assert.Null(result.Document);
        Assert.NotEmpty(result.Report.Errors);
        Assert.False(File.Exists(Path.Combine(workspace.Root, "report.docx")));
    }

    [Fact]
    public async Task Create_never_overwrites_an_existing_name()
    {
        using var workspace = new CreateWorkspace();
        await workspace.Client.CreateAsync("workspace", "report.docx");
        var original = File.ReadAllBytes(Path.Combine(workspace.Root, "report.docx"));

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            workspace.Client.CreateAsync("workspace", "report.docx"));

        Assert.Equal(ProviderErrorCode.AlreadyExists, error.Code);
        Assert.Equal(original, File.ReadAllBytes(Path.Combine(workspace.Root, "report.docx")));
    }

    [Theory]
    [InlineData("CON.docx")]
    [InlineData("nul.docx")]
    [InlineData("LPT1.docx")]
    public async Task Create_rejects_reserved_windows_device_names(string name)
    {
        using var workspace = new CreateWorkspace();

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            workspace.Client.CreateAsync("workspace", name));

        Assert.Equal(ProviderErrorCode.InvalidArgument, error.Code);
        Assert.Contains("reserved Windows device name", error.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Empty(Directory.GetFiles(workspace.Root, "*.docx"));
    }

    [Fact]
    public async Task Create_rejects_paths_and_unsupported_formats()
    {
        using var workspace = new CreateWorkspace();

        var path = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            workspace.Client.CreateAsync("workspace", "drafts/report.docx"));
        var format = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            workspace.Client.CreateAsync("workspace", "budget.xlsx"));

        Assert.Equal(ProviderErrorCode.InvalidArgument, path.Code);
        Assert.Equal(ProviderErrorCode.InvalidArgument, format.Code);
        Assert.Empty(Directory.GetFiles(workspace.Root, "*.docx"));
    }

    [Fact]
    public async Task Registration_failure_leaves_the_created_file_intact()
    {
        using var workspace = new CreateWorkspace();
        Directory.CreateDirectory(Path.Combine(workspace.Root, ".officeagent", "index.json"));

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            workspace.Client.CreateAsync("workspace", "report.docx"));

        Assert.Equal(ProviderErrorCode.IO, error.Code);
        Assert.Contains("unregistered", error.Message);
        Assert.True(File.Exists(Path.Combine(workspace.Root, "report.docx")));
    }

    [Fact]
    public async Task Provider_without_creation_capability_is_rejected()
    {
        var registry = new DocumentProviderRegistry(new IDocumentProvider[] { new ReadOnlyProvider() });
        var client = new OfficeAgentClient(registry, new WordModule());

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            client.CreateAsync("readonly", "report.docx"));

        Assert.Equal(ProviderErrorCode.ConfigurationError, error.Code);
    }

    [Fact]
    public async Task Custom_service_composition_selects_the_factory_by_extension()
    {
        using var workspace = new CreateWorkspace();
        var word = new WordModule();
        var provider = new FileSystemDocumentProvider(new FileSystemDocumentProviderOptions
        {
            ConnectionId = "custom",
            RootPath = workspace.Root
        });
        var client = new OfficeAgentClient(
            new OfficeAgentEngine(new IFormatModule[] { word }),
            new DocumentProviderRegistry(new IDocumentProvider[] { provider }),
            loggerFactory: null,
            blankDocumentFactories: new IBlankDocumentFactory[] { new OtherBlankFactory(), word });

        var result = await client.CreateAsync("custom", "report.docx");

        Assert.True(result.Committed);
        Assert.Equal("report.docx", result.Document!.Name);
    }

    [Fact]
    public void Create_tool_has_an_independent_opt_in()
    {
        using var workspace = new CreateWorkspace();
        var tools = new OfficeAgentTools(workspace.Client);

        Assert.DoesNotContain("create_document", tools.AsAIFunctions().Select(f => f.Name));
        Assert.DoesNotContain("create_document", tools.AsAIFunctions(
            new OfficeAgentToolsOptions { AllowRegistration = true }).Select(f => f.Name));
        Assert.Contains("create_document", tools.AsAIFunctions(
            new OfficeAgentToolsOptions { AllowCreation = true }).Select(f => f.Name));
        Assert.Contains("create_document", tools.AsAIFunctions(
            new OfficeAgentToolsOptions { AllowRegistration = true, AllowCreation = true }).Select(f => f.Name));
    }

    [Fact]
    public async Task Create_tool_returns_a_usable_document_id()
    {
        using var workspace = new CreateWorkspace();
        var tools = new OfficeAgentTools(workspace.Client);
        var plan = """
            { "operations": [ { "op": "insert",
                                "target": { "paraId": "auto-0000", "expect": "" },
                                "position": "Before",
                                "text": "Quarterly Report" } ] }
            """;

        using var created = JsonDocument.Parse(
            await tools.CreateDocument("workspace", "report.docx", plan));

        Assert.True(created.RootElement.GetProperty("committed").GetBoolean());
        var id = created.RootElement.GetProperty("outputDocumentId").GetString()!;
        using var inspected = JsonDocument.Parse(await tools.InspectDocument("workspace", id));
        Assert.Equal("Quarterly Report",
            inspected.RootElement.GetProperty("paragraphs")[0].GetProperty("Text").GetString());
    }

    [Fact]
    public async Task Create_tool_returns_structured_errors()
    {
        using var workspace = new CreateWorkspace();
        var tools = new OfficeAgentTools(workspace.Client);

        using var badPlan = JsonDocument.Parse(
            await tools.CreateDocument("workspace", "report.docx", "{ not json"));
        using var badName = JsonDocument.Parse(
            await tools.CreateDocument("workspace", "drafts/report.docx"));
        using var invalidCharacter = JsonDocument.Parse(
            await tools.CreateDocument("workspace", "bad<name.docx"));
        await tools.CreateDocument("workspace", "taken.docx");
        using var duplicate = JsonDocument.Parse(
            await tools.CreateDocument("workspace", "taken.docx"));

        Assert.Equal("invalid-json",
            badPlan.RootElement.GetProperty("errors")[0].GetProperty("Code").GetString());
        Assert.Equal("invalid-argument",
            badName.RootElement.GetProperty("errors")[0].GetProperty("Code").GetString());
        Assert.Equal("invalid-argument",
            invalidCharacter.RootElement.GetProperty("errors")[0].GetProperty("Code").GetString());
        Assert.Contains("invalid filename",
            invalidCharacter.RootElement.GetProperty("errors")[0].GetProperty("Message").GetString(),
            StringComparison.OrdinalIgnoreCase);
        Assert.Equal("already-exists",
            duplicate.RootElement.GetProperty("errors")[0].GetProperty("Code").GetString());
    }

    private static DocumentPlan InsertPlan(string text) => new()
    {
        Operations = new PlanOperation[]
        {
            new InsertOp
            {
                Target = new TextSpanAnchor { ParaId = BlankAnchor, Expect = string.Empty },
                Position = InsertPosition.Before,
                Text = text
            }
        }
    };

    private sealed class ReadOnlyProvider : IDocumentProvider
    {
        public string Provider => "readonly";
        public string ConnectionId => "readonly";
        public Task<DocumentReference> RegisterAsync(string source, CancellationToken cancellationToken = default) =>
            throw new NotSupportedException();
        public Task<DocumentContent> OpenReadAsync(DocumentReference reference, CancellationToken cancellationToken = default) =>
            throw new NotSupportedException();
        public Task<DocumentReference> SaveAsync(DocumentReference source, Stream content, SaveDocumentOptions options, CancellationToken cancellationToken = default) =>
            throw new NotSupportedException();
        public Task RemoveAsync(DocumentReference reference, CancellationToken cancellationToken = default) =>
            throw new NotSupportedException();
    }

    private sealed class OtherBlankFactory : IBlankDocumentFactory
    {
        public string Extension => ".other";
        public byte[] CreateBlank() => throw new InvalidOperationException("Wrong factory selected.");
    }

    private sealed class CreateWorkspace : IDisposable
    {
        private readonly ServiceProvider _services;

        public CreateWorkspace()
        {
            Root = Path.Combine(Path.GetTempPath(), $"officeagent-create-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Root);
            var services = new ServiceCollection();
            services.AddWordFormat();
            services.AddFileSystemDocumentProvider("workspace", Root);
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
}
