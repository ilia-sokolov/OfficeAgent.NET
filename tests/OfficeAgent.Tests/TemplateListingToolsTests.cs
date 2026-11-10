using System.Security.Claims;
using System.Text.Json;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.Abstractions;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Mcp;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// The agent-facing side of listing: <c>list_documents</c> and <c>find_templates</c> on the Agent
/// Framework and MCP surfaces, the listing fields <c>list_connections</c> reports, and the Office
/// template folders the server discovers, over a machine supplied by the test.
/// </summary>
public class TemplateListingToolsTests
{
    // ── list_documents ────────────────────────────────────────────────────────

    [Fact]
    public async Task List_documents_returns_sources_that_open_document_accepts()
    {
        using var workspace = new DocumentListingTests.Workspace();
        workspace.Write("Contracts/contract.docx", DocxFactory.Contract());
        workspace.Write("Contracts/notes.docx", DocxFactory.Contract());
        var tools = Tools(workspace.Provider());

        using var listed = JsonDocument.Parse(await tools.ListDocuments("workspace", recursive: true, nameContains: "CONTRACT"));

        Assert.Equal("workspace", listed.RootElement.GetProperty("connectionId").GetString());
        var item = Assert.Single(listed.RootElement.GetProperty("items").EnumerateArray());
        Assert.Equal("Contracts/contract.docx", item.GetProperty("source").GetString());
        Assert.Equal("contract.docx", item.GetProperty("name").GetString());
        Assert.Equal("Contracts", item.GetProperty("folder").GetString());
        Assert.False(item.GetProperty("isTemplate").GetBoolean());
        Assert.True(item.GetProperty("length").GetInt64() > 0);
        Assert.Equal(JsonValueKind.String, item.GetProperty("lastModified").ValueKind);
        Assert.Equal(JsonValueKind.Null, listed.RootElement.GetProperty("continuation").ValueKind);

        using var opened = JsonDocument.Parse(await tools.OpenDocument("workspace", item.GetProperty("source").GetString()!));
        Assert.Equal("contract.docx", opened.RootElement.GetProperty("name").GetString());
    }

    [Fact]
    public async Task List_documents_pages_with_continuation_and_takes_a_comma_list_of_extensions()
    {
        using var workspace = new DocumentListingTests.Workspace();
        foreach (var name in new[] { "a.docx", "b.pptx", "c.docx", "d.docx" }) workspace.Write(name);
        var tools = Tools(workspace.Provider(o => o.AllowedExtensions = new[] { ".docx", ".pptx" }));

        using var first = JsonDocument.Parse(await tools.ListDocuments("workspace", extensions: "docx, .pptx", maximumResults: 3));
        var continuation = first.RootElement.GetProperty("continuation").GetString();
        using var second = JsonDocument.Parse(await tools.ListDocuments("workspace", extensions: "docx, .pptx", maximumResults: 3, continuation: continuation!));
        using var onlyDecks = JsonDocument.Parse(await tools.ListDocuments("workspace", extensions: ".pptx"));

        Assert.Equal(new[] { "a.docx", "b.pptx", "c.docx" }, Names(first, "items"));
        Assert.Equal(new[] { "d.docx" }, Names(second, "items"));
        Assert.Equal(new[] { "b.pptx" }, Names(onlyDecks, "items"));
    }

    [Fact]
    public async Task List_documents_on_a_connection_that_cannot_list_is_a_configuration_error()
    {
        var tools = Tools(new DocumentListingTests.NotListing("plain"));

        using var result = JsonDocument.Parse(await tools.ListDocuments("plain"));

        Assert.Equal(ToolErrorCodes.ConfigurationError, Code(result));
        Assert.Contains("cannot list", result.RootElement.GetProperty("errors")[0].GetProperty("message").GetString());
    }

    [Fact]
    public async Task List_documents_refuses_an_escape_and_a_denied_connection_without_disclosing_paths()
    {
        using var workspace = new DocumentListingTests.Workspace();
        var tools = Tools(workspace.Provider());

        var escaped = await tools.ListDocuments("workspace", folder: "../..");
        using var escape = JsonDocument.Parse(escaped);
        Assert.Equal(ToolErrorCodes.AccessDenied, Code(escape));
        Assert.DoesNotContain(workspace.Root.Replace("\\", "\\\\"), escaped);

        var denied = new OfficeAgentTools(Client(workspace.Provider()), new DenyAll());
        using var forbidden = JsonDocument.Parse(await denied.ListDocuments("workspace"));
        Assert.Equal(ToolErrorCodes.ConnectionForbidden, Code(forbidden));
    }

    [Fact]
    public async Task List_documents_lists_a_session_connection()
    {
        var session = new MemoryDocumentProvider("session");
        var held = session.Add("brief.docx", DocxFactory.Contract());
        var tools = Tools(session);

        using var listed = JsonDocument.Parse(await tools.ListDocuments("session"));

        var item = Assert.Single(listed.RootElement.GetProperty("items").EnumerateArray());
        Assert.Equal(held.ItemId, item.GetProperty("source").GetString());
    }

    // ── find_templates ────────────────────────────────────────────────────────

    [Fact]
    public async Task Find_templates_ranks_by_match_then_location_then_name()
    {
        using var shared = new DocumentListingTests.Workspace();
        using var personal = new DocumentListingTests.Workspace();
        using var project = new DocumentListingTests.Workspace();
        shared.Write("Engagement Letter.dotx");
        shared.Write("Board Letter.dotx");
        personal.Write("Letter.dotx");
        personal.Write("Letterhead.potx");
        project.Write("work.docx");
        project.Write("Templates/Letter.dotx");
        var templateExtensions = new[] { ".docx", ".dotx", ".pptx", ".potx" };
        var tools = Tools(
            shared.Provider(o => { o.ConnectionId = "shared"; o.TemplateLocation = TemplateLocationKind.Workgroup; o.AllowedExtensions = templateExtensions; }),
            personal.Provider(o => { o.ConnectionId = "mine"; o.TemplateLocation = TemplateLocationKind.Personal; o.AllowedExtensions = templateExtensions; }),
            project.Provider(o => { o.ConnectionId = "project"; o.TemplatesFolder = "Templates"; o.AllowedExtensions = templateExtensions; }),
            new MemoryDocumentProvider("session"));

        using var result = JsonDocument.Parse(await tools.FindTemplates(nameContains: "letter"));

        var found = result.RootElement.GetProperty("templates").EnumerateArray()
            .Select(t => $"{t.GetProperty("connectionId").GetString()}:{t.GetProperty("name").GetString()}:{t.GetProperty("location").GetString()}:{t.GetProperty("format").GetString()}")
            .ToArray();
        Assert.Equal(new[]
        {
            "mine:Letter.dotx:personal:word",          // the whole name
            "project:Letter.dotx:connection:word",
            "mine:Letterhead.potx:personal:powerpoint", // the start of the name
            "shared:Board Letter.dotx:workgroup:word",  // a word in the name, workgroup first
            "shared:Engagement Letter.dotx:workgroup:word",
        }, found);
        Assert.Equal(3, result.RootElement.GetProperty("connectionsSearched").GetInt32());
        Assert.Empty(result.RootElement.GetProperty("connectionsFailed").EnumerateArray());

        var letter = result.RootElement.GetProperty("templates")[1];
        Assert.Equal("Templates/Letter.dotx", letter.GetProperty("source").GetString());
    }

    [Fact]
    public async Task Find_templates_filters_by_format_and_limits_results()
    {
        using var personal = new DocumentListingTests.Workspace();
        personal.Write("Deck.potx");
        personal.Write("Memo.dotx");
        personal.Write("Budget.xltx");
        var tools = Tools(personal.Provider(o =>
        {
            o.TemplateLocation = TemplateLocationKind.Personal;
            o.AllowedExtensions = new[] { ".dotx", ".potx", ".xltx" };
        }));

        using var decks = JsonDocument.Parse(await tools.FindTemplates(format: "PowerPoint"));
        using var one = JsonDocument.Parse(await tools.FindTemplates(maximumResults: 1));
        using var unknown = JsonDocument.Parse(await tools.FindTemplates(format: "visio"));

        Assert.Equal(new[] { "Deck.potx" }, Names(decks, "templates"));
        Assert.Single(one.RootElement.GetProperty("templates").EnumerateArray());
        Assert.Equal(ToolErrorCodes.InvalidArgument, Code(unknown));
    }

    [Fact]
    public async Task One_failing_connection_does_not_fail_find_templates()
    {
        using var personal = new DocumentListingTests.Workspace();
        personal.Write("Memo.dotx");
        var tools = Tools(
            new FailingTemplates("library"),
            personal.Provider(o =>
            {
                o.TemplateLocation = TemplateLocationKind.Personal;
                o.AllowedExtensions = new[] { ".dotx" };
            }));

        using var result = JsonDocument.Parse(await tools.FindTemplates());

        Assert.Equal(new[] { "Memo.dotx" }, Names(result, "templates"));
        Assert.Equal(2, result.RootElement.GetProperty("connectionsSearched").GetInt32());
        var failure = Assert.Single(result.RootElement.GetProperty("connectionsFailed").EnumerateArray());
        Assert.Equal("library", failure.GetProperty("connectionId").GetString());
        Assert.Equal(ToolErrorCodes.AccessDenied, failure.GetProperty("code").GetString());
        Assert.DoesNotContain("tenant-secret", result.RootElement.ToString());
    }

    [Fact]
    public async Task Find_templates_skips_connections_the_caller_cannot_read()
    {
        using var personal = new DocumentListingTests.Workspace();
        personal.Write("Memo.dotx");
        var provider = personal.Provider(o =>
        {
            o.TemplateLocation = TemplateLocationKind.Personal;
            o.AllowedExtensions = new[] { ".dotx" };
        });
        var tools = new OfficeAgentTools(Client(provider), new DenyAll());

        using var result = JsonDocument.Parse(await tools.FindTemplates());

        Assert.Empty(result.RootElement.GetProperty("templates").EnumerateArray());
        Assert.Equal(0, result.RootElement.GetProperty("connectionsSearched").GetInt32());
    }

    [Fact]
    public void The_listing_tools_arrive_with_registration()
    {
        var tools = Tools(new MemoryDocumentProvider("session"));

        Assert.DoesNotContain(tools.AsAIFunctions(), f => f.Name is "list_documents" or "find_templates");
        var opted = tools.AsAIFunctions(new OfficeAgentToolsOptions { AllowRegistration = true });
        var listDocuments = Assert.Single(opted, f => f.Name == "list_documents");
        Assert.Contains("open_document", listDocuments.Description);
        Assert.Contains("create_document's template source", listDocuments.Description);
        Assert.Single(opted, f => f.Name == "find_templates");
        Assert.DoesNotContain(tools.AsAIFunctions(new OfficeAgentToolsOptions { AllowRegistration = true, AllowConnectionAddressing = false }),
            f => f.Name is "list_documents" or "find_templates");
    }

    // ── MCP ───────────────────────────────────────────────────────────────────

    [Fact]
    public async Task The_mcp_tools_are_read_only_and_list_a_configured_folder()
    {
        using var workspace = new DocumentListingTests.Workspace();
        workspace.Write("contract.docx", DocxFactory.Contract());
        workspace.Write("Templates/Letter.dotx", DocxFactory.Contract());
        var options = new OfficeAgentMcpOptions
        {
            FileSystemConnections =
            {
                new FileSystemConnectionOptions
                {
                    ConnectionId = "documents",
                    RootPath = workspace.Root,
                    AllowedExtensions = new List<string> { ".docx", ".dotx" },
                    TemplatesFolder = "Templates"
                }
            }
        };
        var services = new ServiceCollection();
        services.AddOfficeAgentMcp(options);
        using var provider = services.BuildServiceProvider();
        var tools = provider.GetRequiredService<OfficeAgentTools>();
        var mcpTools = OfficeAgentMcpServer.CreateTools(tools, options).ToDictionary(t => t.ProtocolTool.Name);

        foreach (var name in new[] { "list_documents", "find_templates" })
        {
            Assert.True(mcpTools[name].ProtocolTool.Annotations!.ReadOnlyHint);
            Assert.False(mcpTools[name].ProtocolTool.Annotations!.DestructiveHint);
        }

        using var listed = JsonDocument.Parse(await tools.ListDocuments("documents", recursive: true));
        Assert.Equal(new[] { "contract.docx", "Letter.dotx" }, Names(listed, "items"));
        using var found = JsonDocument.Parse(await tools.FindTemplates("word"));
        Assert.Equal("Templates/Letter.dotx", found.RootElement.GetProperty("templates")[0].GetProperty("source").GetString());
        Assert.Equal("connection", found.RootElement.GetProperty("templates")[0].GetProperty("location").GetString());
    }

    [Fact]
    public async Task List_connections_reports_listing_templates_and_read_only()
    {
        using var workspace = new DocumentListingTests.Workspace();
        using var templates = new DocumentListingTests.Workspace();
        var options = new OfficeAgentMcpOptions
        {
            AllowCreation = true,
            EphemeralConnectionId = "session",
            FileSystemConnections =
            {
                new FileSystemConnectionOptions { ConnectionId = "documents", RootPath = workspace.Root, TemplatesFolder = "Templates" },
                new FileSystemConnectionOptions
                {
                    ConnectionId = "office-workgroup-templates", RootPath = templates.Root, ReadOnly = true,
                    TemplateLocation = "workgroup", AllowedExtensions = new List<string> { ".docx", ".dotx" }
                }
            },
            SharePointConnections =
            {
                new SharePointConnectionOptions
                {
                    ConnectionId = "library", TenantId = "t", ClientId = "c", ClientSecret = "s",
                    TemplatesFolder = "b!drive/01FOLDER", TemplateLocation = "Organization", ReadOnly = true
                }
            }
        };
        var services = new ServiceCollection();
        services.AddSingleton(new HttpClient());
        services.AddOfficeAgentMcp(options);
        using var provider = services.BuildServiceProvider();

        using var live = JsonDocument.Parse(await OfficeAgentMcpServer.ConnectionsPayloadAsync(provider.GetRequiredService<OfficeAgentTools>(), options));
        using var configured = JsonDocument.Parse(OfficeAgentMcpServer.ConnectionsPayload(options));

        foreach (var payload in new[] { live, configured })
        {
            var entries = payload.RootElement.EnumerateArray().ToDictionary(e => e.GetProperty("connectionId").GetString()!);
            Assert.True(entries["session"].GetProperty("listable").GetBoolean());
            Assert.False(entries["session"].GetProperty("readOnly").GetBoolean());
            Assert.Equal(JsonValueKind.Null, entries["session"].GetProperty("templateLocation").ValueKind);

            Assert.True(entries["documents"].GetProperty("listable").GetBoolean());
            Assert.Equal("Templates", entries["documents"].GetProperty("templatesFolder").GetString());
            Assert.Equal("connection", entries["documents"].GetProperty("templateLocation").GetString());
            Assert.False(entries["documents"].GetProperty("readOnly").GetBoolean());
            Assert.True(entries["documents"].GetProperty("canCreateDocuments").GetBoolean());

            var workgroup = entries["office-workgroup-templates"];
            Assert.Equal("workgroup", workgroup.GetProperty("templateLocation").GetString());
            Assert.True(workgroup.GetProperty("readOnly").GetBoolean());
            // A read-only folder is never offered as somewhere to create a document.
            Assert.False(workgroup.GetProperty("canCreateDocuments").GetBoolean());

            Assert.True(entries["library"].GetProperty("listable").GetBoolean());
            Assert.Equal("organization", entries["library"].GetProperty("templateLocation").GetString());
            Assert.Equal("b!drive/01FOLDER", entries["library"].GetProperty("templatesFolder").GetString());
            Assert.True(entries["library"].GetProperty("readOnly").GetBoolean());
            Assert.False(entries["library"].GetProperty("canCreateDocuments").GetBoolean());
        }
    }

    [Theory]
    [InlineData("Shared", null)]
    [InlineData("1", null)]
    [InlineData("", "../Templates")]
    [InlineData("", "/abs/Templates")]
    public void Bad_template_settings_stop_the_server_and_name_themselves(string location, string? folder)
    {
        using var workspace = new DocumentListingTests.Workspace();
        var options = new OfficeAgentMcpOptions
        {
            FileSystemConnections =
            {
                new FileSystemConnectionOptions
                {
                    ConnectionId = "documents", RootPath = workspace.Root,
                    TemplateLocation = location, TemplatesFolder = folder ?? string.Empty
                }
            }
        };

        var error = Assert.Throws<InvalidOperationException>(() => new ServiceCollection().AddOfficeAgentMcp(options));

        Assert.Contains("'documents'", error.Message);
        Assert.Contains(folder is null ? "TemplateLocation" : "OfficeAgent__FileSystemConnections__0__TemplatesFolder", error.Message);
    }

    [Fact]
    public void A_sharepoint_templates_folder_must_be_a_url_or_a_pair()
    {
        var options = new OfficeAgentMcpOptions
        {
            SharePointConnections = { new SharePointConnectionOptions { ConnectionId = "library", TemplatesFolder = "Templates" } }
        };

        var error = Assert.Throws<InvalidOperationException>(() => new ServiceCollection().AddOfficeAgentMcp(options));

        Assert.Contains("OfficeAgent__SharePointConnections__0__TemplatesFolder", error.Message);
    }

    [Fact]
    public void Template_keys_bind_from_indexed_environment_style_keys()
    {
        var configuration = new ConfigurationBuilder().AddInMemoryCollection(new Dictionary<string, string?>
        {
            ["OfficeAgent:FileSystemConnections:0:ConnectionId"] = "documents",
            ["OfficeAgent:FileSystemConnections:0:RootPath"] = "/docs",
            ["OfficeAgent:FileSystemConnections:0:TemplatesFolder"] = "Templates",
            ["OfficeAgent:FileSystemConnections:0:TemplateLocation"] = "Workgroup",
            ["OfficeAgent:FileSystemConnections:0:ReadOnly"] = "true",
            ["OfficeAgent:SharePointConnections:0:ConnectionId"] = "library",
            ["OfficeAgent:SharePointConnections:0:TemplatesFolder"] = "https://contoso.sharepoint.com/sites/brand/Templates",
            ["OfficeAgent:SharePointConnections:0:TemplateLocation"] = "Organization",
        }).Build();

        var options = OfficeAgentConfiguration.Bind(configuration, new FakeMachine(), _ => { });

        var files = Assert.Single(options.FileSystemConnections);
        Assert.Equal("Templates", files.TemplatesFolder);
        Assert.Equal("Workgroup", files.TemplateLocation);
        Assert.True(files.ReadOnly);
        var library = Assert.Single(options.SharePointConnections);
        Assert.Equal("Organization", library.TemplateLocation);
        Assert.False(options.OfficeTemplates.Enabled);
    }

    // ── Office template locations ─────────────────────────────────────────────

    [Fact]
    public void With_no_registry_value_the_documents_custom_office_templates_folder_is_used()
    {
        using var documents = new DocumentListingTests.Workspace();
        Directory.CreateDirectory(Path.Combine(documents.Root, "Custom Office Templates"));
        var machine = new FakeMachine { DocumentsFolder = documents.Root };
        var log = new List<string>();

        var connections = OfficeTemplateLocations.Discover(machine, log.Add);

        var personal = Assert.Single(connections);
        Assert.Equal(OfficeTemplateLocations.PersonalConnectionId, personal.ConnectionId);
        Assert.Equal(Path.Combine(documents.Root, "Custom Office Templates"), personal.RootPath);
        Assert.True(personal.ReadOnly);
        Assert.Equal("Personal", personal.TemplateLocation);
        Assert.Equal(new[] { ".docx", ".dotx", ".pptx", ".potx", ".xlsx", ".xltx" }, personal.AllowedExtensions);
        Assert.Empty(log);
    }

    [Fact]
    public void Distinct_personal_folders_become_numbered_connections_and_missing_ones_are_logged()
    {
        using var word = new DocumentListingTests.Workspace();
        using var powerPoint = new DocumentListingTests.Workspace();
        using var documents = new DocumentListingTests.Workspace();
        Directory.CreateDirectory(Path.Combine(documents.Root, "Custom Office Templates"));
        var missing = Path.Combine(word.Root, "gone");
        var machine = new FakeMachine
        {
            DocumentsFolder = documents.Root,
            Registry =
            {
                [(@"Software\Microsoft\Office\16.0\Word\Options", "PersonalTemplates")] = word.Root,
                [(@"Software\Microsoft\Office\16.0\PowerPoint\Options", "PersonalTemplates")] = powerPoint.Root,
                [(@"Software\Microsoft\Office\16.0\Excel\Options", "PersonalTemplates")] = missing,
            }
        };
        var log = new List<string>();

        var connections = OfficeTemplateLocations.Discover(machine, log.Add);

        // A registry value replaces the Documents fallback, as it does in Office.
        Assert.Equal(new[] { "office-personal-templates-1", "office-personal-templates-2" }, connections.Select(c => c.ConnectionId));
        Assert.Equal(new[] { word.Root, powerPoint.Root }, connections.Select(c => c.RootPath));
        Assert.Contains(log, line => line.Contains(missing) && line.Contains("does not exist"));
    }

    [Fact]
    public void The_same_personal_folder_named_by_every_application_is_one_connection()
    {
        using var folder = new DocumentListingTests.Workspace();
        var machine = new FakeMachine();
        foreach (var application in new[] { "Word", "PowerPoint", "Excel" })
            machine.Registry[($@"Software\Microsoft\Office\16.0\{application}\Options", "PersonalTemplates")] =
                application == "Excel" ? folder.Root + Path.DirectorySeparatorChar : folder.Root;

        var personal = Assert.Single(OfficeTemplateLocations.Discover(machine, _ => { }));

        Assert.Equal(OfficeTemplateLocations.PersonalConnectionId, personal.ConnectionId);
    }

    [Fact]
    public void The_workgroup_folder_prefers_the_policy_value()
    {
        using var policy = new DocumentListingTests.Workspace();
        using var user = new DocumentListingTests.Workspace();
        var machine = new FakeMachine
        {
            Registry =
            {
                [(@"Software\Policies\Microsoft\Office\16.0\Common\General", "SharedTemplates")] = policy.Root,
                [(@"Software\Microsoft\Office\16.0\Common\General", "SharedTemplates")] = user.Root,
            }
        };

        var workgroup = Assert.Single(OfficeTemplateLocations.Discover(machine, _ => { }));
        Assert.Equal(OfficeTemplateLocations.WorkgroupConnectionId, workgroup.ConnectionId);
        Assert.Equal(policy.Root, workgroup.RootPath);
        Assert.Equal("Workgroup", workgroup.TemplateLocation);

        machine.Registry.Remove((@"Software\Policies\Microsoft\Office\16.0\Common\General", "SharedTemplates"));
        Assert.Equal(user.Root, Assert.Single(OfficeTemplateLocations.Discover(machine, _ => { })).RootPath);
    }

    [Fact]
    public void Missing_folders_are_skipped_and_nothing_is_found_off_windows_and_macos()
    {
        var log = new List<string>();
        var missingDocuments = Path.Combine(Path.GetTempPath(), $"officeagent-none-{Guid.NewGuid():N}");

        Assert.Empty(OfficeTemplateLocations.Discover(new FakeMachine { DocumentsFolder = missingDocuments }, log.Add));
        Assert.Single(log);
        Assert.Empty(OfficeTemplateLocations.Discover(new FakeMachine { IsWindows = false, DocumentsFolder = missingDocuments }, log.Add));
    }

    [Fact]
    public void On_macos_the_group_container_templates_folder_is_used()
    {
        using var home = new DocumentListingTests.Workspace();
        var templates = Path.Combine(home.Root, "Library", "Group Containers", "UBF8T346G9.Office", "User Content", "Templates");
        Directory.CreateDirectory(templates);

        var personal = Assert.Single(OfficeTemplateLocations.Discover(
            new FakeMachine { IsWindows = false, IsMacOS = true, HomeFolder = home.Root }, _ => { }));

        Assert.Equal(templates, personal.RootPath);
    }

    [Fact]
    public async Task Enabled_office_templates_become_read_only_connections_find_templates_searches()
    {
        using var documents = new DocumentListingTests.Workspace();
        using var shared = new DocumentListingTests.Workspace();
        documents.Write("Custom Office Templates/Memo.dotx", DocxFactory.Contract());
        shared.Write("Brand Deck.potx");
        using var root = new DocumentListingTests.Workspace();
        var configuration = new ConfigurationBuilder().AddInMemoryCollection(new Dictionary<string, string?>
        {
            ["OfficeAgent:OfficeTemplates:Enabled"] = "true",
            ["OfficeAgent:FileSystemConnections:0:ConnectionId"] = "documents",
            ["OfficeAgent:FileSystemConnections:0:RootPath"] = root.Root,
        }).Build();
        var machine = new FakeMachine
        {
            DocumentsFolder = documents.Root,
            Registry = { [(@"Software\Microsoft\Office\16.0\Common\General", "SharedTemplates")] = shared.Root }
        };

        var options = OfficeAgentConfiguration.Bind(configuration, machine, _ => { });

        Assert.Equal(new[] { "documents", "office-personal-templates", "office-workgroup-templates" },
            options.FileSystemConnections.Select(c => c.ConnectionId));
        var services = new ServiceCollection();
        services.AddOfficeAgentMcp(options);
        using var provider = services.BuildServiceProvider();
        var tools = provider.GetRequiredService<OfficeAgentTools>();

        using var found = JsonDocument.Parse(await tools.FindTemplates());
        Assert.Equal(new[] { "Brand Deck.potx:workgroup", "Memo.dotx:personal" },
            found.RootElement.GetProperty("templates").EnumerateArray()
                .Select(t => $"{t.GetProperty("name").GetString()}:{t.GetProperty("location").GetString()}"));

        // The templates open, and nothing is ever written into the folders they came from.
        using var opened = JsonDocument.Parse(await tools.OpenDocument("office-personal-templates", "Memo.dotx"));
        Assert.Equal("Memo.dotx", opened.RootElement.GetProperty("name").GetString());
        using var edited = JsonDocument.Parse(await tools.EditDocument("office-personal-templates", "Memo.dotx",
            "[{\"op\":\"changeText\",\"target\":{\"find\":\"Client:\"},\"with\":\"Fabrikam\",\"mode\":\"Direct\"}]", saveMode: "NewVersion"));
        Assert.Equal(ToolErrorCodes.AccessDenied, Code(edited));
        Assert.False(Directory.Exists(Path.Combine(documents.Root, "Custom Office Templates", ".officeagent")));
        Assert.Single(Directory.GetFiles(Path.Combine(documents.Root, "Custom Office Templates")));
    }

    [Fact]
    public void A_configured_connection_keeps_its_id_over_a_discovered_folder()
    {
        using var documents = new DocumentListingTests.Workspace();
        Directory.CreateDirectory(Path.Combine(documents.Root, "Custom Office Templates"));
        var options = new OfficeAgentMcpOptions
        {
            FileSystemConnections = { new FileSystemConnectionOptions { ConnectionId = "office-personal-templates", RootPath = documents.Root } }
        };
        var log = new List<string>();

        OfficeTemplateLocations.AddTo(options, new FakeMachine { DocumentsFolder = documents.Root }, log.Add);

        Assert.Equal(documents.Root, Assert.Single(options.FileSystemConnections).RootPath);
        Assert.Contains(log, line => line.Contains("already uses"));
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private static OfficeAgentClient Client(params IDocumentProvider[] providers) =>
        new(new DocumentProviderRegistry(providers), new WordModule());

    private static OfficeAgentTools Tools(params IDocumentProvider[] providers) => new(Client(providers));

    private static string[] Names(JsonDocument document, string array) =>
        document.RootElement.GetProperty(array).EnumerateArray().Select(e => e.GetProperty("name").GetString()!).ToArray();

    private static string? Code(JsonDocument document) =>
        document.RootElement.GetProperty("errors")[0].GetProperty("code").GetString();

    private sealed class DenyAll : IConnectionAccessPolicy
    {
        public ValueTask<bool> IsAllowedAsync(
            ClaimsPrincipal principal, string connectionId, ConnectionCapability capability, CancellationToken cancellationToken = default) =>
            new(false);
    }

    /// <summary>A template library whose listing always fails, with upstream detail in the message.</summary>
    private sealed class FailingTemplates : IDocumentListingProvider, IConnectionTemplateSource
    {
        public FailingTemplates(string connectionId) => ConnectionId = connectionId;
        public string Provider => "custom";
        public string ConnectionId { get; }
        public string? TemplatesFolder => "https://tenant-secret.example/Templates";
        public TemplateLocationKind TemplateLocation => TemplateLocationKind.Organization;
        public Task<DocumentListPage> ListAsync(DocumentListQuery query, CancellationToken cancellationToken = default) =>
            throw new DocumentProviderException(ProviderErrorCode.AccessDenied,
                "Graph refused https://tenant-secret.example/Templates", Provider, ConnectionId);
        public Task<DocumentReference> RegisterAsync(string source, CancellationToken cancellationToken = default) => throw new NotSupportedException();
        public Task<DocumentContent> OpenReadAsync(DocumentReference reference, CancellationToken cancellationToken = default) => throw new NotSupportedException();
        public Task<DocumentReference> SaveAsync(DocumentReference source, Stream content, SaveDocumentOptions options, CancellationToken cancellationToken = default) => throw new NotSupportedException();
        public Task RemoveAsync(DocumentReference reference, CancellationToken cancellationToken = default) => throw new NotSupportedException();
    }

    /// <summary>A machine with the registry and folders a test describes.</summary>
    private sealed class FakeMachine : IOfficeTemplateEnvironment
    {
        public bool IsWindows { get; init; } = true;
        public bool IsMacOS { get; init; }
        public string? DocumentsFolder { get; init; }
        public string? HomeFolder { get; init; }
        public Dictionary<(string SubKey, string Value), string> Registry { get; } = new();

        public string? ReadUserRegistryString(string subKey, string valueName) =>
            Registry.TryGetValue((subKey, valueName), out var value) ? value : null;

        public bool DirectoryExists(string path) => Directory.Exists(path);
    }
}
