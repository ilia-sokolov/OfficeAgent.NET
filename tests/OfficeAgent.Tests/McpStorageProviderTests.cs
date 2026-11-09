using System.Security.Cryptography;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.Abstractions;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Mcp;

namespace OfficeAgent.Tests;

/// <summary>
/// The iManage, Google Drive, Box, Azure Blob and Amazon S3 connections of the MCP server: each is
/// configured the way a host configures SharePoint, appears in discovery with what it accepts, stops
/// the server on bad configuration, and edits documents through the real provider against that
/// provider's fake service - with no secret in any tool result, description or error.
/// </summary>
public class McpStorageProviderTests
{
    private const string Edit =
        """[ { "op": "changeText", "target": { "find": "Effective date: 2020-01-01." }, "with": "Effective date: 2021-01-01." } ]""";

    private const string EditWithSpan =
        """[ { "op": "changeText", "target": { "find": "Effective date: 2020-01-01." }, "with": "Effective date: 2021-01-01.", "granularity": "Span" } ]""";

    // Distinctive values, so a leak anywhere is found by a plain substring search.
    private const string IManagePassword = "imanage-password-Zq81";
    private const string IManageClientSecret = "imanage-client-secret-Zq82";
    private const string BoxClientSecret = "box-client-secret-Zq83";
    private const string BlobAccountKey = "YmxvYi1hY2NvdW50LWtleS1acTg0";

    private static readonly string DriveKeyPem = NewPrivateKeyPem();

    private static readonly string DriveKeyJson = JsonSerializer.Serialize(new
    {
        type = "service_account",
        client_email = "officeagent@contoso.iam.gserviceaccount.com",
        private_key = DriveKeyPem,
        private_key_id = "drive-key-id-Zq85"
    });

    // ── One rig per provider ──────────────────────────────────────────────────

    /// <summary>A provider's fake service, configured into the MCP server the way a host would.</summary>
    private abstract class Rig : IDisposable
    {
        public abstract string ConnectionId { get; }
        public abstract string Provider { get; }

        /// <summary>Puts a fresh contract in the store and returns the source an agent registers it by.</summary>
        public abstract string Seed(string name);

        /// <summary>The stored bytes of the document seeded under <paramref name="name"/>.</summary>
        public abstract byte[] Saved(string name);

        /// <summary>Asserts the save reached the provider carrying its precondition.</summary>
        public abstract void AssertConditionalSave(string name);

        /// <summary>Makes the next save of <paramref name="name"/> conflict or be refused, and returns the error code expected.</summary>
        public abstract string Obstruct(string name);

        public abstract void Configure(OfficeAgentMcpOptions options, bool creation, bool minimal);

        public abstract StorageTransports Transports();

        /// <summary>The templates folder as the connection's TemplatesFolder setting names it.</summary>
        public abstract string TemplatesFolder { get; }

        /// <summary>Puts a template in <see cref="TemplatesFolder"/>.</summary>
        public abstract void SeedTemplate(string name);

        public virtual void Dispose()
        {
        }
    }

    private sealed class IManageRig : Rig
    {
        private readonly IManageProviderTests.FakeIManage _server = new();
        private readonly Dictionary<string, long> _numbers = new();

        public override string ConnectionId => "matters";
        public override string Provider => "imanage";

        public override string Seed(string name)
        {
            var number = _server.Seed(name, DocxFactory.Contract());
            _numbers[name] = number;
            return _server.IdOf(number);
        }

        public override byte[] Saved(string name) =>
            _server.BytesOf(_numbers[name], _server.VersionCount(_numbers[name]));

        public override void AssertConditionalSave(string name)
        {
            // A new version, added under the document lock after the latest version was read again.
            Assert.Equal(2, _server.VersionCount(_numbers[name]));
            Assert.True(_server.Locks > 0);
            Assert.True(_server.TokenRequests > 0);
        }

        public override string Obstruct(string name)
        {
            var number = _numbers[name];
            _server.OnLock = () => _server.AddVersion(number, DocxFactory.Contract());
            return ToolErrorCodes.VersionConflict;
        }

        public override void Configure(OfficeAgentMcpOptions options, bool creation, bool minimal) =>
            options.IManageConnections.Add(new IManageConnectionOptions
            {
                ConnectionId = ConnectionId,
                ServerUrl = IManageProviderTests.FakeIManage.Server,
                CustomerId = "42",
                ClientId = "officeagent-client",
                ClientSecret = IManageClientSecret,
                Username = "svc-officeagent",
                Password = IManagePassword,
                CreationFolderId = creation ? "ACTIVE!773" : string.Empty,
                DefaultChangeGranularity = minimal ? "Minimal" : "Span"
            });

        public override StorageTransports Transports() => new() { Http = _ => _server.Http };

        private string? _templates;

        public override string TemplatesFolder => _templates ??= _server.AddFolder("Precedents", parent: "ACTIVE!1");

        public override void SeedTemplate(string name) => _server.SeedIn(TemplatesFolder, name, DocxFactory.Contract());

        public override void Dispose() => _server.Dispose();
    }

    private sealed class GoogleDriveRig : Rig
    {
        private readonly GoogleDriveProviderTests.FakeDrive _drive = new();
        private readonly Dictionary<string, string> _ids = new();

        public override string ConnectionId => "drive";
        public override string Provider => "googledrive";

        public override string Seed(string name)
        {
            var id = _drive.Seed(name, DocxFactory.Contract());
            _ids[name] = id;
            return id;
        }

        public override byte[] Saved(string name) => _drive.Bytes(_ids[name]);

        public override void AssertConditionalSave(string name)
        {
            // Drive has no server-side precondition, so the provider reads the head revision
            // again between the download and the upload, and refuses a mismatch.
            var id = _ids[name];
            var requests = _drive.Requests;
            var download = requests.FindLastIndex(r => r.StartsWith($"GET {GoogleDriveProviderTests.FakeDrive.Api}/files/{id}", StringComparison.Ordinal) && r.Contains("alt=media"));
            var upload = requests.FindIndex(r => r.Contains($"{GoogleDriveProviderTests.FakeDrive.Upload}/files/{id}"));
            Assert.True(download >= 0 && upload > download, string.Join("\n", requests));
            Assert.Contains(requests.Skip(download + 1).Take(upload - download - 1),
                r => r.StartsWith($"GET {GoogleDriveProviderTests.FakeDrive.Api}/files/{id}", StringComparison.Ordinal) && !r.Contains("alt=media"));
            Assert.True(_drive.TokenRequests > 0);
        }

        public override string Obstruct(string name)
        {
            _drive.Lock(_ids[name]);
            return ToolErrorCodes.WriteRejected;
        }

        public override void Configure(OfficeAgentMcpOptions options, bool creation, bool minimal) =>
            options.GoogleDriveConnections.Add(new GoogleDriveConnectionOptions
            {
                ConnectionId = ConnectionId,
                KeyJson = DriveKeyJson,
                ApiBaseUrl = GoogleDriveProviderTests.FakeDrive.Api,
                UploadBaseUrl = GoogleDriveProviderTests.FakeDrive.Upload,
                CreationFolderId = creation ? GoogleDriveProviderTests.FakeDrive.FolderId : string.Empty,
                DefaultChangeGranularity = minimal ? "Minimal" : "Span"
            });

        public override StorageTransports Transports() => new() { Http = _ => _drive.Http };

        private string? _templates;

        public override string TemplatesFolder => _templates ??= _drive.SeedFolder("Templates", parent: "sharedRoot");

        public override void SeedTemplate(string name) => _drive.Seed(name, DocxFactory.Contract(), parent: TemplatesFolder);

        public override void Dispose() => _drive.Dispose();
    }

    private sealed class BoxRig : Rig
    {
        private readonly BoxProviderTests.FakeBox _box = new();
        private readonly Dictionary<string, string> _ids = new();
        private readonly Dictionary<string, string> _openedETags = new();

        public override string ConnectionId => "box";
        public override string Provider => "box";

        public override string Seed(string name)
        {
            var id = _box.Seed(name, DocxFactory.Contract());
            _ids[name] = id;
            _openedETags[name] = _box.ETag(id);
            return id;
        }

        public override byte[] Saved(string name) => _box.Bytes(_ids[name]);

        public override void AssertConditionalSave(string name)
        {
            // Box itself refuses the upload when the file changed: If-Match carries the opened etag.
            Assert.Equal(_openedETags[name], _box.LastIfMatch);
            Assert.True(_box.TokenRequests > 0);
        }

        public override string Obstruct(string name)
        {
            var id = _ids[name];
            _box.BeforeNextUpload = () => _box.EditElsewhere(id, new byte[] { 9 });
            return ToolErrorCodes.VersionConflict;
        }

        public override void Configure(OfficeAgentMcpOptions options, bool creation, bool minimal) =>
            options.BoxConnections.Add(new BoxConnectionOptions
            {
                ConnectionId = ConnectionId,
                ClientId = "box-client",
                ClientSecret = BoxClientSecret,
                SubjectType = "Enterprise",
                SubjectId = "424242",
                ApiBaseUrl = BoxProviderTests.FakeBox.Api,
                UploadBaseUrl = BoxProviderTests.FakeBox.Upload,
                CreationFolderId = creation ? BoxProviderTests.FakeBox.FolderId : string.Empty,
                DefaultChangeGranularity = minimal ? "Minimal" : "Span"
            });

        public override StorageTransports Transports() => new() { Http = _ => _box.Http };

        private string? _templates;

        public override string TemplatesFolder => _templates ??= _box.SeedFolder("Templates", parent: "0");

        public override void SeedTemplate(string name) => _box.Seed(name, DocxFactory.Contract(), parent: TemplatesFolder);

        public override void Dispose() => _box.Dispose();
    }

    private abstract class ObjectStoreRig : Rig
    {
        private readonly Dictionary<string, string> _openedETags = new();

        protected abstract FakeObjectStore Store { get; }

        public override string Seed(string name)
        {
            _openedETags[name] = Store.Seed("contracts/" + name, DocxFactory.Contract());
            return name;
        }

        public override byte[] Saved(string name) => Store.BytesOf("contracts/" + name);

        public override void AssertConditionalSave(string name)
        {
            // The store refuses the write itself unless the object still has the opened ETag.
            var put = Store.Requests.Single(r => r.Method == "PUT" && r.Name == "contracts/" + name);
            Assert.Equal(_openedETags[name], put.Headers["If-Match"]);
        }

        public override string Obstruct(string name)
        {
            Store.BeforeWrite = written => Store.Touch(written);
            return ToolErrorCodes.VersionConflict;
        }

        public override string TemplatesFolder => "templates";

        public override void SeedTemplate(string name) => Store.Seed("contracts/templates/" + name, DocxFactory.Contract());
    }

    private sealed class AzureBlobRig : ObjectStoreRig
    {
        private readonly FakeBlobService _store = new();

        protected override FakeObjectStore Store => _store;
        public override string ConnectionId => "blobs";
        public override string Provider => "azureblob";

        public override void Configure(OfficeAgentMcpOptions options, bool creation, bool minimal) =>
            options.AzureBlobConnections.Add(new AzureBlobConnectionOptions
            {
                ConnectionId = ConnectionId,
                ConnectionString = $"DefaultEndpointsProtocol=https;AccountName={FakeBlobService.Account};AccountKey={BlobAccountKey};EndpointSuffix=core.windows.net",
                ContainerName = FakeBlobService.Container,
                Prefix = "contracts/",
                DefaultChangeGranularity = minimal ? "Minimal" : "Span"
            });

        public override StorageTransports Transports() => new() { AzureBlob = _ => _store.Client() };
    }

    private sealed class AmazonS3Rig : ObjectStoreRig
    {
        private readonly FakeS3Service _store = new();

        protected override FakeObjectStore Store => _store;
        public override string ConnectionId => "bucket";
        public override string Provider => "amazons3";

        public override void Configure(OfficeAgentMcpOptions options, bool creation, bool minimal) =>
            options.AmazonS3Connections.Add(new AmazonS3ConnectionOptions
            {
                ConnectionId = ConnectionId,
                BucketName = FakeS3Service.Bucket,
                KeyPrefix = "contracts/",
                Region = "us-east-1",
                DefaultChangeGranularity = minimal ? "Minimal" : "Span"
            });

        public override StorageTransports Transports() => new() { AmazonS3 = _ => _store.Client() };
    }

    private static Rig NewRig(string provider) => provider switch
    {
        "imanage" => new IManageRig(),
        "googledrive" => new GoogleDriveRig(),
        "box" => new BoxRig(),
        "azureblob" => new AzureBlobRig(),
        "amazons3" => new AmazonS3Rig(),
        _ => throw new ArgumentOutOfRangeException(nameof(provider))
    };

    /// <summary>Every provider this suite covers, for theories.</summary>
    public static TheoryData<string> Providers => new() { "imanage", "googledrive", "box", "azureblob", "amazons3" };

    private sealed class Server : IDisposable
    {
        private readonly ServiceProvider _services;

        public Server(OfficeAgentMcpOptions options, StorageTransports? transports)
        {
            Options = options;
            var services = new ServiceCollection();
            if (transports is not null) services.AddSingleton(transports);
            services.AddOfficeAgentMcp(options);
            _services = services.BuildServiceProvider();
            Tools = _services.GetRequiredService<OfficeAgentTools>();
            McpTools = OfficeAgentMcpServer.CreateTools(Tools, options);
        }

        public OfficeAgentMcpOptions Options { get; }
        public OfficeAgentTools Tools { get; }
        public IList<ModelContextProtocol.Server.McpServerTool> McpTools { get; }
        public OfficeAgentClient Client => _services.GetRequiredService<OfficeAgentClient>();

        public string Description(string tool) => McpTools.Single(t => t.ProtocolTool.Name == tool).ProtocolTool.Description ?? string.Empty;

        public IDocumentProvider Provider(string connectionId) =>
            _services.GetServices<IDocumentProvider>().Single(p => p.ConnectionId == connectionId);

        public void Dispose() => _services.Dispose();
    }

    private static Server Start(Rig rig, bool creation = true, bool minimal = true, bool allowCreation = true)
    {
        var options = new OfficeAgentMcpOptions { AllowCreation = allowCreation };
        rig.Configure(options, creation, minimal);
        return new Server(options, rig.Transports());
    }

    // ── End to end ────────────────────────────────────────────────────────────

    [Theory]
    [MemberData(nameof(Providers))]
    public async Task A_document_is_registered_inspected_and_edited_through_the_provider(string provider)
    {
        using var rig = NewRig(provider);
        using var server = Start(rig);
        var source = rig.Seed("contract.docx");

        using var registered = JsonDocument.Parse(await server.Tools.RegisterDocument(rig.ConnectionId, source));
        Assert.True(registered.RootElement.TryGetProperty("documentId", out var documentId), registered.RootElement.ToString());
        Assert.Equal("contract.docx", registered.RootElement.GetProperty("name").GetString());

        var inspected = await server.Tools.InspectDocument(rig.ConnectionId, documentId.GetString()!);
        Assert.Contains(DocxFactory.DateText, inspected);

        // The connection's Minimal default applies to a plan that states no granularity.
        using var edited = JsonDocument.Parse(await server.Tools.EditDocument(rig.ConnectionId, source, Edit));
        Assert.True(edited.RootElement.GetProperty("committed").GetBoolean(), edited.RootElement.ToString());
        Assert.Equal("2020-01-01.", Assert.Single(DeletedText(rig.Saved("contract.docx"))));
        Assert.Equal("2021-01-01.", Assert.Single(InsertedText(rig.Saved("contract.docx"))));
        rig.AssertConditionalSave("contract.docx");
    }

    [Theory]
    [MemberData(nameof(Providers))]
    public async Task A_plan_that_states_its_granularity_wins_over_the_connection_default(string provider)
    {
        using var rig = NewRig(provider);
        using var server = Start(rig, minimal: true);
        var source = rig.Seed("span.docx");

        Assert.Equal(ChangeGranularity.Minimal, server.Client.DefaultChangeGranularityFor(rig.ConnectionId));
        using var edited = JsonDocument.Parse(await server.Tools.EditDocument(rig.ConnectionId, source, EditWithSpan));

        Assert.True(edited.RootElement.GetProperty("committed").GetBoolean(), edited.RootElement.ToString());
        Assert.Equal(DocxFactory.DateText, Assert.Single(DeletedText(rig.Saved("span.docx"))));
    }

    [Theory]
    [MemberData(nameof(Providers))]
    public void Without_a_configured_default_the_connection_marks_the_whole_span(string provider)
    {
        using var rig = NewRig(provider);
        using var server = Start(rig, minimal: false);

        Assert.Equal(ChangeGranularity.Span, server.Client.DefaultChangeGranularityFor(rig.ConnectionId));
    }

    [Theory]
    [MemberData(nameof(Providers))]
    public async Task A_conflict_or_refusal_reaches_the_agent_with_its_code_and_writes_nothing(string provider)
    {
        using var rig = NewRig(provider);
        using var server = Start(rig);
        var source = rig.Seed("contract.docx");
        var before = rig.Saved("contract.docx");
        var expected = rig.Obstruct("contract.docx");

        var result = await server.Tools.EditDocument(rig.ConnectionId, source, Edit);

        Assert.Contains($"\"{expected}\"", result);
        Assert.Contains("\"committed\":false", result.Replace(" ", string.Empty));
        Assert.DoesNotContain("\"writeOutcome\":\"unknown\"", result.Replace(" ", string.Empty));
        if (expected == ToolErrorCodes.WriteRejected)
            Assert.Equal(before, rig.Saved("contract.docx"));
    }

    [Theory]
    [MemberData(nameof(Providers))]
    public async Task A_disallowed_extension_is_refused_with_its_code(string provider)
    {
        using var rig = NewRig(provider);
        using var server = Start(rig);
        var source = rig.Seed("notes.xlsx");

        var result = await server.Tools.RegisterDocument(rig.ConnectionId, source);

        Assert.Contains($"\"{ToolErrorCodes.ExtensionNotAllowed}\"", result);
    }

    [Theory]
    [MemberData(nameof(Providers))]
    public async Task Create_document_files_a_new_document_through_the_provider(string provider)
    {
        using var rig = NewRig(provider);
        using var server = Start(rig, creation: true);
        Assert.Contains("create_document", server.McpTools.Select(t => t.ProtocolTool.Name));

        using var created = JsonDocument.Parse(await server.Tools.CreateDocument(rig.ConnectionId, "brief.docx"));

        Assert.True(created.RootElement.GetProperty("committed").GetBoolean(), created.RootElement.ToString());
        Assert.Equal("brief.docx", created.RootElement.GetProperty("outputName").GetString());
    }

    [Theory]
    [MemberData(nameof(Providers))]
    public async Task A_configured_templates_folder_is_listed_and_its_templates_register(string provider)
    {
        using var rig = NewRig(provider);
        var options = new OfficeAgentMcpOptions();
        rig.Configure(options, creation: true, minimal: false);
        var connection = (StorageConnectionOptions)Connection(options);
        connection.TemplatesFolder = rig.TemplatesFolder;
        connection.TemplateLocation = "organization";
        connection.AllowedExtensions = new List<string> { ".docx", ".dotx" };
        rig.SeedTemplate("letter.dotx");
        using var server = new Server(options, rig.Transports());

        var listing = Assert.IsAssignableFrom<IDocumentListingProvider>(server.Provider(rig.ConnectionId));
        var templates = Assert.IsAssignableFrom<IConnectionTemplateSource>(listing);
        var page = await listing.ListAsync(new DocumentListQuery { Templates = true });

        Assert.Equal(TemplateLocationKind.Organization, templates.TemplateLocation);
        Assert.Equal(rig.TemplatesFolder, templates.TemplatesFolder);
        var entry = Assert.Single(page.Items);
        Assert.True(entry.IsTemplate);
        using var registered = JsonDocument.Parse(await server.Tools.RegisterDocument(rig.ConnectionId, entry.Source));
        Assert.Equal("letter.dotx", registered.RootElement.GetProperty("name").GetString());
    }

    [Theory]
    [MemberData(nameof(Providers))]
    public void A_connection_without_a_templates_folder_reports_none(string provider)
    {
        using var rig = NewRig(provider);
        using var server = Start(rig);

        var templates = Assert.IsAssignableFrom<IConnectionTemplateSource>(server.Provider(rig.ConnectionId));

        Assert.Null(templates.TemplatesFolder);
        Assert.Equal(TemplateLocationKind.None, templates.TemplateLocation);
    }

    // ── Discovery ─────────────────────────────────────────────────────────────

    [Theory]
    [MemberData(nameof(Providers))]
    public async Task List_connections_reports_the_connection_and_what_it_can_do(string provider)
    {
        using var rig = NewRig(provider);
        var options = new OfficeAgentMcpOptions { AllowCreation = true };
        rig.Configure(options, creation: true, minimal: true);
        var connection = Assert.IsAssignableFrom<StorageConnectionOptions>(Connection(options));
        connection.AllowedExtensions = new List<string> { "docx", ".XLSX" };
        connection.DefaultChangeMode = "direct";
        connection.CalculateFormulas = true;
        using var server = new Server(options, rig.Transports());

        using var payload = JsonDocument.Parse(await OfficeAgentMcpServer.ConnectionsPayloadAsync(server.Tools, options));
        var entry = Assert.Single(payload.RootElement.EnumerateArray());

        Assert.Equal(rig.ConnectionId, entry.GetProperty("connectionId").GetString());
        Assert.Equal(rig.Provider, entry.GetProperty("provider").GetString());
        Assert.True(entry.GetProperty("canCreateDocuments").GetBoolean());
        Assert.Equal(new[] { ".docx", ".xlsx" }, entry.GetProperty("allowedExtensions").EnumerateArray().Select(e => e.GetString()));
        Assert.Equal("Direct", entry.GetProperty("defaultChangeMode").GetString());
        Assert.Equal("Minimal", entry.GetProperty("defaultChangeGranularity").GetString());
        Assert.True(entry.GetProperty("calculatesFormulas").GetBoolean());
        Assert.Contains($"\"{rig.ConnectionId}\" ({rig.Provider})", OfficeAgentMcpServer.InstructionsFor(options));
    }

    [Theory]
    [InlineData("imanage", "ACTIVE!12345.1", "an iManage document id such as \"ACTIVE!12345.1\"")]
    [InlineData("googledrive", "drive.google.com/file/d/", "a Google Drive file id or the file's Drive or Docs URL")]
    [InlineData("box", "app.box.com/file/", "a Box file id or the file's Box URL")]
    [InlineData("azureblob", "blob name relative to", "a blob name relative to this connection's folder")]
    [InlineData("amazons3", "object key relative to", "an object key relative to this connection's folder")]
    public void Register_and_open_describe_only_the_sources_of_configured_providers(string provider, string source, string instruction)
    {
        using var rig = NewRig(provider);
        using var server = Start(rig);

        foreach (var tool in new[] { "register_document", "open_document" })
        {
            var description = server.Description(tool);
            Assert.Contains(source, description);
            Assert.DoesNotContain("SharePoint", description);
            Assert.DoesNotContain("filesystem", description);
        }

        Assert.Contains($"a register_document source is {instruction}", OfficeAgentMcpServer.InstructionsFor(server.Options));
    }

    [Fact]
    public void Several_providers_are_described_together_and_a_filesystem_only_server_keeps_its_wording()
    {
        using var root = new TemporaryDirectory();
        var options = new OfficeAgentMcpOptions();
        options.FileSystemConnections.Add(new FileSystemConnectionOptions { ConnectionId = "documents", RootPath = root.Path });
        using var plain = new Server(options, transports: null);
        Assert.Contains("for a filesystem connection, a path under its root", plain.Description("register_document"));
        Assert.DoesNotContain("iManage", plain.Description("register_document"));

        using var iManage = new IManageRig();
        using var box = new BoxRig();
        iManage.Configure(options, creation: false, minimal: false);
        box.Configure(options, creation: false, minimal: false);
        using var mixed = new Server(options, iManage.Transports());

        var description = mixed.Description("register_document");
        Assert.Contains("for a filesystem connection", description);
        Assert.Contains("for an iManage connection", description);
        Assert.Contains("for a Box connection", description);
    }

    [Theory]
    [MemberData(nameof(Providers))]
    public void Creation_is_advertised_only_where_the_connection_has_a_destination(string provider)
    {
        using var rig = NewRig(provider);
        var objectStore = provider is "azureblob" or "amazons3";

        var without = new OfficeAgentMcpOptions { AllowCreation = true, AllowRegistration = false };
        rig.Configure(without, creation: false, minimal: false);
        var withFolder = new OfficeAgentMcpOptions { AllowCreation = true, AllowRegistration = false };
        rig.Configure(withFolder, creation: true, minimal: false);

        // An object store creates under its own container or bucket and prefix, so it always has
        // a destination; the others need a creation folder.
        Assert.Equal(objectStore, CanCreate(without));
        Assert.True(CanCreate(withFolder));
        using (var server = new Server(without, rig.Transports()))
            Assert.Equal(objectStore, server.McpTools.Any(t => t.ProtocolTool.Name == "create_document"));
        using (var server = new Server(withFolder, rig.Transports()))
            Assert.Contains(server.McpTools, t => t.ProtocolTool.Name == "create_document");
        if (!objectStore)
            Assert.Contains("create_document is not configured for this connection",
                OfficeAgentMcpServer.InstructionsFor(new OfficeAgentMcpOptions
                {
                    AllowCreation = true,
                    EphemeralConnectionId = "session",
                    IManageConnections = { new IManageConnectionOptions { ConnectionId = "matters" } }
                }));

        static bool CanCreate(OfficeAgentMcpOptions options)
        {
            using var payload = JsonDocument.Parse(OfficeAgentMcpServer.ConnectionsPayload(options));
            return Assert.Single(payload.RootElement.EnumerateArray()).GetProperty("canCreateDocuments").GetBoolean();
        }
    }

    // ── Secrets ───────────────────────────────────────────────────────────────

    [Fact]
    public async Task No_secret_appears_in_any_tool_result_description_or_instruction()
    {
        var rigs = new[] { "imanage", "googledrive", "box", "azureblob", "amazons3" }.Select(NewRig).ToList();
        try
        {
            var options = new OfficeAgentMcpOptions { AllowCreation = true };
            foreach (var rig in rigs) rig.Configure(options, creation: true, minimal: true);
            var http = rigs
                .Where(r => r.Transports().Http is not null)
                .ToDictionary(r => r.Provider, r => r.Transports().Http!);
            var transports = new StorageTransports
            {
                Http = name => http[name](name),
                AzureBlob = rigs.OfType<AzureBlobRig>().Single().Transports().AzureBlob,
                AmazonS3 = rigs.OfType<AmazonS3Rig>().Single().Transports().AmazonS3
            };
            using var server = new Server(options, transports);

            var seen = new List<string>
            {
                OfficeAgentMcpServer.InstructionsFor(options),
                OfficeAgentMcpServer.ConnectionsPayload(options),
                await OfficeAgentMcpServer.ConnectionsPayloadAsync(server.Tools, options),
                await server.Tools.DescribeCapabilities()
            };
            seen.AddRange(server.McpTools.Select(t => JsonSerializer.Serialize(t.ProtocolTool)));

            foreach (var rig in rigs)
            {
                var source = rig.Seed("contract.docx");
                seen.Add(await server.Tools.RegisterDocument(rig.ConnectionId, source));
                seen.Add(await server.Tools.OpenDocument(rig.ConnectionId, source));
                seen.Add(await server.Tools.EditDocument(rig.ConnectionId, source, Edit));
                seen.Add(await server.Tools.RegisterDocument(rig.ConnectionId, "https://elsewhere.example/contract.docx"));
                rig.Obstruct("contract.docx");
                seen.Add(await server.Tools.EditDocument(rig.ConnectionId, source, Edit));
            }

            foreach (var text in seen)
                AssertNoSecret(text);
        }
        finally
        {
            foreach (var rig in rigs) rig.Dispose();
        }
    }

    // ── Startup checks ────────────────────────────────────────────────────────

    public static TheoryData<string, string[]> BadConfigurations => new()
    {
        { "imanage-no-server", new[] { "iManage connection 'matters'", "ServerUrl", "OfficeAgent__IManageConnections__0__ServerUrl" } },
        { "imanage-http-server", new[] { "iManage connection 'matters'", "ServerUrl" } },
        { "imanage-no-customer", new[] { "CustomerId", "OfficeAgent__IManageConnections__0__CustomerId" } },
        { "imanage-no-client", new[] { "ClientId" } },
        { "imanage-no-password", new[] { "Password", "OfficeAgent__IManageConnections__0__Password", "environment variable" } },
        { "imanage-refresh-without-token", new[] { "RefreshToken", "RefreshToken grant" } },
        { "imanage-client-credentials-without-secret", new[] { "ClientSecret", "ClientCredentials grant" } },
        { "imanage-bad-grant", new[] { "GrantType 'Passwrd'", "Expected Password, RefreshToken or ClientCredentials" } },
        { "imanage-bad-folder", new[] { "CreationFolderId", "ACTIVE!773" } },
        { "imanage-bad-granularity", new[] { "DefaultChangeGranularity 'Minimum'", "Expected Span or Minimal" } },
        { "imanage-no-id", new[] { "iManage connection 0 has no ConnectionId", "OfficeAgent__IManageConnections__0__ConnectionId" } },
        { "drive-no-key", new[] { "Google Drive connection 'drive'", "KeyFilePath", "KeyJson" } },
        { "drive-two-keys", new[] { "exactly one service account key" } },
        { "drive-missing-key-file", new[] { "KeyFilePath that names no file" } },
        { "drive-unreadable-key", new[] { "Google Drive connection 'drive' is not configured correctly", "private key" } },
        { "box-no-secret", new[] { "Box connection 'box'", "ClientSecret", "OfficeAgent__BoxConnections__0__ClientSecret" } },
        { "box-no-subject", new[] { "SubjectId" } },
        { "box-bad-subject-type", new[] { "SubjectType 'Group'", "Expected Enterprise or User" } },
        { "box-bad-folder", new[] { "numeric Box folder id" } },
        { "blob-no-container", new[] { "Azure Blob connection 'blobs'", "ContainerName" } },
        { "blob-no-account", new[] { "ServiceUri", "ConnectionString" } },
        { "blob-both-accounts", new[] { "exactly one way to reach the account" } },
        { "blob-bad-connection-string", new[] { "Azure Blob connection 'blobs' is not configured correctly", "connection string is not valid" } },
        { "blob-bad-prefix", new[] { "Azure Blob connection 'blobs' is not configured correctly" } },
        { "s3-no-bucket", new[] { "Amazon S3 connection 'bucket'", "BucketName" } },
        { "s3-bad-endpoint", new[] { "ServiceUrl" } },
        { "s3-bad-mode", new[] { "DefaultChangeMode 'Tracke'", "Tracked or Direct" } },
        { "imanage-bad-templates", new[] { "iManage connection 'matters'", "TemplatesFolder", "ACTIVE!780" } },
        { "drive-bad-templates", new[] { "Google Drive connection 'drive'", "TemplatesFolder", "Google Drive folder id" } },
        { "box-bad-templates", new[] { "Box connection 'box'", "TemplatesFolder that is not a numeric Box folder id" } },
        { "blob-bad-templates", new[] { "Azure Blob connection 'blobs'", "TemplatesFolder", "inside the connection's prefix" } },
        { "s3-bad-template-location", new[] { "TemplateLocation 'Global'", "Expected Connection, Personal, Workgroup, Organization or None" } },
        { "s3-template-location-without-folder", new[] { "TemplateLocation Workgroup but no TemplatesFolder", "OfficeAgent__AmazonS3Connections__0__TemplatesFolder" } },
        { "s3-templates-folder-with-location-none", new[] { "TemplatesFolder and TemplateLocation None" } },
    };

    [Theory]
    [MemberData(nameof(BadConfigurations))]
    public void Bad_configuration_stops_the_server_at_startup_and_names_the_setting(string configuration, string[] expected)
    {
        var options = BadOptions(configuration);

        var error = Assert.Throws<InvalidOperationException>(() => OfficeAgentMcpServer.BuildToolset(options));

        foreach (var fragment in expected)
            Assert.Contains(fragment, error.Message);
        AssertNoSecret(error.ToString());
    }

    [Fact]
    public void Connections_bind_from_configuration_like_sharepoint_connections()
    {
        var configuration = new ConfigurationBuilder().AddInMemoryCollection(new Dictionary<string, string?>
        {
            ["OfficeAgent:IManageConnections:0:ConnectionId"] = "matters",
            ["OfficeAgent:IManageConnections:0:ServerUrl"] = "https://cloudimanage.com",
            ["OfficeAgent:IManageConnections:0:GrantType"] = "RefreshToken",
            ["OfficeAgent:IManageConnections:0:LockDuringSave"] = "false",
            ["OfficeAgent:IManageConnections:0:AllowedExtensions:0"] = ".pptx",
            ["OfficeAgent:GoogleDriveConnections:0:ConnectionId"] = "drive",
            ["OfficeAgent:GoogleDriveConnections:0:Subject"] = "legal@contoso.com",
            ["OfficeAgent:BoxConnections:0:ConnectionId"] = "box",
            ["OfficeAgent:BoxConnections:0:SubjectType"] = "User",
            ["OfficeAgent:AzureBlobConnections:0:ConnectionId"] = "blobs",
            ["OfficeAgent:AzureBlobConnections:0:Prefix"] = "contracts/",
            ["OfficeAgent:AzureBlobConnections:0:TemplatesFolder"] = "templates",
            ["OfficeAgent:AzureBlobConnections:0:TemplateLocation"] = "Organization",
            ["OfficeAgent:BoxConnections:0:TemplatesFolder"] = "4242",
            ["OfficeAgent:AmazonS3Connections:0:ConnectionId"] = "bucket",
            ["OfficeAgent:AmazonS3Connections:0:ForcePathStyle"] = "true",
            ["OfficeAgent:AmazonS3Connections:0:DefaultChangeGranularity"] = "Minimal",
        }).Build();

        var options = OfficeAgentConfiguration.Bind(configuration);

        var iManage = Assert.Single(options.IManageConnections);
        Assert.Equal("RefreshToken", iManage.GrantType);
        Assert.False(iManage.LockDuringSave);
        Assert.Equal(new[] { ".pptx" }, iManage.AllowedExtensions);
        Assert.Equal("legal@contoso.com", Assert.Single(options.GoogleDriveConnections).Subject);
        Assert.Equal("User", Assert.Single(options.BoxConnections).SubjectType);
        Assert.Equal(50 * 1024 * 1024, options.BoxConnections[0].MaximumBytes);
        Assert.Equal("contracts/", Assert.Single(options.AzureBlobConnections).Prefix);
        Assert.Equal(("templates", "Organization"), (options.AzureBlobConnections[0].TemplatesFolder, options.AzureBlobConnections[0].TemplateLocation));
        Assert.Equal(("4242", ""), (options.BoxConnections[0].TemplatesFolder, options.BoxConnections[0].TemplateLocation));
        Assert.True(Assert.Single(options.AmazonS3Connections).ForcePathStyle);
        Assert.Equal("Minimal", options.AmazonS3Connections[0].DefaultChangeGranularity);
    }

    [Fact]
    public void A_value_the_binder_cannot_read_is_reported_by_name_instead_of_dropping_the_connection()
    {
        var configuration = new ConfigurationBuilder().AddInMemoryCollection(new Dictionary<string, string?>
        {
            ["OfficeAgent:BoxConnections:0:ConnectionId"] = "box",
            ["OfficeAgent:BoxConnections:0:MaximumBytes"] = "fifty megabytes",
        }).Build();

        var error = Assert.Throws<InvalidOperationException>(() => OfficeAgentConfiguration.Bind(configuration));

        Assert.Contains("OfficeAgent:BoxConnections:0", error.Message);
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private static object Connection(OfficeAgentMcpOptions options) =>
        options.IManageConnections.Cast<object>()
            .Concat(options.GoogleDriveConnections)
            .Concat(options.BoxConnections)
            .Concat(options.AzureBlobConnections)
            .Concat(options.AmazonS3Connections)
            .Single();

    private static OfficeAgentMcpOptions BadOptions(string configuration)
    {
        IManageConnectionOptions IManage() => new()
        {
            ConnectionId = "matters", ServerUrl = "https://cloudimanage.com", CustomerId = "42",
            ClientId = "client", ClientSecret = IManageClientSecret, Username = "svc", Password = IManagePassword
        };
        GoogleDriveConnectionOptions Drive() => new() { ConnectionId = "drive", KeyJson = DriveKeyJson };
        BoxConnectionOptions Box() => new()
        {
            ConnectionId = "box", ClientId = "client", ClientSecret = BoxClientSecret, SubjectId = "424242"
        };
        AzureBlobConnectionOptions Blob() => new()
        {
            ConnectionId = "blobs", ContainerName = "docs", ServiceUri = "https://fakeaccount.blob.core.windows.net"
        };
        AmazonS3ConnectionOptions S3() => new() { ConnectionId = "bucket", BucketName = "contracts", Region = "eu-west-1" };

        var options = new OfficeAgentMcpOptions();
        switch (configuration)
        {
            case "imanage-no-server": options.IManageConnections.Add(With(IManage(), c => c.ServerUrl = "")); break;
            case "imanage-http-server": options.IManageConnections.Add(With(IManage(), c => c.ServerUrl = "http://cloudimanage.com")); break;
            case "imanage-no-customer": options.IManageConnections.Add(With(IManage(), c => c.CustomerId = " ")); break;
            case "imanage-no-client": options.IManageConnections.Add(With(IManage(), c => c.ClientId = "")); break;
            case "imanage-no-password": options.IManageConnections.Add(With(IManage(), c => c.Password = "")); break;
            case "imanage-refresh-without-token": options.IManageConnections.Add(With(IManage(), c => c.GrantType = "RefreshToken")); break;
            case "imanage-client-credentials-without-secret":
                options.IManageConnections.Add(With(IManage(), c => { c.GrantType = "clientcredentials"; c.ClientSecret = ""; }));
                break;
            case "imanage-bad-grant": options.IManageConnections.Add(With(IManage(), c => c.GrantType = "Passwrd")); break;
            case "imanage-bad-folder": options.IManageConnections.Add(With(IManage(), c => c.CreationFolderId = "ACTIVE!773.1")); break;
            case "imanage-bad-granularity": options.IManageConnections.Add(With(IManage(), c => c.DefaultChangeGranularity = "Minimum")); break;
            case "imanage-no-id": options.IManageConnections.Add(With(IManage(), c => c.ConnectionId = "")); break;
            case "drive-no-key": options.GoogleDriveConnections.Add(With(Drive(), c => c.KeyJson = "")); break;
            case "drive-two-keys": options.GoogleDriveConnections.Add(With(Drive(), c => c.KeyFilePath = typeof(McpStorageProviderTests).Assembly.Location)); break;
            case "drive-missing-key-file":
                options.GoogleDriveConnections.Add(With(Drive(), c => { c.KeyJson = ""; c.KeyFilePath = Path.Combine(Path.GetTempPath(), $"missing-{Guid.NewGuid():N}.json"); }));
                break;
            case "drive-unreadable-key":
                options.GoogleDriveConnections.Add(With(Drive(), c => c.KeyJson = DriveKeyJson.Replace("MII", "XXX")));
                break;
            case "box-no-secret": options.BoxConnections.Add(With(Box(), c => c.ClientSecret = "")); break;
            case "box-no-subject": options.BoxConnections.Add(With(Box(), c => c.SubjectId = "")); break;
            case "box-bad-subject-type": options.BoxConnections.Add(With(Box(), c => c.SubjectType = "Group")); break;
            case "box-bad-folder": options.BoxConnections.Add(With(Box(), c => c.CreationFolderId = "Contracts")); break;
            case "blob-no-container": options.AzureBlobConnections.Add(With(Blob(), c => c.ContainerName = "")); break;
            case "blob-no-account": options.AzureBlobConnections.Add(With(Blob(), c => c.ServiceUri = "")); break;
            case "blob-both-accounts":
                options.AzureBlobConnections.Add(With(Blob(), c => c.ConnectionString = $"AccountName=a;AccountKey={BlobAccountKey}"));
                break;
            case "blob-bad-connection-string":
                options.AzureBlobConnections.Add(With(Blob(), c => { c.ServiceUri = ""; c.ConnectionString = $"AccountKey={BlobAccountKey};Nonsense"; }));
                break;
            case "blob-bad-prefix": options.AzureBlobConnections.Add(With(Blob(), c => c.Prefix = "../outside")); break;
            case "s3-no-bucket": options.AmazonS3Connections.Add(With(S3(), c => c.BucketName = "")); break;
            case "s3-bad-endpoint": options.AmazonS3Connections.Add(With(S3(), c => c.ServiceUrl = "minio:9000")); break;
            case "s3-bad-mode": options.AmazonS3Connections.Add(With(S3(), c => c.DefaultChangeMode = "Tracke")); break;
            case "imanage-bad-templates": options.IManageConnections.Add(With(IManage(), c => c.TemplatesFolder = "Precedents")); break;
            case "drive-bad-templates": options.GoogleDriveConnections.Add(With(Drive(), c => c.TemplatesFolder = "Templates/Firm")); break;
            case "box-bad-templates": options.BoxConnections.Add(With(Box(), c => c.TemplatesFolder = "Templates")); break;
            case "blob-bad-templates": options.AzureBlobConnections.Add(With(Blob(), c => c.TemplatesFolder = "../templates")); break;
            case "s3-bad-template-location":
                options.AmazonS3Connections.Add(With(S3(), c => { c.TemplatesFolder = "templates"; c.TemplateLocation = "Global"; }));
                break;
            case "s3-template-location-without-folder": options.AmazonS3Connections.Add(With(S3(), c => c.TemplateLocation = "workgroup")); break;
            case "s3-templates-folder-with-location-none":
                options.AmazonS3Connections.Add(With(S3(), c => { c.TemplatesFolder = "templates"; c.TemplateLocation = "None"; }));
                break;
            default: throw new ArgumentOutOfRangeException(nameof(configuration));
        }
        return options;

        static T With<T>(T value, Action<T> change)
        {
            change(value);
            return value;
        }
    }

    private static void AssertNoSecret(string text)
    {
        foreach (var secret in new[] { IManagePassword, IManageClientSecret, BoxClientSecret, BlobAccountKey, "drive-key-id-Zq85" })
            Assert.DoesNotContain(secret, text);
        // Any recognisable slice of the private key.
        Assert.DoesNotContain("PRIVATE KEY", text);
        Assert.DoesNotContain(DriveKeyPem.Split('\n')[1].Trim(), text);
    }

    private static string NewPrivateKeyPem()
    {
        using var rsa = RSA.Create(2048);
        return rsa.ExportPkcs8PrivateKeyPem();
    }

    private static List<string> DeletedText(byte[] document) =>
        Revisions<DeletedRun>(document);

    private static List<string> InsertedText(byte[] document) =>
        Revisions<InsertedRun>(document);

    private static List<string> Revisions<T>(byte[] document) where T : DocumentFormat.OpenXml.OpenXmlElement
    {
        using var stream = new MemoryStream(document);
        using var package = WordprocessingDocument.Open(stream, isEditable: false);
        return package.MainDocumentPart!.Document.Body!.Descendants<T>().Select(run => run.InnerText).ToList();
    }

    private sealed class TemporaryDirectory : IDisposable
    {
        public TemporaryDirectory()
        {
            Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"officeagent-mcp-storage-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Path);
        }

        public string Path { get; }

        public void Dispose()
        {
            if (Directory.Exists(Path)) Directory.Delete(Path, recursive: true);
        }
    }
}
