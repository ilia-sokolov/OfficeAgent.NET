// Interactive sample: a Microsoft Agent Framework (MAF) agent backed by Azure
// OpenAI that edits a Word document through OfficeAgent.NET's opaque-id
// register-by-reference flow.
//
// The host registers one document with a provider connection and receives an
// opaque id. The LLM only ever sees that id; it cannot register, delete, or
// otherwise touch storage. Each successful apply returns the next id to use for
// follow-up edits.
//
// Storage provider (auto-selected):
//   - By default the sample uses a temp-rooted FILESYSTEM provider and stages a
//     local fixture under it (zero configuration).
//   - If SharePoint configuration is present (AGENT_SHAREPOINT_DOC set), the
//     sample uses the SHAREPOINT provider instead and registers an existing
//     document by its SharePoint/OneDrive URL or its driveId/itemId pair.
//
// Required environment variables (Azure OpenAI):
//
//   AZURE_OPENAI_ENDPOINT     e.g. https://my-resource.openai.azure.com
//   AZURE_OPENAI_DEPLOYMENT   the chat deployment name, e.g. gpt-4o-mini
//
// Authentication (pick one):
//
//   AZURE_OPENAI_API_KEY      uses ApiKeyCredential
//   (none)                    uses DefaultAzureCredential - recommended for
//                             managed identity / az login developer flows
//
// Filesystem mode (the default) - optional:
//
//   AGENT_DOC          path to seed the conversation with (default: ./sample.docx,
//                      generated on first run with a small contract fixture).
//   AGENT_STORAGE_DIR  filesystem root the provider registers paths under
//                      (default: a per-run temp directory).
//
// SharePoint mode - set AGENT_SHAREPOINT_DOC to enable; then:
//
//   AGENT_SHAREPOINT_DOC         the existing .docx to edit, as a SharePoint/OneDrive
//                                URL (e.g. "https://contoso.sharepoint.com/:w:/s/…")
//                                or a "driveId/itemId" pair (e.g. "b!9a3f…/01ABCDEF").
//   Authentication:
//     AGENT_SHAREPOINT_TENANT_ID + AGENT_SHAREPOINT_CLIENT_ID +
//       AGENT_SHAREPOINT_CLIENT_SECRET  for the app-only (client-credentials) flow.
//   Optional:
//     AGENT_SHAREPOINT_CONNECTION_ID  connection id (default: "sharepoint").
//
// Run:
//   dotnet run --project samples/AgentEdit

using System.ClientModel;
using Azure.AI.OpenAI;
using Azure.Identity;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Agents.AI;
using Microsoft.Extensions.AI;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using OfficeAgent.Abstractions;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.SharePoint;
using OfficeAgent.Word;

string endpoint = Required("AZURE_OPENAI_ENDPOINT");
string deployment = Required("AZURE_OPENAI_DEPLOYMENT");
string? apiKey = Environment.GetEnvironmentVariable("AZURE_OPENAI_API_KEY");

// SharePoint mode turns on as soon as a document source is configured; otherwise
// the sample stays on the zero-config filesystem provider.
string? sharePointDoc = Environment.GetEnvironmentVariable("AGENT_SHAREPOINT_DOC");
bool useSharePoint = !string.IsNullOrWhiteSpace(sharePointDoc);

string connectionId = useSharePoint
    ? (Environment.GetEnvironmentVariable("AGENT_SHAREPOINT_CONNECTION_ID") ?? "sharepoint")
    : "workspace";
string providerName = useSharePoint
    ? SharePointDocumentProvider.ProviderName
    : FileSystemDocumentProvider.ProviderName;

// Filesystem mode prepares a local storage root and a fixture; SharePoint mode
// needs neither (it edits an existing document already in the library).
string storageRoot = string.Empty;
string seedPath = string.Empty;
if (useSharePoint)
{
    Console.WriteLine($"Provider: sharepoint  connectionId={connectionId}");
}
else
{
    seedPath = Environment.GetEnvironmentVariable("AGENT_DOC") ?? "sample.docx";
    storageRoot = Environment.GetEnvironmentVariable("AGENT_STORAGE_DIR")
                  ?? Path.Combine(Path.GetTempPath(), $"officeagent-agentedit-{Guid.NewGuid():N}");
    Directory.CreateDirectory(storageRoot);
    Console.WriteLine($"Provider: filesystem  storage={storageRoot}");
    if (!File.Exists(seedPath))
    {
        File.WriteAllBytes(seedPath, BuildSampleContract());
        Console.WriteLine($"Generated fresh fixture at {Path.GetFullPath(seedPath)}");
    }
}

// ── Host: OfficeAgent + the selected document provider + logging ───────────
using var host = Host.CreateDefaultBuilder()
    .ConfigureLogging(b => b
        .SetMinimumLevel(LogLevel.Information)
        .AddSimpleConsole(o => { o.SingleLine = true; o.TimestampFormat = "HH:mm:ss "; }))
    .ConfigureServices(s =>
    {
        s.AddWordFormat();
        if (useSharePoint)
            ConfigureSharePoint(s, connectionId);
        else
            ConfigureFileSystem(s, connectionId, storageRoot);
        s.AddOfficeAgent();
    })
    .Build();

var officeClient = host.Services.GetRequiredService<OfficeAgentClient>();
var log = host.Services.GetRequiredService<ILogger<Program>>();

// ── Seed: register a document with the connection and get an opaque id back ──
// The provider is a registry of references; the host owns where the file lives,
// and gets back an opaque id to address it from then on.
DocumentReference seeded;
if (useSharePoint)
{
    // SharePoint registers an EXISTING document by its URL or driveId/itemId pair -
    // the sample does not create content in your library.
    seeded = await officeClient.RegisterAsync(connectionId, sharePointDoc!);
    log.LogInformation(
        "Registered SharePoint document. connectionId={ConnectionId} source={Source} documentId={DocumentId}",
        connectionId, sharePointDoc, seeded.ItemId);
}
else
{
    // Filesystem stages a local fixture under the storage root, then registers it.
    var stagedPath = StageFilesystemFixture(storageRoot, seedPath);
    seeded = await officeClient.RegisterAsync(connectionId, stagedPath);
    log.LogInformation(
        "Registered document with storage. connectionId={ConnectionId} path={Path} documentId={DocumentId}",
        connectionId, stagedPath, seeded.ItemId);
}

// ── Azure OpenAI IChatClient ───────────────────────────────────────────────
AzureOpenAIClient azureClient = apiKey is { Length: > 0 }
    ? new AzureOpenAIClient(new Uri(endpoint), new ApiKeyCredential(apiKey))
    : new AzureOpenAIClient(new Uri(endpoint), new DefaultAzureCredential());

IChatClient chatClient = azureClient
    .GetChatClient(deployment)
    .AsIChatClient()
    .AsBuilder()
    .UseFunctionInvocation()       // MEAI auto-executes tool calls
    .UseLogging(host.Services.GetRequiredService<ILoggerFactory>())
    .Build();

// ── OfficeAgent tools projected as MAF tools ───────────────────────────────
var officeTools = new OfficeAgentTools(officeClient);
IList<AITool> tools = officeTools.AsAIFunctions().Cast<AITool>().ToList();

var instructions = $$"""
    CURRENT DOCUMENT - use these values in every tool call:
      connectionId = "{{connectionId}}"
      documentId   = "{{seeded.ItemId}}"

    Each user message also starts with a [Current document: …] line carrying the
    up-to-date documentId (it changes after every save). Always use the id from
    the latest message. NEVER ask the user for a connectionId or documentId -
    the host supplies them; the user does not know or manage these values.

    You are a Microsoft Word contract-editing assistant.

    After a successful apply_plan, the returned outputDocumentId becomes the
    current document; the host reflects it in the next [Current document: …] line.

    Verb cookbook (concrete shapes are in the preview_plan description):
    - "replace X with Y"                    → changeText (mode "Tracked" unless asked otherwise)
    - "highlight / bold / italic / colour"  → format with run properties on a text span
    - "set heading style / alignment"       → format with styleId / alignment on a paragraph
    - "add / remove rows in the table"      → insertTableRows / removeTableRows on the table#N path from inspect_document.nodes
    - "add / remove a column"               → insertTableColumns / removeTableColumns
    - "fill the <Tag> placeholder"          → fill with target.tag = "<Tag>"
    - "set the document title / property"   → setProperty
    - "accept / reject revisions"           → revision
    - "insert / remove an image"            → insertImage (host-supplied base64Bytes, or imageConnectionId+imageDocumentId pre-registered by the host) / removeImage by image#N

    {{OfficeAgentTools.SystemPromptGuidance}}

    When you have completed a request, briefly summarise what changed.
    """;

AIAgent agent = new ChatClientAgent(
    chatClient,
    instructions: instructions,
    name: "OfficeAgent",
    description: "Edits Word documents using the OfficeAgent.NET toolkit (provider-secured).",
    tools: tools,
    loggerFactory: host.Services.GetRequiredService<ILoggerFactory>(),
    services: host.Services);

var session = await agent.CreateSessionAsync();

Console.WriteLine();
Console.WriteLine("OfficeAgent ready. Try requests like:");
Console.WriteLine("  - What's in the document?");
Console.WriteLine("  - Replace every 'Acme Corp' with 'Globex Inc.' as a tracked change.");
Console.WriteLine("  - Highlight 'Effective date' in yellow.");
Console.WriteLine("  - Add a row to the milestones table: Beta, 2026-08-15.");
Console.WriteLine("  - Set the document title to 'Service Agreement v2'.");
Console.WriteLine("  - Fill the ClientName content control with 'Globex'.");
Console.WriteLine("Type 'export <output.docx>' to write the current revision to disk, or 'quit' to exit.");
Console.WriteLine();

string activeDocumentId = seeded.ItemId;

while (true)
{
    Console.Write("> ");
    var input = Console.ReadLine();
    if (string.IsNullOrWhiteSpace(input)) continue;
    var trimmed = input.Trim();
    if (trimmed.Equals("quit", StringComparison.OrdinalIgnoreCase)) break;

    if (trimmed.StartsWith("export ", StringComparison.OrdinalIgnoreCase))
    {
        var outputPath = trimmed["export ".Length..].Trim();
        try
        {
            var exported = await ExportDocumentAsync(officeClient, providerName, connectionId, activeDocumentId, outputPath);
            log.LogInformation(
                "Returned {DocumentId} to the user as {OutputPath} ({ContentType})",
                activeDocumentId,
                exported.Path,
                exported.ContentType);
        }
        catch (Exception ex)
        {
            log.LogError(ex, "Export failed");
        }
        continue;
    }

    try
    {
        // Prefix every turn with the live document context. The static system prompt
        // goes stale as soon as apply_plan mints a new id; carrying the current id in
        // the latest message keeps the model anchored to the right document and stops
        // it from asking the user for ids.
        var turnInput =
            $"[Current document: connectionId=\"{connectionId}\", documentId=\"{activeDocumentId}\"]\n{input}";

        var response = await agent.RunAsync(turnInput, session);
        Console.WriteLine();
        Console.WriteLine(response.ToString());
        Console.WriteLine();

        // Best-effort: pick up the latest outputDocumentId returned by apply_plan
        // so follow-up turns and `export` track the most recent revision.
        activeDocumentId = LatestSavedId(response, activeDocumentId);
    }
    catch (Exception ex)
    {
        log.LogError(ex, "Agent run failed");
    }
}

return 0;

static async Task<(string Path, string ContentType)> ExportDocumentAsync(
    OfficeAgentClient client,
    string providerName,
    string connectionId,
    string documentId,
    string outputPath)
{
    // The host-not the LLM-resolves the opaque id and delivers the bytes. A web
    // or chat host would pass this stream to its download/attachment API instead.
    // The reference is built for whichever provider is in use; resolution is by
    // (provider, connectionId).
    using var content = await client.OpenReadAsync(
        DocumentReference.For(providerName, connectionId, documentId));
    using var file = File.Create(outputPath);
    await content.Stream.CopyToAsync(file);

    return (
        Path.GetFullPath(outputPath),
        content.Reference.ContentType ?? "application/octet-stream");
}

static string Required(string name) =>
    Environment.GetEnvironmentVariable(name)
    ?? throw new InvalidOperationException(
        $"Environment variable {name} is required. See the comment block at the top of Program.cs.");

static void ConfigureFileSystem(IServiceCollection services, string connectionId, string storageRoot) =>
    services.AddFileSystemDocumentProvider(connectionId, storageRoot);

static void ConfigureSharePoint(IServiceCollection services, string connectionId)
{
    services.AddSingleton(new HttpClient());

    // The app client-credentials flow acquires Graph tokens from the configured Entra app.
    // For a per-user identity instead, configure On-Behalf-Of on a hosted HTTP server.
    var options = new AppOnlyOptions
    {
        TenantId = Required("AGENT_SHAREPOINT_TENANT_ID"),
        ClientId = Required("AGENT_SHAREPOINT_CLIENT_ID"),
        ClientSecret = Required("AGENT_SHAREPOINT_CLIENT_SECRET")
    };
    services.AddSingleton<IAccessTokenProvider>(sp =>
        new AppOnlyAccessTokenProvider(options, sp.GetRequiredService<HttpClient>()));

    services.AddSharePointDocumentProvider(connectionId, o => o.AllowedExtensions = new[] { ".docx" });
}

static string StageFilesystemFixture(string storageRoot, string seedPath)
{
    // A per-call subdirectory keeps the display name intact while avoiding
    // collisions across runs. The provider stores only the reference, not the bytes.
    var seedDirectory = Path.Combine(storageRoot, Guid.NewGuid().ToString("N"));
    Directory.CreateDirectory(seedDirectory);
    var stagedPath = Path.Combine(seedDirectory, Path.GetFileName(seedPath));
    File.Copy(seedPath, stagedPath, overwrite: true);
    return stagedPath;
}

static string LatestSavedId(AgentResponse response, string fallback)
{
    // apply_plan tool results expose outputDocumentId in their JSON payload. Scan in
    // reverse so the most recent save wins.
    var text = response.ToString() ?? string.Empty;
    var key = "\"outputDocumentId\":\"";
    var last = text.LastIndexOf(key, StringComparison.Ordinal);
    if (last < 0) return fallback;
    var start = last + key.Length;
    var end = text.IndexOf('"', start);
    return end > start ? text.Substring(start, end - start) : fallback;
}

// ── A small in-memory contract fixture for first-run convenience ────────────
static byte[] BuildSampleContract()
{
    using var ms = new MemoryStream();
    using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
    {
        var main = doc.AddMainDocumentPart();

        var heading = Para("00000001", "Heading1", new Run(new Text("Service Agreement")));
        var clause = Para("00000002", null,
            new Run(new Text("Acme Corp shall provide services to Acme Corp.") { Space = SpaceProcessingModeValues.Preserve }));

        var clientControl = new SdtRun(
            new SdtProperties(new Tag { Val = "ClientName" }, new SdtId { Val = 101 }),
            new SdtContentRun(new Run(new Text("PLACEHOLDER"))));
        var clientPara = Para("00000003", null,
            new Run(new Text("Client: ") { Space = SpaceProcessingModeValues.Preserve }));
        clientPara.AppendChild(clientControl);

        var date = Para("00000004", null, new Run(new Text("Effective date: 2026-06-05.")));

        var table = new Table(
            new TableProperties(new TableStyle { Val = "TableGrid" },
                new TableBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 4 },
                    new BottomBorder { Val = BorderValues.Single, Size = 4 },
                    new LeftBorder { Val = BorderValues.Single, Size = 4 },
                    new RightBorder { Val = BorderValues.Single, Size = 4 },
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                    new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 })),
            Row(bold: true, "Milestone", "Date"),
            Row(bold: false, "Kickoff", "2026-06-15"));

        main.Document = new Document(new Body(heading, clause, clientPara, date, table));

        var stylePart = main.AddNewPart<StyleDefinitionsPart>();
        stylePart.Styles = new Styles(
            new Style { Type = StyleValues.Paragraph, StyleId = "Heading1", StyleName = new StyleName { Val = "heading 1" } },
            new Style { Type = StyleValues.Paragraph, StyleId = "Quote", StyleName = new StyleName { Val = "Quote" } });

        main.Document.Save();
        stylePart.Styles.Save();
    }
    return ms.ToArray();

    static Paragraph Para(string id, string? styleId, params OpenXmlElement[] runs)
    {
        var p = new Paragraph { ParagraphId = id };
        if (styleId is not null)
            p.ParagraphProperties = new ParagraphProperties(new ParagraphStyleId { Val = styleId });
        foreach (var r in runs) p.AppendChild(r);
        return p;
    }

    static TableRow Row(bool bold, params string[] cells)
    {
        var row = new TableRow();
        foreach (var text in cells)
        {
            var run = new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
            if (bold) run.RunProperties = new RunProperties(new Bold());
            row.AppendChild(new TableCell(new Paragraph(run)));
        }
        return row;
    }
}
