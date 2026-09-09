// Contract review agent: Microsoft Agent Framework for judgement, OfficeAgent.NET
// for the document.
//
// The model reads the contract and records structured findings. It has no tool
// that writes. Host code turns findings into one plan, previews it, and writes a
// reviewed copy - so the review is either fully applied or not applied at all,
// and the source contract is never a write target.
//
// Required environment variables (Azure OpenAI):
//   AZURE_OPENAI_ENDPOINT     e.g. https://my-resource.openai.azure.com
//   AZURE_OPENAI_DEPLOYMENT   the chat deployment name
//   AZURE_OPENAI_API_KEY      optional; omit to use DefaultAzureCredential
//
// Optional:
//   REVIEW_CONTRACT   path to an existing .docx to review (default: a generated
//                     contract.docx inside REVIEW_STORAGE)
//   REVIEW_PLAYBOOK   path to the playbook JSON (default: bundled playbook.json)
//   REVIEW_STORAGE    filesystem root for the provider (default: a temp directory)
//
// Run:
//   dotnet run --project samples/ContractReview

using System.ClientModel;
using Azure.AI.OpenAI;
using Azure.Identity;
using Microsoft.Extensions.AI;
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Samples.ContractReview;
using OfficeAgent.Word;

string endpoint = Required("AZURE_OPENAI_ENDPOINT");
string deployment = Required("AZURE_OPENAI_DEPLOYMENT");
string? apiKey = Environment.GetEnvironmentVariable("AZURE_OPENAI_API_KEY");

string storageRoot = Environment.GetEnvironmentVariable("REVIEW_STORAGE")
    ?? Path.Combine(Path.GetTempPath(), "officeagent-review-" + Guid.NewGuid().ToString("N")[..8]);
Directory.CreateDirectory(storageRoot);

string? configuredContract = Environment.GetEnvironmentVariable("REVIEW_CONTRACT");
string contractPath;
if (string.IsNullOrWhiteSpace(configuredContract))
{
    contractPath = Path.Combine(storageRoot, "contract.docx");
    File.WriteAllBytes(contractPath, ContractFixture.Create());
    Console.WriteLine($"Generated a sample contract at {contractPath}");
}
else
{
    contractPath = Path.GetFullPath(configuredContract);
    if (!File.Exists(contractPath))
    {
        throw new FileNotFoundException(
            $"REVIEW_CONTRACT does not exist: {contractPath}",
            contractPath);
    }
}

if (!IsUnder(storageRoot, contractPath))
{
    // The provider only registers paths under its root, so copy an external
    // contract in rather than widening the connection boundary. Never overwrite
    // something already in the storage root - it is a directory the caller may
    // have chosen, not scratch space we own.
    var staged = Path.Combine(storageRoot, Path.GetFileName(contractPath));
    if (File.Exists(staged))
    {
        var stem = Path.GetFileNameWithoutExtension(staged);
        var ext = Path.GetExtension(staged);
        staged = Path.Combine(storageRoot, $"{stem}.{Guid.NewGuid():N}{ext}");
    }

    File.Copy(contractPath, staged);
    contractPath = staged;
}

// The bundled playbook is copied beside the executable, not into whatever
// directory dotnet run was invoked from.
string playbookPath = Environment.GetEnvironmentVariable("REVIEW_PLAYBOOK")
    ?? Path.Combine(AppContext.BaseDirectory, "playbook.json");
var playbook = Playbook.FromFile(playbookPath);

// ── OfficeAgent: one connection, one registered document ───────────────────
const string ConnectionId = "contracts";

using var services = new ServiceCollection()
    .AddWordFormat()
    .AddFileSystemDocumentProvider(ConnectionId, storageRoot, o => o.AllowedExtensions = new[] { ".docx" })
    .AddOfficeAgent()
    .BuildServiceProvider();

var officeClient = services.GetRequiredService<OfficeAgentClient>();
var contract = await officeClient.RegisterAsync(ConnectionId, contractPath);

// ── Azure OpenAI chat client with automatic function invocation ────────────
AzureOpenAIClient azureClient = apiKey is { Length: > 0 }
    ? new AzureOpenAIClient(new Uri(endpoint), new ApiKeyCredential(apiKey))
    : new AzureOpenAIClient(new Uri(endpoint), new DefaultAzureCredential());

IChatClient chatClient = azureClient
    .GetChatClient(deployment)
    .AsIChatClient()
    .AsBuilder()
    .UseFunctionInvocation()
    .Build();

// ── Review ─────────────────────────────────────────────────────────────────
var agent = new ReviewAgent(chatClient, officeClient);
var runner = new ReviewRunner(officeClient);

string outputName = Path.GetFileNameWithoutExtension(contractPath) + ".reviewed.docx";

Console.WriteLine($"Reviewing {Path.GetFileName(contractPath)} against '{playbook.Name}' ({playbook.Rules.Count} rules)...");
Console.WriteLine();

var result = await runner.RunAsync(
    ConnectionId,
    contract.ItemId,
    playbook,
    (screening, ct) => agent.JudgeAsync(ConnectionId, contract.ItemId, playbook, screening, ct),
    outputName);

Console.WriteLine(ReviewRunner.Report(playbook, result));

if (result.Status == ReviewStatus.Applied)
{
    Console.WriteLine($"Open {Path.Combine(storageRoot, result.OutputName!)} in Word to accept or reject the redlines.");
}

// A path-boundary test, not a string prefix: "C:\docs-old" is not under
// "C:\docs", though one string starts with the other.
static bool IsUnder(string root, string candidate)
{
    var relative = Path.GetRelativePath(Path.GetFullPath(root), Path.GetFullPath(candidate));
    return relative != "."
        && !relative.StartsWith(".." + Path.DirectorySeparatorChar, StringComparison.Ordinal)
        && relative != ".."
        && !Path.IsPathRooted(relative);
}

static string Required(string name) =>
    Environment.GetEnvironmentVariable(name)
    ?? throw new InvalidOperationException($"Environment variable {name} is required.");
