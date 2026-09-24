// Requirements review: a Microsoft Agent Framework agent calls TypeSafe Jev as a function
// tool, a drafting agent proposes wording, OfficeAgent.NET handles Word, and host code routes.
//
// Offline by default: Jev answers come from a scripted HTTP handler in the Jev wire format,
// and both agents use scripted IChatClient instances. No network, no credentials.
//
//   dotnet run --project samples/RequirementsReview -- --out ./review-out
//   dotnet run --project samples/RequirementsReview -- --out ./review-out --approve-as ops-lead@example.com
//
// Live paths are opt-in and read credentials from the environment only:
//   --live-jev       TYPESAFE_API_KEY
//   --live-agents    AZURE_OPENAI_ENDPOINT, AZURE_OPENAI_DEPLOYMENT, AZURE_OPENAI_API_KEY
//
// Other options:
//   --document <path>        review an existing .docx instead of the generated supplier proposal
//                            (edit privacy-requirements.json to match its headings)
//   --legacy                 run the earlier installation-method-statement scenario
//   --reject-as <reviewer>   record a rejection instead of an approval
//
// --approve-as stands in for an authenticated reviewer pressing "approve" after reading the
// preview. A real host takes the reviewer id from its sign-in, never from an argument.

using System.ClientModel;
using Microsoft.Extensions.AI;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Samples.RequirementsReview;
using OfficeAgent.Word;
using OpenAI;

var options = Args.Parse(args);

var outDir = Path.GetFullPath(options.Out ?? Path.Combine(Path.GetTempPath(), "requirements-review-" + Guid.NewGuid().ToString("N")[..8]));
Directory.CreateDirectory(outDir);

// A rerun into the same folder keeps everything an earlier run wrote: the document it
// reviewed, its reviewed copy, and its audit record. New files get a time stamp.
var stamp = DateTime.UtcNow.ToString("yyyyMMddTHHmmss");
var suffix = File.Exists(Path.Combine(outDir, "audit.json")) ? "." + stamp : string.Empty;

string documentPath;
if (options.Document is { } supplied)
{
    var source = Path.GetFullPath(supplied);
    documentPath = Path.Combine(outDir, Path.GetFileName(source));
    if (!string.Equals(source, documentPath, StringComparison.OrdinalIgnoreCase))
    {
        if (!File.Exists(documentPath))
        {
            File.Copy(source, documentPath);
        }
        else if (!File.ReadAllBytes(source).AsSpan().SequenceEqual(File.ReadAllBytes(documentPath)))
        {
            Console.Error.WriteLine($"A different {Path.GetFileName(source)} is already in {outDir}. Use a new --out folder.");
            return 1;
        }
    }
}
else
{
    documentPath = Path.Combine(outDir, options.Legacy ? MethodStatementFixture.FileName : PrivacyProposalFixture.FileName);
    if (!File.Exists(documentPath))
    {
        File.WriteAllBytes(documentPath, options.Legacy ? MethodStatementFixture.Create() : PrivacyProposalFixture.Create());
    }
}

var requirements = RequirementSet.FromFile(Path.Combine(
    AppContext.BaseDirectory, options.Legacy ? "requirements.json" : "privacy-requirements.json"));
var policy = ReviewPolicy.FromFile(Path.Combine(AppContext.BaseDirectory, "review-policy.json"));

const string ConnectionId = "reviews";
using var services = new ServiceCollection()
    .AddLogging(b => b.AddSimpleConsole(o => o.SingleLine = true).SetMinimumLevel(LogLevel.Information))
    .AddWordFormat()
    .AddFileSystemDocumentProvider(ConnectionId, outDir, o => o.AllowedExtensions = new[] { ".docx" })
    .AddOfficeAgent()
    .BuildServiceProvider();

var logger = services.GetRequiredService<ILoggerFactory>().CreateLogger("RequirementsReview");
var office = services.GetRequiredService<OfficeAgentClient>();
var document = await office.RegisterAsync(ConnectionId, documentPath);

// ── Evaluator: the real Jev adapter either way; only the transport differs offline ──
var jevOptions = new JevOptions
{
    Model = policy.EvaluatorModel,
    Timeout = policy.EvaluatorTimeout,
    MaxEvidenceCharacters = policy.MaxEvidenceCharacters,
};

IRequirementEvaluator evaluator;
EvaluatorInfo evaluatorInfo;
if (options.LiveJev)
{
    var key = Environment.GetEnvironmentVariable("TYPESAFE_API_KEY");
    if (string.IsNullOrWhiteSpace(key))
    {
        Console.Error.WriteLine("--live-jev needs TYPESAFE_API_KEY in the environment.");
        return 1;
    }

    evaluator = new JevRequirementEvaluator(new HttpClient(), jevOptions, key, logger);
    evaluatorInfo = new EvaluatorInfo(JevRequirementEvaluator.ProviderName, jevOptions.Endpoint.ToString());
}
else
{
    var scriptedTransport = options.Legacy
        ? new ScriptedJevHandler()
        : new ScriptedJevHandler(id => ScriptedAnswer.PrivacyDemo.TryGetValue(id, out var answer)
            ? ScriptedJevHandler.Json(System.Net.HttpStatusCode.OK, answer.ToResponse(ScriptedJevHandler.ScriptedModel).ToJsonString())
            : ScriptedJevHandler.Json(System.Net.HttpStatusCode.InternalServerError, "{\"detail\":\"no scripted answer\"}"));
    evaluator = new JevRequirementEvaluator(new HttpClient(scriptedTransport), jevOptions with { ProviderName = ScriptedJevHandler.ProviderLabel }, "offline-placeholder", logger);
    evaluatorInfo = new EvaluatorInfo(ScriptedJevHandler.ProviderLabel, "offline: ScriptedJevHandler (answers invented for the demo)");
}

// ── Agent Framework reviewer and drafter over live or scripted chat clients ──
IChatClient chatClient;
if (options.LiveAgents)
{
    var endpoint = Environment.GetEnvironmentVariable("AZURE_OPENAI_ENDPOINT");
    var deployment = Environment.GetEnvironmentVariable("AZURE_OPENAI_DEPLOYMENT");
    var key = Environment.GetEnvironmentVariable("AZURE_OPENAI_API_KEY");
    if (string.IsNullOrWhiteSpace(endpoint) || string.IsNullOrWhiteSpace(deployment) || string.IsNullOrWhiteSpace(key))
    {
        Console.Error.WriteLine("--live-agents needs AZURE_OPENAI_ENDPOINT, AZURE_OPENAI_DEPLOYMENT and AZURE_OPENAI_API_KEY.");
        return 1;
    }

    chatClient = new OpenAIClient(
            new ApiKeyCredential(key),
            new OpenAIClientOptions { Endpoint = new Uri(endpoint.TrimEnd('/') + "/openai/v1/") })
        .GetChatClient(deployment)
        .AsIChatClient();
}
else
{
    chatClient = new ScriptedChatClient();
}

var reviewChatClient = options.LiveAgents ? chatClient : new ScriptedReviewChatClient();
var workflow = new RequirementsWorkflow(
    office, requirements, policy,
    new AgentRequirementReviewer(reviewChatClient, evaluator),
    new AgentProposalDrafter(chatClient),
    logger: logger);

// ── Propose: evidence, checks, typed judgements, routing, drafting, preview. No write. ──
var proposal = await workflow.ProposeAsync(ConnectionId, document.ItemId);
CommitOutcome? commit = null;

if (options.ReviewerId is { } reviewer)
{
    if (proposal.Status != ProposalStatus.AwaitingApproval)
    {
        Console.Error.WriteLine($"Nothing to approve: proposal is {proposal.Status}.");
    }
    else
    {
        var decision = options.Reject ? ReviewerDecision.Rejected : ReviewerDecision.Approved;
        var approval = ReviewerApproval.For(proposal, reviewer, decision, DateTimeOffset.UtcNow);
        var outputName = $"{Path.GetFileNameWithoutExtension(documentPath)}.reviewed{suffix}.docx";
        commit = await new AuthorizedCommitter(office).CommitAsync(proposal, approval, outputName);
    }
}

var audit = AuditRecord.Build(proposal, evaluatorInfo, commit);
await File.WriteAllTextAsync(Path.Combine(outDir, $"audit{suffix}.json"), AuditRecord.ToJson(audit));
var report = Report.Render(proposal, commit);
await File.WriteAllTextAsync(Path.Combine(outDir, $"report{suffix}.md"), report);

Console.WriteLine();
Console.WriteLine(report);
Console.WriteLine($"Output directory: {outDir}");

return proposal.Status switch
{
    ProposalStatus.AwaitingApproval when commit is null => 0,
    ProposalStatus.AwaitingApproval when commit!.Status is CommitStatus.Committed or CommitStatus.RejectedByReviewer => 0,
    ProposalStatus.NothingToPropose => 0,
    _ => 2,
};

internal sealed record Args(string? Out, string? Document, bool LiveJev, bool LiveAgents, string? ReviewerId, bool Reject, bool Legacy)
{
    public static Args Parse(string[] args)
    {
        string? output = null, documentPath = null, reviewer = null;
        bool liveJev = false, liveAgents = false, reject = false, legacy = false;
        for (var i = 0; i < args.Length; i++)
        {
            switch (args[i])
            {
                case "--out": output = Value(args, ref i); break;
                case "--document": documentPath = Value(args, ref i); break;
                case "--live-jev": liveJev = true; break;
                case "--live-agents": liveAgents = true; break;
                case "--live-drafter": liveAgents = true; break; // Older spelling.
                case "--legacy": legacy = true; break;
                case "--approve-as": reviewer = Value(args, ref i); break;
                case "--reject-as": reviewer = Value(args, ref i); reject = true; break;
                default: throw new ArgumentException($"Unknown argument '{args[i]}'.");
            }
        }

        return new Args(output, documentPath, liveJev, liveAgents, reviewer, reject, legacy);
    }

    private static string Value(string[] args, ref int i) =>
        ++i < args.Length ? args[i] : throw new ArgumentException($"{args[i - 1]} needs a value.");
}
