using System.IO.Compression;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using Microsoft.Extensions.AI;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Samples.RequirementsReview;
using OfficeAgent.Word;

namespace RequirementsReview.Tests;

/// <summary>
/// A disposable review environment: a temp-rooted filesystem connection with the method statement
/// registered, the bundled requirement set and policy, and helpers for the produced .docx.
/// </summary>
internal sealed class Harness : IDisposable
{
    public const string ConnectionId = "reviews";
    public const string Reviewer = "ops-lead@example.com";
    public const string FakeKey = "tsk_TEST_ONLY_7f3a9c_do-not-log";

    private readonly ServiceProvider _services;

    private Harness(ServiceProvider services, string root, string path, string documentId)
    {
        _services = services;
        Root = root;
        SourcePath = path;
        DocumentId = documentId;
        Client = services.GetRequiredService<OfficeAgentClient>();
    }

    public string Root { get; }

    public string SourcePath { get; }

    public string DocumentId { get; }

    public OfficeAgentClient Client { get; }

    public RequirementSet Requirements { get; set; } = RequirementSet.FromFile(Path.Combine(AppContext.BaseDirectory, "requirements.json"));

    public ReviewPolicy Policy { get; } = ReviewPolicy.FromFile(Path.Combine(AppContext.BaseDirectory, "review-policy.json"));

    public CapturingLogger Log { get; } = new();

    public static async Task<Harness> CreateAsync(MethodStatementOptions? options = null)
    {
        var root = Path.Combine(Path.GetTempPath(), "oa-req-tests-" + Guid.NewGuid().ToString("N")[..10]);
        Directory.CreateDirectory(root);
        var path = Path.Combine(root, MethodStatementFixture.FileName);
        File.WriteAllBytes(path, MethodStatementFixture.Create(options));

        var services = new ServiceCollection()
            .AddWordFormat()
            .AddFileSystemDocumentProvider(ConnectionId, root, o => o.AllowedExtensions = new[] { ".docx" })
            .AddOfficeAgent()
            .BuildServiceProvider();

        var client = services.GetRequiredService<OfficeAgentClient>();
        var reference = await client.RegisterAsync(ConnectionId, path);
        return new Harness(services, root, path, reference.ItemId);
    }

    public static async Task<Harness> CreatePrivacyAsync()
    {
        var root = Path.Combine(Path.GetTempPath(), "oa-privacy-tests-" + Guid.NewGuid().ToString("N")[..10]);
        Directory.CreateDirectory(root);
        var path = Path.Combine(root, PrivacyProposalFixture.FileName);
        File.WriteAllBytes(path, PrivacyProposalFixture.Create());

        var services = new ServiceCollection()
            .AddWordFormat()
            .AddFileSystemDocumentProvider(ConnectionId, root, o => o.AllowedExtensions = new[] { ".docx" })
            .AddOfficeAgent()
            .BuildServiceProvider();

        var client = services.GetRequiredService<OfficeAgentClient>();
        var reference = await client.RegisterAsync(ConnectionId, path);
        var harness = new Harness(services, root, path, reference.ItemId)
        {
            Requirements = RequirementSet.FromFile(Path.Combine(AppContext.BaseDirectory, "privacy-requirements.json")),
        };
        return harness;
    }

    /// <summary>The real Jev adapter over a scripted transport.</summary>
    public JevRequirementEvaluator Jev(ScriptedJevHandler? handler = null, TimeSpan? timeout = null) => new(
        new HttpClient(handler ?? new ScriptedJevHandler()),
        new JevOptions
        {
            ProviderName = ScriptedJevHandler.ProviderLabel,
            Model = Policy.EvaluatorModel,
            Timeout = timeout ?? Policy.EvaluatorTimeout,
            MaxEvidenceCharacters = Policy.MaxEvidenceCharacters,
            RetryBaseDelay = TimeSpan.FromMilliseconds(10),
        },
        FakeKey,
        Log);

    public RequirementsWorkflow Workflow(IRequirementEvaluator? evaluator = null, IChatClient? chat = null, ReviewPolicy? policy = null, IChatClient? reviewChat = null) => new(
        Client,
        Requirements,
        policy ?? Policy,
        new AgentRequirementReviewer(reviewChat ?? new ScriptedReviewChatClient(), evaluator ?? Jev()),
        new AgentProposalDrafter(chat ?? new ScriptedChatClient()),
        logger: Log);

    public Task<ReviewProposal> ProposeAsync(IRequirementEvaluator? evaluator = null, IChatClient? chat = null, ReviewPolicy? policy = null, IChatClient? reviewChat = null) =>
        Workflow(evaluator, chat, policy, reviewChat).ProposeAsync(ConnectionId, DocumentId);

    public Task<CommitOutcome> ApproveAsync(ReviewProposal proposal, string output = "method-statement.reviewed.docx", string reviewer = Reviewer) =>
        new AuthorizedCommitter(Client).CommitAsync(
            proposal,
            ReviewerApproval.For(proposal, reviewer, ReviewerDecision.Approved, DateTimeOffset.UtcNow),
            output);

    public string PartText(string fileName, string partPath)
    {
        using var zip = ZipFile.OpenRead(Path.Combine(Root, fileName));
        var entry = zip.GetEntry(partPath);
        if (entry is null)
        {
            return string.Empty;
        }

        using var reader = new StreamReader(entry.Open(), Encoding.UTF8);
        return reader.ReadToEnd();
    }

    public string DocumentXml(string fileName) => PartText(fileName, "word/document.xml");

    public string CommentsXml(string fileName) => PartText(fileName, "word/comments.xml");

    public bool Exists(string fileName) => File.Exists(Path.Combine(Root, fileName));

    public int DocxCount() => Directory.GetFiles(Root, "*.docx").Length;

    public string SourceSha256() => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(SourcePath))).ToLowerInvariant();

    public void Dispose()
    {
        _services.Dispose();
        try
        {
            Directory.Delete(Root, recursive: true);
        }
        catch (IOException)
        {
            // A locked temp directory is not a test failure.
        }
    }
}

/// <summary>Builds scripted decisions and responses.</summary>
internal static class Script
{
    public static RequirementDecision Decision(
        string requirementId,
        string evidence = "sufficient",
        double evidenceConfidence = 0.95,
        string result = "satisfied",
        double resultConfidence = 0.95,
        double materiality = 0.1,
        double materialityConfidence = 0.9) => new()
    {
        RequirementId = requirementId,
        Provider = "scripted",
        Model = "scripted",
        EvidenceStatus = new ChoiceJudgement(evidence, evidenceConfidence, new Dictionary<string, double>()),
        Result = new ChoiceJudgement(result, resultConfidence, new Dictionary<string, double>()),
        Materiality = new ScoreJudgement(materiality, materialityConfidence, new Dictionary<string, double>()),
    };

    /// <summary>An evaluator that returns the demo answer for each requirement.</summary>
    public static ScriptedRequirementEvaluator Demo() => new((r, _, _) => Task.FromResult(r.Id switch
    {
        "REQ-INS-01" => Decision(r.Id, result: "satisfied"),
        "REQ-REC-01" => Decision(r.Id, result: "violated", materiality: 2.1),
        _ => Decision(r.Id, evidence: "insufficient", result: "not_determined"),
    }));

    public static ScriptedRequirementEvaluator Always(Func<RegisteredRequirement, RequirementDecision> decide) =>
        new((r, _, _) => Task.FromResult(decide(r)));

    public static HttpResponseMessage Ok(string body) => ScriptedJevHandler.Json(HttpStatusCode.OK, body);

    public static string DemoBody(string requirementId) =>
        ScriptedAnswer.Demo[requirementId].ToResponse("jev-test").ToJsonString();
}

/// <summary>Captures log output so tests can check what reaches logs.</summary>
internal sealed class CapturingLogger : ILogger
{
    private readonly List<string> _lines = new();

    public IReadOnlyList<string> Lines
    {
        get
        {
            lock (_lines)
            {
                return _lines.ToList();
            }
        }
    }

    public IDisposable? BeginScope<TState>(TState state)
        where TState : notnull => null;

    public bool IsEnabled(LogLevel logLevel) => true;

    public void Log<TState>(LogLevel logLevel, EventId eventId, TState state, Exception? exception, Func<TState, Exception?, string> formatter)
    {
        lock (_lines)
        {
            _lines.Add($"{logLevel}: {formatter(state, exception)} {exception}");
        }
    }
}
