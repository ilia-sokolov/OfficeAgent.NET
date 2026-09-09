using System.IO.Compression;
using System.Text;
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Samples.ContractReview;
using OfficeAgent.Word;

namespace ContractReview.Tests;

/// <summary>
/// A disposable review environment: a temp-rooted filesystem connection with the
/// contract fixture registered, plus helpers for asserting on the produced .docx.
/// </summary>
internal sealed class ReviewHarness : IDisposable
{
    public const string ConnectionId = "contracts";

    private readonly ServiceProvider _services;

    private ReviewHarness(ServiceProvider services, string root, string contractPath, string documentId)
    {
        _services = services;
        Root = root;
        ContractPath = contractPath;
        DocumentId = documentId;
        Client = services.GetRequiredService<OfficeAgentClient>();
    }

    public string Root { get; }

    public string ContractPath { get; }

    public string DocumentId { get; }

    public OfficeAgentClient Client { get; }

    public static async Task<ReviewHarness> CreateAsync(byte[]? contract = null, string fileName = "contract.docx")
    {
        var root = Path.Combine(Path.GetTempPath(), "oa-review-tests-" + Guid.NewGuid().ToString("N")[..10]);
        Directory.CreateDirectory(root);

        var path = Path.Combine(root, fileName);
        File.WriteAllBytes(path, contract ?? ContractFixture.Create());

        var services = new ServiceCollection()
            .AddWordFormat()
            .AddFileSystemDocumentProvider(ConnectionId, root, o => o.AllowedExtensions = new[] { ".docx" })
            .AddOfficeAgent()
            .BuildServiceProvider();

        var client = services.GetRequiredService<OfficeAgentClient>();
        var reference = await client.RegisterAsync(ConnectionId, path);

        return new ReviewHarness(services, root, path, reference.ItemId);
    }

    /// <summary>Reads a part of a produced .docx as text.</summary>
    public string PartText(string outputName, string partPath)
    {
        var path = Path.Combine(Root, outputName);
        using var zip = ZipFile.OpenRead(path);
        var entry = zip.GetEntry(partPath);
        if (entry is null)
        {
            return string.Empty;
        }

        using var stream = entry.Open();
        using var reader = new StreamReader(stream, Encoding.UTF8);
        return reader.ReadToEnd();
    }

    public string DocumentXml(string outputName) => PartText(outputName, "word/document.xml");

    public string CommentsXml(string outputName) => PartText(outputName, "word/comments.xml");

    public bool OutputExists(string outputName) => File.Exists(Path.Combine(Root, outputName));

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

/// <summary>Playbooks used by the tests.</summary>
internal static class TestPlaybooks
{
    public static Playbook PaymentRedline() => Playbook.FromJson(
        """
        {
          "name": "Test",
          "rules": [
            {
              "id": "PAY-01",
              "title": "Payment terms",
              "severity": "high",
              "patterns": ["60 days"],
              "regex": false,
              "rationale": "Payment must be within 30 days.",
              "action": "redline"
            }
          ]
        }
        """);

    public static Playbook TwoRulesOneSpan() => Playbook.FromJson(
        """
        {
          "name": "Overlapping",
          "rules": [
            {
              "id": "LOW-01",
              "title": "Low severity on the same span",
              "severity": "low",
              "patterns": ["60 days"],
              "regex": false,
              "rationale": "Low rule.",
              "action": "redline"
            },
            {
              "id": "HIGH-01",
              "title": "High severity on the same span",
              "severity": "high",
              "patterns": ["60 days"],
              "regex": false,
              "rationale": "High rule.",
              "action": "redline"
            }
          ]
        }
        """);

    public static Playbook CommentOnly() => Playbook.FromJson(
        """
        {
          "name": "Comment only",
          "rules": [
            {
              "id": "REN-01",
              "title": "Automatic renewal",
              "severity": "medium",
              "patterns": ["automatically renew"],
              "regex": false,
              "rationale": "Automatic renewal needs 60 days notice.",
              "action": "commentOnly"
            }
          ]
        }
        """);

    public static Playbook Undetectable() => Playbook.FromJson(
        """
        {
          "name": "Undetectable",
          "rules": [
            {
              "id": "NONE-01",
              "title": "Matches nothing",
              "severity": "medium",
              "patterns": ["force majeure clause that is not present"],
              "regex": false,
              "rationale": "Never matches.",
              "action": "commentOnly"
            }
          ]
        }
        """);
}
