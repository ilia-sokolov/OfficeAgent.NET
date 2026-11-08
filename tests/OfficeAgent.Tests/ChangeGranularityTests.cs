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
using OfficeAgent.Mcp;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// <c>changeText</c> with granularity <c>Minimal</c> marks only the words that change (#28,
/// decided in discussion #38). Words are runs of non-whitespace compared exactly; the words
/// both sides share at the start and the end stay as they are.
/// </summary>
public class ChangeGranularityTests
{
    private const string Payment = "Payment is due within thirty days of invoice.";
    private const string Cap = "Liability is limited to the cap in clause 9.";
    private const string Notice = "Notice is due within thirty days unless extended in writing.";

    private static OfficeAgentClient Client() => new(new WordModule());

    // ── the narrowed edits ────────────────────────────────────────────────────

    [Fact]
    public void Only_the_changed_word_is_marked()
    {
        var document = Document(Plain(Payment));
        var op = Change(document, "thirty days", "forty-five days");

        var preview = Preview(document, op);
        var change = Assert.Single(preview.Changes);
        Assert.Equal("thirty", change.Before);
        Assert.Equal("forty-five", change.After);
        Assert.Null(change.Warning);

        var result = Commit(document, op);
        AssertValid(result);
        using var body = new OpenBody(result);
        var deleted = Assert.Single(body.Root.Descendants<DeletedRun>());
        var inserted = Assert.Single(body.Root.Descendants<InsertedRun>());
        Assert.Equal("thirty", deleted.InnerText);
        Assert.Equal("forty-five", inserted.InnerText);
        Assert.Equal("Payment is due within forty-five days of invoice.", VisibleText(body.Root));

        AssertResolves(result, accepted: "Payment is due within forty-five days of invoice.", rejected: Payment);
    }

    [Fact]
    public void Adding_words_in_front_is_a_pure_insertion()
    {
        var document = Document(Plain(Cap));
        var op = Change(document, "the cap", "125% of the cap");

        var change = Assert.Single(Preview(document, op).Changes);
        Assert.Equal(string.Empty, change.Before);
        Assert.Equal("125% of ", change.After);

        var result = Commit(document, op);
        AssertValid(result);
        using var body = new OpenBody(result);
        Assert.Empty(body.Root.Descendants<DeletedRun>());
        Assert.Equal("125% of ", Assert.Single(body.Root.Descendants<InsertedRun>()).InnerText);
        Assert.Equal("Liability is limited to 125% of the cap in clause 9.", VisibleText(body.Root));

        AssertResolves(result, accepted: "Liability is limited to 125% of the cap in clause 9.", rejected: Cap);
    }

    [Fact]
    public void Adding_words_after_a_shared_start_inserts_after_it()
    {
        var document = Document(Plain(Cap));
        var op = Change(document, "the cap in clause 9", "the cap in clause 9 as adjusted");

        var result = Commit(document, op);
        AssertValid(result);
        using var body = new OpenBody(result);
        Assert.Empty(body.Root.Descendants<DeletedRun>());
        Assert.Equal(" as adjusted", Assert.Single(body.Root.Descendants<InsertedRun>()).InnerText);
        AssertResolves(result, accepted: "Liability is limited to the cap in clause 9 as adjusted.", rejected: Cap);
    }

    [Fact]
    public void Dropping_words_is_a_pure_deletion()
    {
        var document = Document(Plain(Notice));
        var op = Change(document, "thirty days unless extended in writing", "thirty days in writing");

        var change = Assert.Single(Preview(document, op).Changes);
        Assert.Equal("unless extended ", change.Before);
        Assert.Equal(string.Empty, change.After);

        var result = Commit(document, op);
        AssertValid(result);
        using var body = new OpenBody(result);
        Assert.Equal("unless extended ", Assert.Single(body.Root.Descendants<DeletedRun>()).InnerText);
        Assert.Empty(body.Root.Descendants<InsertedRun>());
        AssertResolves(result, accepted: "Notice is due within thirty days in writing.", rejected: Notice);
    }

    [Theory]
    [InlineData("45 days", "forty-five days", "45", "forty-five")]
    [InlineData("thirty days", "Thirty days", "thirty", "Thirty")]
    [InlineData("thirty days", "thirty days.", "days", "days.")]
    [InlineData("the cap", "the cap and floor", "", " and floor")]
    [InlineData("alpha beta gamma", "alpha gamma", "beta ", "")]
    [InlineData("alpha  beta", "alpha beta", "  ", " ")]
    [InlineData("old text", "", "old text", "")]
    [InlineData("old text", "new words", "old text", "new words")]
    public void Words_are_whitespace_separated_and_compared_exactly(string expect, string with, string before, string after)
    {
        var document = Document(Plain($"Payment is due within {expect} of invoice"));
        var change = Assert.Single(Preview(document, Change(document, expect, with)).Changes);

        Assert.Equal(before, change.Before);
        Assert.Equal(after, change.After);
    }

    [Fact]
    public void Nothing_differing_marks_nothing()
    {
        var document = Document(Plain(Payment));
        var op = Change(document, "thirty days", "thirty days");

        var change = Assert.Single(Preview(document, op).Changes);
        Assert.Equal(string.Empty, change.Before);
        Assert.Equal(string.Empty, change.After);

        var result = Commit(document, op);
        AssertValid(result);
        using var body = new OpenBody(result);
        Assert.Empty(body.Root.Descendants<DeletedRun>());
        Assert.Empty(body.Root.Descendants<InsertedRun>());
        Assert.Equal(Payment, VisibleText(body.Root));
    }

    [Fact]
    public void Span_stays_the_default_and_marks_the_whole_match()
    {
        var document = Document(Plain(Payment));
        var op = new ChangeTextOp { Target = Span(document, "thirty days"), With = "forty-five days", Mode = ChangeMode.Tracked };

        Assert.Equal(ChangeGranularity.Span, op.Granularity);
        var result = Commit(document, op);
        AssertValid(result);
        using var body = new OpenBody(result);
        Assert.Equal("thirty days", Assert.Single(body.Root.Descendants<DeletedRun>()).InnerText);
        Assert.Equal("forty-five days", Assert.Single(body.Root.Descendants<InsertedRun>()).InnerText);
    }

    [Fact]
    public void A_direct_edit_ignores_granularity()
    {
        var document = Document(Plain(Payment));
        var minimal = Commit(document, Change(document, "thirty days", "forty-five days", ChangeMode.Direct));
        var span = Commit(document, new ChangeTextOp
        {
            Target = Span(document, "thirty days"), With = "forty-five days", Mode = ChangeMode.Direct
        });

        AssertValid(minimal);
        using var a = new OpenBody(minimal);
        using var b = new OpenBody(span);
        Assert.Equal(b.Root.OuterXml, a.Root.OuterXml);
        Assert.Empty(a.Root.Descendants<InsertedRun>());
    }

    // ── formatting ────────────────────────────────────────────────────────────

    [Fact]
    public void Words_to_delete_in_different_formatting_fall_back_to_span_with_a_warning()
    {
        // "thirty" is bold and " days" is not, so the narrowed deletion "thirty days" would
        // cross two formats.
        var document = Document(Plain("Payment is due within "), Bold("thirty"), Plain(" days of invoice."));
        var op = Change(document, "within thirty days of", "within forty-five weeks of");

        var change = Assert.Single(Preview(document, op).Changes);
        Assert.Equal("within thirty days of", change.Before);
        Assert.Equal("within forty-five weeks of", change.After);
        Assert.NotNull(change.Warning);
        Assert.Contains("fell back to Span", change.Warning);

        using var committed = Client().Commit(Handle(document), Plan(op));
        Assert.True(committed.Committed);
        Assert.Contains("fell back to Span", Assert.Single(committed.Report.Changes).Warning);
        var result = committed.ToBytes();
        AssertValid(result);
        using var body = new OpenBody(result);
        Assert.Equal("within thirty days of", Assert.Single(body.Root.Descendants<DeletedRun>()).InnerText);
        Assert.Equal("within forty-five weeks of", Assert.Single(body.Root.Descendants<InsertedRun>()).InnerText);
        AssertResolves(result, accepted: "Payment is due within forty-five weeks of invoice.",
            rejected: "Payment is due within thirty days of invoice.");
    }

    [Fact]
    public void Mixed_formatting_outside_the_changed_words_still_narrows()
    {
        // The match crosses bold and plain text, but the one word that changes is plain.
        var document = Document(Plain("Payment is due within "), Bold("thirty"), Plain(" days of invoice."));
        var op = Change(document, "within thirty days of", "within thirty weeks of");

        var change = Assert.Single(Preview(document, op).Changes);
        Assert.Equal("days", change.Before);
        Assert.Equal("weeks", change.After);
        Assert.Null(change.Warning);

        var result = Commit(document, op);
        AssertValid(result);
        using var body = new OpenBody(result);
        Assert.Equal("days", Assert.Single(body.Root.Descendants<DeletedRun>()).InnerText);
        var inserted = Assert.Single(body.Root.Descendants<InsertedRun>());
        Assert.Equal("weeks", inserted.InnerText);
        Assert.Null(inserted.Descendants<Bold>().FirstOrDefault());
        // The bold word is still bold and untouched.
        Assert.Contains(body.Root.Descendants<Run>(), run => run.InnerText == "thirty" && run.RunProperties?.Bold is not null
            && run.Ancestors<DeletedRun>().FirstOrDefault() is null);
    }

    [Fact]
    public void An_insertion_takes_the_formatting_of_the_word_before_it()
    {
        var document = Document(Plain("Liability is limited to "), Bold("the cap"), Plain(" in clause 9."));
        var result = Commit(document, Change(document, "the cap", "the cap amount"));

        AssertValid(result);
        using var body = new OpenBody(result);
        var inserted = Assert.Single(body.Root.Descendants<InsertedRun>());
        Assert.Equal(" amount", inserted.InnerText);
        Assert.NotNull(inserted.Descendants<Bold>().FirstOrDefault());
    }

    // ── through the agent tools ──────────────────────────────────────────────

    [Fact]
    public async Task The_wire_takes_granularity_and_reports_the_warning()
    {
        var tools = new OfficeAgentTools(Client());
        var document = Document(Plain("Payment is due within "), Bold("thirty"), Plain(" days of invoice."));
        var json = await tools.EditDocumentContent(Convert.ToBase64String(document), """
            [ { "op": "changeText", "target": { "find": "within thirty days of" },
                "with": "within forty-five weeks of", "mode": "Tracked", "granularity": "minimal" } ]
            """, preview: true);

        using var parsed = JsonDocument.Parse(json);
        Assert.True(parsed.RootElement.GetProperty("isValid").GetBoolean(), json);
        var change = parsed.RootElement.GetProperty("changes")[0];
        Assert.Contains("fell back to Span", change.GetProperty("warning").GetString());
    }

    [Fact]
    public async Task Every_match_is_narrowed_on_its_own()
    {
        var tools = new OfficeAgentTools(Client());
        var document = Document(
            new[] { Plain("Payment is due within thirty days of invoice.") },
            new[] { Plain("Disputes are raised within thirty days of receipt.") });
        var json = await tools.EditDocumentContent(Convert.ToBase64String(document), """
            [ { "op": "changeText", "target": { "find": "thirty days", "match": "all", "wholeWord": true },
                "with": "forty-five days", "mode": "Tracked", "granularity": "Minimal" } ]
            """);

        using var parsed = JsonDocument.Parse(json);
        Assert.True(parsed.RootElement.GetProperty("committed").GetBoolean(), json);
        var changes = parsed.RootElement.GetProperty("changes").EnumerateArray().ToList();
        Assert.Equal(2, changes.Count);
        Assert.All(changes, c =>
        {
            Assert.Equal("thirty", c.GetProperty("before").GetString());
            Assert.Equal("forty-five", c.GetProperty("after").GetString());
            Assert.Equal(JsonValueKind.Null, c.GetProperty("warning").ValueKind);
        });

        var result = Convert.FromBase64String(parsed.RootElement.GetProperty("contentBase64").GetString()!);
        AssertValid(result);
        using var body = new OpenBody(result);
        Assert.Equal(new[] { "thirty", "thirty" }, body.Root.Descendants<DeletedRun>().Select(d => d.InnerText));
        Assert.Equal(new[] { "forty-five", "forty-five" }, body.Root.Descendants<InsertedRun>().Select(i => i.InnerText));
    }

    // ── the connection default ────────────────────────────────────────────────

    [Fact]
    public void Span_is_the_connection_default_unless_configured()
    {
        using var plain = new Workspace(null);
        using var minimal = new Workspace(ChangeGranularity.Minimal);

        Assert.Equal(ChangeGranularity.Span, plain.Client.DefaultChangeGranularityFor("workspace"));
        Assert.Equal(ChangeGranularity.Minimal, minimal.Client.DefaultChangeGranularityFor("workspace"));
        Assert.Equal(ChangeGranularity.Span, minimal.Client.DefaultChangeGranularityFor("no-such-connection"));
        Assert.Equal(ChangeGranularity.Span, minimal.Client.DefaultChangeGranularityFor(""));
    }

    [Theory]
    [InlineData(null, null, "thirty days")]
    [InlineData(ChangeGranularity.Minimal, null, "thirty")]
    [InlineData(ChangeGranularity.Minimal, "Span", "thirty days")]
    [InlineData(null, "Minimal", "thirty")]
    public async Task An_omitted_granularity_follows_the_connection_and_an_explicit_one_wins(
        ChangeGranularity? connectionDefault, string? planGranularity, string deleted)
    {
        using var workspace = new Workspace(connectionDefault);
        var path = Path.Combine(workspace.Root, "contract.docx");
        await File.WriteAllBytesAsync(path, Document(Plain(Payment)));
        var granularity = planGranularity is null ? "" : $", \"granularity\": \"{planGranularity}\"";

        var json = await new OfficeAgentTools(workspace.Client).EditDocument("workspace", "contract.docx", $$"""
            [ { "op": "changeText", "target": { "find": "thirty days" }, "with": "forty-five days"{{granularity}} } ]
            """);

        using var parsed = JsonDocument.Parse(json);
        Assert.True(parsed.RootElement.GetProperty("committed").GetBoolean(), json);
        var result = await File.ReadAllBytesAsync(path);
        AssertValid(result);
        using var body = new OpenBody(result);
        Assert.Equal(deleted, Assert.Single(body.Root.Descendants<DeletedRun>()).InnerText);
    }

    [Fact]
    public void The_mcp_server_reads_the_connection_default_and_names_a_bad_value()
    {
        var root = Path.Combine(Path.GetTempPath(), $"officeagent-granularity-{Guid.NewGuid():N}");
        Directory.CreateDirectory(root);
        try
        {
            Assert.Equal("Span", new FileSystemConnectionOptions().DefaultChangeGranularity);

            var services = new ServiceCollection();
            services.AddOfficeAgentMcp(new OfficeAgentMcpOptions
            {
                FileSystemConnections =
                {
                    new FileSystemConnectionOptions { ConnectionId = "documents", RootPath = root, DefaultChangeGranularity = "minimal" }
                }
            });
            using var provider = services.BuildServiceProvider();
            Assert.Equal(ChangeGranularity.Minimal,
                provider.GetRequiredService<OfficeAgentClient>().DefaultChangeGranularityFor("documents"));

            var refused = Assert.Throws<InvalidOperationException>(() => new ServiceCollection().AddOfficeAgentMcp(new OfficeAgentMcpOptions
            {
                FileSystemConnections =
                {
                    new FileSystemConnectionOptions { ConnectionId = "documents", RootPath = root, DefaultChangeGranularity = "Minimum" }
                }
            }));
            Assert.Contains("Expected Span or Minimal", refused.Message);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    // ── helpers ───────────────────────────────────────────────────────────────

    private static ChangeTextOp Change(byte[] document, string expect, string with, ChangeMode mode = ChangeMode.Tracked) => new()
    {
        Target = Span(document, expect),
        With = with,
        Mode = mode,
        Granularity = ChangeGranularity.Minimal
    };

    private static TextSpanAnchor Span(byte[] document, string text) =>
        (TextSpanAnchor)Client().Find(Handle(document), new FindQuery(text) { Options = new MatchOptions { CaseSensitive = true } })
            .First().Anchor;

    private static StreamHandle Handle(byte[] document) => new(new MemoryStream(document, writable: false), "input.docx");

    private static DocumentPlan Plan(params PlanOperation[] operations) => new()
    {
        Revision = new RevisionMetadata { Author = "Legal Ops" },
        Operations = operations
    };

    private static ChangeReport Preview(byte[] document, PlanOperation op)
    {
        using var previewed = Client().Apply(Handle(document), Plan(op), ApplyOptions.Preview);
        Assert.True(previewed.Report.IsValid, string.Join("; ", previewed.Report.Errors.Select(e => $"{e.Code}: {e.Message}")));
        return previewed.Report;
    }

    private static byte[] Commit(byte[] document, params PlanOperation[] operations)
    {
        using var applied = Client().Commit(Handle(document), Plan(operations));
        Assert.True(applied.Committed, string.Join("; ", applied.Report.Errors.Select(e => $"{e.Code}: {e.Message}")));
        return applied.ToBytes();
    }

    /// <summary>Accepting gives the edited text and rejecting the original, exactly.</summary>
    private static void AssertResolves(byte[] result, string accepted, string rejected)
    {
        foreach (var (action, expected) in new[] { (RevisionAction.Accept, accepted), (RevisionAction.Reject, rejected) })
        {
            var resolved = Commit(result, new RevisionOp { Target = new NodeAnchor { Kind = "revision", Path = "all" }, Action = action });
            AssertValid(resolved);
            using var body = new OpenBody(resolved);
            Assert.Empty(body.Root.Descendants<DeletedRun>());
            Assert.Empty(body.Root.Descendants<InsertedRun>());
            Assert.Equal(expected, VisibleText(body.Root));
        }
    }

    private static Run Plain(string text) => new(new Text(text) { Space = SpaceProcessingModeValues.Preserve });

    private static Run Bold(string text) => new(new RunProperties(new Bold()), new Text(text) { Space = SpaceProcessingModeValues.Preserve });

    private static byte[] Document(params Run[] runs) => Document(new[] { runs });

    private static byte[] Document(params Run[][] paragraphs)
    {
        var buffer = new MemoryStream();
        using (var package = WordprocessingDocument.Create(buffer, WordprocessingDocumentType.Document))
        {
            var main = package.AddMainDocumentPart();
            main.Document = new Document(new Body(paragraphs.Select((runs, i) =>
            {
                var paragraph = new Paragraph { ParagraphId = (0x10 + i).ToString("X8") };
                paragraph.Append(runs);
                return paragraph;
            })));
        }
        return buffer.ToArray();
    }

    /// <summary>The text a reader sees: run text, without struck-through deletions.</summary>
    private static string VisibleText(OpenXmlElement root) =>
        string.Join("\n", root.Descendants<Paragraph>().Select(p => string.Concat(p.Descendants<Text>().Select(t => t.Text))));

    private sealed class OpenBody : IDisposable
    {
        private readonly MemoryStream _stream;
        private readonly WordprocessingDocument _document;

        public Body Root { get; }

        public OpenBody(byte[] bytes)
        {
            _stream = new MemoryStream(bytes);
            _document = WordprocessingDocument.Open(_stream, isEditable: false);
            Root = _document.MainDocumentPart!.Document.Body!;
        }

        public void Dispose()
        {
            _document.Dispose();
            _stream.Dispose();
        }
    }

    private static void AssertValid(byte[] document)
    {
        using var stream = new MemoryStream(document);
        using var opened = WordprocessingDocument.Open(stream, isEditable: false);
        var problems = new OpenXmlValidator(FileFormatVersions.Office2019)
            .Validate(opened)
            .Select(e => $"{e.Path?.XPath}: {e.Description}")
            .ToList();
        Assert.True(problems.Count == 0, string.Join("; ", problems.Take(3)));
    }

    private sealed class Workspace : IDisposable
    {
        private readonly ServiceProvider _services;

        public Workspace(ChangeGranularity? defaultGranularity)
        {
            Root = Path.Combine(Path.GetTempPath(), $"officeagent-granularity-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Root);
            var services = new ServiceCollection();
            services.AddWordFormat();
            services.AddFileSystemDocumentProvider("workspace", Root, o =>
            {
                // Leaving it unset is the case that must still mean Span.
                if (defaultGranularity is { } granularity) o.DefaultChangeGranularity = granularity;
            });
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
