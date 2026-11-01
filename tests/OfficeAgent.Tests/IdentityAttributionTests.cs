using System.Net;
using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Excel;
using OfficeAgent.Mcp;
using OfficeAgent.PowerPoint;
using OfficeAgent.SharePoint;
using OfficeAgent.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;
using DocFormat = OfficeAgent.Abstractions.DocumentFormat;

namespace OfficeAgent.Tests;

/// <summary>
/// Who a hosted edit is attributed to. The actor comes only from a token Entra has accepted,
/// never from a header alone or from plan JSON, and it reaches the document only by opt-in.
/// </summary>
public class IdentityAttributionTests
{
    private static string Jwt(object claims) =>
        "e30." +
        Convert.ToBase64String(Encoding.UTF8.GetBytes(JsonSerializer.Serialize(claims)))
            .TrimEnd('=').Replace('+', '-').Replace('/', '_') +
        ".signature";

    private static readonly string UserToken = Jwt(new
    {
        oid = "00000000-aaaa-bbbb-cccc-000000000001",
        sub = "pairwise-subject",
        iss = "https://login.microsoftonline.com/tenant/v2.0",
        name = "Alice Smith"
    });

    private static OnBehalfOfOptions Options() => new()
    {
        TenantId = "00000000-0000-0000-0000-000000000000",
        ClientId = "api-client-id",
        ClientSecret = "api-secret",
        Authority = "https://login.fake"
    };

    private static OnBehalfOfAccessTokenProvider Exchange(HttpStatusCode status = HttpStatusCode.OK) =>
        new(Options(), new HttpClient(new FakeEntra(status)), () => GraphUserContext.CurrentUserAccessToken);

    // ── the actor ─────────────────────────────────────────────────────────

    [Fact]
    public async Task An_exchange_inside_a_nested_call_identifies_the_user_to_the_request()
    {
        using (GraphUserContext.Push(UserToken))
        {
            Assert.Null(new OnBehalfOfActorProvider().GetCurrentActor());

            // The exchange runs deep in an async call, as it does inside a provider read; a
            // plain async-local value set there would be gone by the time the receipt is written.
            await Task.Run(() => Exchange().GetAccessTokenAsync());

            var actor = new OnBehalfOfActorProvider().GetCurrentActor();
            Assert.NotNull(actor);
            Assert.Equal("00000000-aaaa-bbbb-cccc-000000000001", actor!.Subject);
            Assert.Equal("https://login.microsoftonline.com/tenant/v2.0", actor.Issuer);
            Assert.Equal("Alice Smith", actor.DisplayName);
        }

        Assert.Null(new OnBehalfOfActorProvider().GetCurrentActor());
    }

    [Fact]
    public async Task A_cached_exchange_still_identifies_the_user()
    {
        var tokens = Exchange();
        using (GraphUserContext.Push(UserToken)) await tokens.GetAccessTokenAsync();

        using (GraphUserContext.Push(UserToken))
        {
            await tokens.GetAccessTokenAsync();
            Assert.NotNull(new OnBehalfOfActorProvider().GetCurrentActor());
        }
    }

    [Fact]
    public async Task A_refused_exchange_gives_no_actor()
    {
        using (GraphUserContext.Push(UserToken))
        {
            await Assert.ThrowsAnyAsync<Exception>(() =>
                Exchange(HttpStatusCode.BadRequest).GetAccessTokenAsync());

            Assert.Null(new OnBehalfOfActorProvider().GetCurrentActor());
        }
    }

    [Fact]
    public void A_header_with_no_exchange_gives_no_actor()
    {
        using (GraphUserContext.Push(UserToken))
            Assert.Null(new OnBehalfOfActorProvider().GetCurrentActor());
    }

    [Theory]
    [InlineData("not-a-jwt")]
    [InlineData("a.!!!.c")]
    public async Task An_unreadable_token_gives_no_actor(string token)
    {
        using (GraphUserContext.Push(token))
        {
            await Exchange().GetAccessTokenAsync();
            Assert.Null(new OnBehalfOfActorProvider().GetCurrentActor());
        }
    }

    [Fact]
    public async Task A_token_without_an_object_id_gives_no_actor()
    {
        // sub is per application, so it is not the stable identity a receipt promises.
        using (GraphUserContext.Push(Jwt(new { sub = "pairwise-subject", name = "Alice Smith" })))
        {
            await Exchange().GetAccessTokenAsync();
            Assert.Null(new OnBehalfOfActorProvider().GetCurrentActor());
        }
    }

    // ── the author ────────────────────────────────────────────────────────

    [Fact]
    public async Task By_default_the_user_is_on_the_receipt_and_not_in_the_document()
    {
        var (bytes, receipt) = await CommitAsUser(new OfficeAgentClientOptions());

        Assert.Equal("Alice Smith", receipt.Actor!.DisplayName);
        Assert.Equal(new[] { "Agent Name" }, RevisionAuthors(bytes));
        Assert.Equal(new[] { "Agent Name" }, CommentAuthors(bytes));
    }

    [Fact]
    public async Task With_the_setting_on_the_user_authors_the_document_and_the_plan_stays_on_the_receipt()
    {
        var (bytes, receipt) = await CommitAsUser(new OfficeAgentClientOptions { RevisionAuthorFromActor = true });

        Assert.Equal(new[] { "Alice Smith" }, RevisionAuthors(bytes));
        Assert.Equal(new[] { "Alice Smith" }, CommentAuthors(bytes));
        Assert.Equal("AS", CommentInitials(bytes));
        Assert.Equal("Agent Name", receipt.Revision.Author);
        Assert.Equal("Alice Smith", receipt.Actor!.DisplayName);
    }

    [Fact]
    public async Task With_the_setting_on_and_no_user_nothing_changes()
    {
        var (bytes, receipt) = await CommitAsUser(
            new OfficeAgentClientOptions { RevisionAuthorFromActor = true }, actor: null);

        Assert.Null(receipt.Actor);
        Assert.Equal(new[] { "Agent Name" }, RevisionAuthors(bytes));
    }

    [Fact]
    public async Task A_name_a_revision_cannot_carry_leaves_the_plan_author()
    {
        var (bytes, _) = await CommitAsUser(
            new OfficeAgentClientOptions { RevisionAuthorFromActor = true },
            new AuditActor { Subject = "oid", DisplayName = "Alice\nSmith" });

        Assert.Equal(new[] { "Agent Name" }, RevisionAuthors(bytes));
    }

    [Fact]
    public void A_trusted_author_reaches_workbook_and_deck_comments()
    {
        var options = new ApplyOptions { DryRun = false, RevisionAuthor = "Alice Smith" };

        var workbook = new OfficeAgentClient(new ExcelModule());
        using (var applied = workbook.Apply(
                   new StreamHandle(new MemoryStream(new ExcelModule().CreateBlank()), "c.xlsx"),
                   new DocumentPlan
                   {
                       Format = DocFormat.Excel,
                       Operations = new PlanOperation[]
                       {
                           new CommentOp { Target = new CellAnchor { SheetId = 1, Address = "A1" }, Text = "Check" }
                       }
                   },
                   options))
        {
            Assert.True(applied.Committed, string.Join("; ", applied.Report.Errors.Select(e => e.Message)));
            using var document = SpreadsheetDocument.Open(new MemoryStream(applied.ToBytes()), false);
            var comments = document.WorkbookPart!.WorksheetParts.Single().WorksheetCommentsPart!.Comments!;
            Assert.Contains("Alice Smith", comments.InnerText);
        }

        var deck = new OfficeAgentClient(new PowerPointModule());
        using (var applied = deck.Apply(
                   new StreamHandle(new MemoryStream(PptxFactory.Deck()), "c.pptx"),
                   new DocumentPlan
                   {
                       Format = DocFormat.PowerPoint,
                       Operations = new PlanOperation[]
                       {
                           new CommentOp { Target = new NodeAnchor { Kind = "slide", Path = "slide#256" }, Text = "Check" }
                       }
                   },
                   options))
        {
            Assert.True(applied.Committed, string.Join("; ", applied.Report.Errors.Select(e => e.Message)));
            using var document = PresentationDocument.Open(new MemoryStream(applied.ToBytes()), false);
            var author = document.PresentationPart!.GetPartsOfType<PowerPointAuthorsPart>().Single();
            Assert.Contains("Alice Smith", author.AuthorList!.OuterXml);
        }
    }

    // ── the server ────────────────────────────────────────────────────────

    [Fact]
    public void The_server_binds_the_setting_and_registers_the_actor_provider()
    {
        var configuration = new ConfigurationBuilder()
            .AddInMemoryCollection(new Dictionary<string, string?>
            {
                ["OfficeAgent:RevisionAuthorFromActor"] = "true"
            })
            .Build();
        var options = OfficeAgentConfiguration.Bind(configuration);
        Assert.True(options.RevisionAuthorFromActor);
        Assert.False(new OfficeAgentMcpOptions().RevisionAuthorFromActor);

        var services = new ServiceCollection();
        services.AddOfficeAgentMcp(options);
        using var provider = services.BuildServiceProvider();

        Assert.IsType<OnBehalfOfActorProvider>(provider.GetRequiredService<IAuditActorProvider>());
        Assert.True(provider.GetRequiredService<OfficeAgentClientOptions>().RevisionAuthorFromActor);
    }

    // ── helpers ───────────────────────────────────────────────────────────

    private static async Task<(byte[] Bytes, ApplyReceipt Receipt)> CommitAsUser(
        OfficeAgentClientOptions options, AuditActor? actor) =>
        await Commit(options, actor);

    private static Task<(byte[] Bytes, ApplyReceipt Receipt)> CommitAsUser(OfficeAgentClientOptions options) =>
        Commit(options, new AuditActor { Subject = "oid", Issuer = "tenant", DisplayName = "Alice Smith" });

    private static async Task<(byte[] Bytes, ApplyReceipt Receipt)> Commit(OfficeAgentClientOptions options, AuditActor? actor)
    {
        var root = Path.Combine(Path.GetTempPath(), $"officeagent-identity-{Guid.NewGuid():N}");
        Directory.CreateDirectory(root);
        try
        {
            var services = new ServiceCollection();
            services.AddWordFormat();
            services.AddFileSystemDocumentProvider("workspace", root);
            services.AddSingleton<IAuditActorProvider>(new FixedActor(actor));
            services.AddSingleton(options);
            services.AddOfficeAgent();
            using var provider = services.BuildServiceProvider();
            var client = provider.GetRequiredService<OfficeAgentClient>();

            var reference = await client.RegisterBytesAsync("workspace", root, DocxFactory.Contract(), "contract.docx");
            var paragraph = (await client.InspectAsync(reference)).Paragraphs.First(p => p.Text.Contains("Acme Corp"));
            var target = new TextSpanAnchor { ParaId = paragraph.ParaId, Expect = "Acme Corp", Occurrence = 0 };

            var result = await client.CommitAsync(reference, new DocumentPlan
            {
                Revision = new RevisionMetadata { Author = "Agent Name" },
                Operations = new PlanOperation[]
                {
                    new ChangeTextOp { Target = target, With = "Globex", Mode = ChangeMode.Tracked },
                    new CommentOp
                    {
                        Target = new TextSpanAnchor { ParaId = paragraph.ParaId, Expect = "services", Occurrence = 0 },
                        Text = "Check",
                        Author = "Agent Name"
                    }
                }
            });

            Assert.True(result.Committed, string.Join("; ", result.Report.Errors.Select(e => e.Message)));
            using var saved = await client.OpenReadAsync(result.Document!);
            using var copy = new MemoryStream();
            await saved.Stream.CopyToAsync(copy);
            return (copy.ToArray(), result.Receipt!);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    private static string[] RevisionAuthors(byte[] bytes)
    {
        using var document = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        var body = document.MainDocumentPart!.Document!.Body!;
        return body.Descendants<W.InsertedRun>().Select(r => r.Author!.Value!)
            .Concat(body.Descendants<W.DeletedRun>().Select(r => r.Author!.Value!))
            .Distinct()
            .ToArray();
    }

    private static string[] CommentAuthors(byte[] bytes)
    {
        using var document = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        return document.MainDocumentPart!.WordprocessingCommentsPart!.Comments!
            .Elements<W.Comment>().Select(c => c.Author!.Value!).Distinct().ToArray();
    }

    private static string CommentInitials(byte[] bytes)
    {
        using var document = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        return document.MainDocumentPart!.WordprocessingCommentsPart!.Comments!
            .Elements<W.Comment>().Single().Initials!.Value!;
    }

    private sealed class FixedActor : IAuditActorProvider
    {
        private readonly AuditActor? _actor;
        public FixedActor(AuditActor? actor) => _actor = actor;
        public AuditActor? GetCurrentActor() => _actor;
    }

    private sealed class FakeEntra : HttpMessageHandler
    {
        private readonly HttpStatusCode _status;
        public FakeEntra(HttpStatusCode status) => _status = status;

        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken token)
        {
            var body = _status == HttpStatusCode.OK
                ? """{ "access_token": "graph", "expires_in": 3600 }"""
                : """{ "error": "invalid_grant", "error_description": "refused" }""";
            return Task.FromResult(new HttpResponseMessage(_status)
            {
                Content = new StringContent(body, Encoding.UTF8, "application/json")
            });
        }
    }
}
