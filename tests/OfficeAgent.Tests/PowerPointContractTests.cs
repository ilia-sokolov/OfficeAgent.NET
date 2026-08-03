using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.Abstractions;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.PowerPoint;
using OfficeAgent.Word;
using DocFormat = OfficeAgent.Abstractions.DocumentFormat;

namespace OfficeAgent.Tests;

/// <summary>
/// The contract where the two format modules meet: a shared verb vocabulary in which
/// each module implements a subset, divergences are refused with something the agent can
/// act on, and one client serves both formats.
/// </summary>
public class PowerPointContractTests
{
    [Fact]
    public void Adding_Resolve_to_the_comment_verb_does_not_change_what_Word_accepts()
    {
        var client = new OfficeAgentClient(new WordModule());

        // CommentAction gained Resolve for PowerPoint. Word has no resolved state, so it
        // must keep refusing rather than quietly adding a comment instead.
        var report = client.Preview(
            new StreamHandle(new MemoryStream(DocxFactory.Contract())),
            new DocumentPlan
            {
                Operations = new PlanOperation[]
                {
                    new CommentOp
                    {
                        Target = new TextSpanAnchor { ParaId = "w14:00000002", Expect = "Acme Corp" },
                        Text = "Check this.",
                        Action = CommentAction.Resolve
                    }
                }
            });

        Assert.False(report.IsValid);
        Assert.Equal(ValidationErrorCodes.UnsupportedOperation, Assert.Single(report.Errors).Code);
    }

    [Fact]
    public void Word_still_adds_comments_exactly_as_before()
    {
        var client = new OfficeAgentClient(new WordModule());
        var hits = client.Find(
            new StreamHandle(new MemoryStream(DocxFactory.Contract())),
            new FindQuery { Pattern = "Acme Corp" });

        var report = client.Preview(
            new StreamHandle(new MemoryStream(DocxFactory.Contract())),
            new DocumentPlan
            {
                Operations = new PlanOperation[]
                {
                    new CommentOp { Target = hits[0].Anchor!, Text = "Check this." }
                }
            });

        Assert.True(report.IsValid,
            string.Join("; ", report.Errors.Select(e => $"{e.Code}: {e.Message}")));
    }

    [Fact]
    public void Tracked_mode_is_refused_on_a_deck_with_an_actionable_message()
    {
        var client = new OfficeAgentClient(new PowerPointModule());
        var hits = client.Find(
            new StreamHandle(new MemoryStream(PptxFactory.Deck())),
            new FindQuery { Pattern = PptxFactory.TitleText });

        var report = client.Preview(
            new StreamHandle(new MemoryStream(PptxFactory.Deck())),
            new DocumentPlan
            {
                Format = DocFormat.PowerPoint,
                Operations = new PlanOperation[]
                {
                    new ChangeTextOp
                    {
                        Target = hits[0].Anchor!,
                        With = "Annual Review",
                        Mode = ChangeMode.Tracked
                    }
                }
            });

        var error = Assert.Single(report.Errors);
        Assert.Equal(ValidationErrorCodes.InvalidOperation, error.Code);
        // The agent must be told what to do instead, not merely that it failed.
        Assert.Contains("Direct", error.Message);
    }

    [Fact]
    public void Verbs_the_deck_module_does_not_implement_are_named_not_silently_skipped()
    {
        var client = new OfficeAgentClient(new PowerPointModule());
        var hits = client.Find(
            new StreamHandle(new MemoryStream(PptxFactory.Deck())),
            new FindQuery { Pattern = PptxFactory.TitleText });

        var report = client.Preview(
            new StreamHandle(new MemoryStream(PptxFactory.Deck())),
            new DocumentPlan
            {
                Format = DocFormat.PowerPoint,
                Operations = new PlanOperation[]
                {
                    new SetPropertyOp { Target = new NodeAnchor { Kind = "docProperty", Path = "core/title" }, Value = "Nope" }
                }
            });

        Assert.False(report.IsValid);
        Assert.Equal(ValidationErrorCodes.UnsupportedOperation, Assert.Single(report.Errors).Code);
    }

    [Fact]
    public void One_client_routes_each_document_to_the_module_that_handles_it()
    {
        var client = new OfficeAgentClient(new WordModule(), new PowerPointModule());

        Assert.Equal(DocFormat.Word, client.Inspect(DocxFactory.Contract()).Format);
        Assert.Equal(DocFormat.PowerPoint, client.Inspect(PptxFactory.Deck()).Format);
    }

    [Fact]
    public async Task A_deck_reports_the_presentation_media_type_not_a_generic_binary()
    {
        using var workspace = new ContentTypeWorkspace();

        var deck = await workspace.Client.CreateAsync("workspace", "review.pptx");
        var document = await workspace.Client.CreateAsync("workspace", "notes.docx");

        // Hosts serve downloads with this value: a deck arriving as octet-stream prompts
        // a save-as for an unnamed binary instead of opening in PowerPoint.
        Assert.Equal(
            "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            deck.Document!.ContentType);
        Assert.Equal(
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            document.Document!.ContentType);
    }

    [Fact]
    public void A_word_handler_contributed_to_the_container_is_not_offered_to_a_deck()
    {
        var services = new ServiceCollection();
        services.AddWordFormat();
        services.AddPowerPointFormat();
        // A host extending Word with a verb PowerPoint does not implement. Registered as
        // a bare IOperationHandler, which is the only thing AddWordFormat consumes.
        services.AddSingleton<IOperationHandler, WordOnlySetPropertyHandler>();
        services.AddOfficeAgent();

        using var provider = services.BuildServiceProvider();
        var client = provider.GetRequiredService<OfficeAgentClient>();

        var hits = client.Find(
            new StreamHandle(new MemoryStream(PptxFactory.Deck())),
            new FindQuery { Pattern = PptxFactory.TitleText });

        var report = client.Preview(
            new StreamHandle(new MemoryStream(PptxFactory.Deck())),
            new DocumentPlan
            {
                Format = DocFormat.PowerPoint,
                Operations = new PlanOperation[] { new SetPropertyOp { Target = new NodeAnchor { Kind = "docProperty", Path = "core/title" }, Value = "Nope" } }
            });

        // The deck must still refuse the verb: the Word handler leaking into the
        // PowerPoint module would "succeed" by writing WordprocessingML into a slide.
        Assert.Equal(ValidationErrorCodes.UnsupportedOperation, Assert.Single(report.Errors).Code);
    }

    /// <summary>Stands in for a host-contributed Word extension handler for a verb the PowerPoint module does not implement.</summary>
    private sealed class WordOnlySetPropertyHandler : IOperationHandler
    {
        public bool CanHandle(PlanOperation operation) => operation is SetPropertyOp;

        public OperationPreview Preview(ApplyContext context, PlanOperation operation) =>
            OperationPreview.Ok(new ProposedChange { Verb = "setProperty", Target = operation.Target });

        public void Apply(ApplyContext context, PlanOperation operation)
        {
        }
    }

    private sealed class ContentTypeWorkspace : IDisposable
    {
        private readonly ServiceProvider _services;

        public ContentTypeWorkspace()
        {
            Root = Path.Combine(Path.GetTempPath(), $"officeagent-ct-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Root);

            var services = new ServiceCollection();
            services.AddWordFormat();
            services.AddPowerPointFormat();
            services.AddFileSystemDocumentProvider("workspace", Root, o =>
                o.AllowedExtensions = new[] { ".docx", ".pptx" });
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

    [Fact]
    public void The_agent_surface_describes_both_formats()
    {
        // An agent told to default to Tracked would have every deck edit refused, and one
        // told the starting anchor is always auto-0000 would mis-target a new deck.
        Assert.Contains("PowerPoint", OfficeAgentTools.SystemPromptGuidance);
        Assert.Contains("Direct", OfficeAgentTools.SystemPromptGuidance);
        Assert.Contains("slide256/shape2/p0", OfficeAgentTools.CreationPromptGuidance);

        var create = new OfficeAgentTools(new OfficeAgentClient(new PowerPointModule()))
            .AsAIFunctions(new OfficeAgentToolsOptions { AllowCreation = true })
            .Single(f => f.Name == "create_document");
        Assert.Contains(".pptx", create.Description);
    }
}
