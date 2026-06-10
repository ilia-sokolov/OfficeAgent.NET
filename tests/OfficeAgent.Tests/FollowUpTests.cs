using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.Abstractions;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

public class FollowUpTests
{
    private static OfficeAgentClient Office() => new(new WordModule());
    private static StreamHandle Handle(byte[] bytes) => new(new MemoryStream(bytes), "doc.docx");

    [Fact]
    public void Inspect_and_find_cover_header_text()
    {
        var bytes = DocWithHeader("DRAFT COPY");
        var office = Office();

        var inspect = office.Inspect(Handle(bytes));
        Assert.Contains(inspect.Paragraphs, p => p.Text == "DRAFT COPY");

        var hits = office.Find(Handle(bytes), new FindQuery("DRAFT"));
        Assert.Single(hits);
    }

    [Fact]
    public void ChangeText_edits_header_paragraph()
    {
        var bytes = DocWithHeader("DRAFT COPY");
        var office = Office();
        var inspect = office.Inspect(Handle(bytes));
        var headerParaId = inspect.Paragraphs.First(p => p.Text == "DRAFT COPY").ParaId;

        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new ChangeTextOp
                {
                    Target = new TextSpanAnchor { ParaId = headerParaId, Expect = "DRAFT", Occurrence = 0 },
                    With = "FINAL",
                    Mode = ChangeMode.Direct
                }
            }
        };

        var result = office.Commit(Handle(bytes), plan);
        Assert.True(result.Committed);

        using var doc = WordprocessingDocument.Open(new MemoryStream(OfficeAgentClient.ToBytes(result)), false);
        var headerText = string.Concat(doc.MainDocumentPart!.HeaderParts.First().Header!
            .Descendants<Text>().Select(t => t.Text));
        Assert.Equal("FINAL COPY", headerText);
    }

    [Fact]
    public void Header_edit_changes_snapshot_etag()
    {
        var bytes = DocWithHeader("DRAFT COPY");
        var office = Office();
        var before = office.Inspect(Handle(bytes)).Snapshot.ETag;
        var headerParaId = office.Inspect(Handle(bytes)).Paragraphs.First(p => p.Text == "DRAFT COPY").ParaId;

        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new ChangeTextOp
                {
                    Target = new TextSpanAnchor { ParaId = headerParaId, Expect = "DRAFT", Occurrence = 0 },
                    With = "FINAL",
                    Mode = ChangeMode.Direct
                }
            }
        };

        var edited = OfficeAgentClient.ToBytes(office.Commit(Handle(bytes), plan));
        var after = office.Inspect(Handle(edited)).Snapshot.ETag;

        Assert.NotEqual(before, after);
    }

    [Fact]
    public async Task Async_inspect_and_commit_round_trip_over_files()
    {
        var input = Path.Combine(Path.GetTempPath(), $"officeagent-{Guid.NewGuid():N}.docx");
        await File.WriteAllBytesAsync(input, DocxFactory.Contract());
        try
        {
            var office = Office();
            var inspect = await office.InspectAsync(new FileHandle(input));
            var paraId = inspect.Paragraphs.First(p => p.Text.Contains("shall provide")).ParaId;

            var plan = new DocumentPlan
            {
                Operations = new PlanOperation[]
                {
                    new ChangeTextOp
                    {
                        Target = new TextSpanAnchor { ParaId = paraId, Expect = "Acme Corp", Occurrence = 0 },
                        With = "Globex",
                        Mode = ChangeMode.Direct
                    }
                }
            };

            var result = await office.CommitAsync(new FileHandle(input), plan);
            Assert.True(result.Committed);
        }
        finally
        {
            File.Delete(input);
        }
    }

    [Fact]
    public async Task Cancelled_apply_returns_structured_error_from_tool()
    {
        using var workspace = new ToolsWorkspace();
        var tools = new OfficeAgent.AgentFramework.OfficeAgentTools(workspace.Client);
        var added = await workspace.Client.RegisterBytesAsync("workspace", workspace.Root, DocxFactory.Contract(), "contract.docx");

        using var cts = new CancellationTokenSource();
        cts.Cancel();

        var planJson = "{ \"operations\": [] }";
        var report = await tools.ApplyPlan("workspace", added.ItemId, planJson, cancellationToken: cts.Token);

        using var parsed = JsonDocument.Parse(report);
        Assert.False(parsed.RootElement.GetProperty("isValid").GetBoolean());
        Assert.Equal("cancelled", parsed.RootElement.GetProperty("errors")[0].GetProperty("Code").GetString());
    }

    [Fact]
    public void DI_AddWordFormat_composes_a_working_client()
    {
        var services = new ServiceCollection();
        services.AddWordFormat();
        services.AddOfficeAgent();

        using var provider = services.BuildServiceProvider();
        var client = provider.GetRequiredService<OfficeAgentClient>();
        var inspect = client.Inspect(Handle(DocxFactory.Contract()));

        Assert.Equal(OfficeAgent.Abstractions.DocumentFormat.Word, inspect.Format);
    }

    [Fact]
    public async Task DI_contributed_handler_extends_the_module()
    {
        var services = new ServiceCollection();
        services.AddSingleton<IOperationHandler, NoOpStampHandler>();
        services.AddWordFormat();
        services.AddOfficeAgent();

        using var provider = services.BuildServiceProvider();
        var client = provider.GetRequiredService<OfficeAgentClient>();

        var input = Path.Combine(Path.GetTempPath(), $"officeagent-{Guid.NewGuid():N}.docx");
        await File.WriteAllBytesAsync(input, DocxFactory.Contract());
        try
        {
            // FillOp targeting a NodeAnchor is claimed by no built-in handler, only the contributed one.
            var plan = new DocumentPlan
            {
                Operations = new PlanOperation[]
                {
                    new FillOp { Target = new NodeAnchor { Kind = "stamp" }, Value = "x" }
                }
            };

            var report = client.Preview(new FileHandle(input), plan);
            Assert.True(report.IsValid);
            Assert.Contains(report.Changes, c => c.Verb == "stamp");
        }
        finally
        {
            File.Delete(input);
        }
    }

    [Fact]
    public async Task Reserved_verb_is_no_longer_part_of_the_contract()
    {
        using var workspace = new ToolsWorkspace();
        var tools = new OfficeAgent.AgentFramework.OfficeAgentTools(workspace.Client);
        var added = await workspace.Client.RegisterBytesAsync("workspace", workspace.Root, DocxFactory.Contract(), "contract.docx");

        var planJson = "{ \"operations\": [ { \"op\": \"setTheme\", " +
                       "\"target\": { \"$anchor\": \"style\", \"styleId\": \"Normal\" }, \"theme\": \"Dark\" } ] }";

        var report = await tools.PreviewPlan("workspace", added.ItemId, planJson);
        using var parsed = JsonDocument.Parse(report);

        Assert.False(parsed.RootElement.GetProperty("isValid").GetBoolean());
        Assert.Equal("invalid-json", parsed.RootElement.GetProperty("errors")[0].GetProperty("Code").GetString());
    }

    private sealed class ToolsWorkspace : IDisposable
    {
        private readonly ServiceProvider _serviceProvider;

        public ToolsWorkspace()
        {
            Root = Path.Combine(Path.GetTempPath(), $"officeagent-followup-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Root);

            var services = new ServiceCollection();
            services.AddWordFormat();
            services.AddFileSystemDocumentProvider("workspace", Root);
            services.AddOfficeAgent();
            _serviceProvider = services.BuildServiceProvider();
            Client = _serviceProvider.GetRequiredService<OfficeAgentClient>();
        }

        public string Root { get; }
        public OfficeAgentClient Client { get; }

        public void Dispose()
        {
            _serviceProvider.Dispose();
            if (Directory.Exists(Root)) Directory.Delete(Root, recursive: true);
        }
    }

    private static byte[] DocWithHeader(string headerText)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();

            var headerPart = main.AddNewPart<HeaderPart>();
            headerPart.Header = new Header(new Paragraph(
                new Run(new Text(headerText) { Space = SpaceProcessingModeValues.Preserve })));
            headerPart.Header.Save();
            var headerId = main.GetIdOfPart(headerPart);

            var sectPr = new SectionProperties(
                new HeaderReference { Type = HeaderFooterValues.Default, Id = headerId });

            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Body text"))),
                sectPr));
            main.Document.Save();
        }
        return ms.ToArray();
    }

    /// <summary>A trivial contributed handler that previews a deterministic stamp change.</summary>
    private sealed class NoOpStampHandler : IOperationHandler
    {
        public bool CanHandle(PlanOperation operation) =>
            operation is FillOp { Target: NodeAnchor { Kind: "stamp" } };

        public OperationPreview Preview(ApplyContext context, PlanOperation operation) =>
            OperationPreview.Ok(new ProposedChange { Verb = "stamp", Before = "", After = "stamped" });

        public void Apply(ApplyContext context, PlanOperation operation) { }
    }
}
