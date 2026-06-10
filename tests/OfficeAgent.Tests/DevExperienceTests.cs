using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

public class DevExperienceTests
{
    private static OfficeAgentClient Office() => new(new WordModule());

    [Fact]
    public async Task Replace_first_occurrence_via_storage_flow()
    {
        using var workspace = new Workspace();
        var client = workspace.Client();
        var doc = await client.RegisterBytesAsync("workspace", workspace.Root, DocxFactory.Contract(), "contract.docx");

        var hits = await client.FindAsync("workspace", doc.ItemId, new FindQuery("Acme Corp"));
        Assert.NotEmpty(hits);
        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new ChangeTextOp { Target = hits[0].Anchor, With = "Globex Inc.", Mode = ChangeMode.Direct }
            }
        };

        var result = await client.CommitAsync("workspace", doc.ItemId, plan);
        Assert.True(result.Committed);

        using var content = await client.OpenReadAsync(result.Document!);
        using var saved = WordprocessingDocument.Open(content.Stream, false);
        var text = string.Concat(saved.MainDocumentPart!.Document.Body!.Descendants<Text>().Select(t => t.Text));
        Assert.Contains("Globex Inc.", text);
    }

    [Fact]
    public async Task Replace_every_occurrence_via_storage_flow()
    {
        using var workspace = new Workspace();
        var client = workspace.Client();
        var doc = await client.RegisterBytesAsync("workspace", workspace.Root, DocxFactory.Contract(), "contract.docx");

        var hits = await client.FindAsync("workspace", doc.ItemId, new FindQuery("Acme Corp"));
        // Apply in descending occurrence order per paragraph so an earlier replacement
        // does not shift a later occurrence's index.
        var ordered = hits
            .OrderBy(h => (h.Anchor as TextSpanAnchor)?.ParaId, StringComparer.Ordinal)
            .ThenByDescending(h => (h.Anchor as TextSpanAnchor)?.Occurrence ?? 0)
            .ToArray();
        var ops = ordered
            .Select(h => (PlanOperation)new ChangeTextOp { Target = h.Anchor, With = "Globex", Mode = ChangeMode.Direct })
            .ToArray();

        var result = await client.CommitAsync("workspace", doc.ItemId, new DocumentPlan { Operations = ops });
        Assert.True(result.Committed);
        Assert.Equal(2, result.Report.Changes.Count);

        using var content = await client.OpenReadAsync(result.Document!);
        using var saved = WordprocessingDocument.Open(content.Stream, false);
        var text = string.Concat(saved.MainDocumentPart!.Document.Body!.Descendants<Text>().Select(t => t.Text));
        Assert.DoesNotContain("Acme Corp", text);
    }

    [Fact]
    public async Task Fill_content_control_by_tag_via_storage_flow()
    {
        using var workspace = new Workspace();
        var client = workspace.Client();
        var doc = await client.RegisterBytesAsync("workspace", workspace.Root, DocxFactory.Contract(), "contract.docx");

        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new FillOp
                {
                    Target = new StructuralAnchor { Id = $"cc:{DocxFactory.ClientControlTag}", Tag = DocxFactory.ClientControlTag, Kind = "contentControl" },
                    Value = "Globex"
                }
            }
        };

        var result = await client.CommitAsync("workspace", doc.ItemId, plan);
        Assert.True(result.Committed);

        using var content = await client.OpenReadAsync(result.Document!);
        using var saved = WordprocessingDocument.Open(content.Stream, false);
        var sdt = saved.MainDocumentPart!.Document.Body!.Descendants<SdtRun>().Single();
        var value = string.Concat(sdt.Descendants<Text>().Select(t => t.Text));
        Assert.Equal("Globex", value);
    }

    [Fact]
    public void ApplyResult_instance_methods_match_legacy_statics()
    {
        var bytes = DocxFactory.Contract();
        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new SetPropertyOp { Target = new NodeAnchor { Kind = "docProperty", Path = "core/title" }, Value = "T" }
            }
        };

        using var result = Office().Commit(new StreamHandle(new MemoryStream(bytes)), plan);
        Assert.True(result.Committed);

        var fromInstance = result.ToBytes();
        var fromStatic = OfficeAgentClient.ToBytes(result);

        Assert.Equal(fromInstance, fromStatic);
    }

    [Fact]
    public void ValidationError_exposes_typed_CodeKind()
    {
        var err = new ValidationError(ValidationErrorCode.StaleSnapshot, "drifted");
        Assert.Equal(ValidationErrorCodes.StaleSnapshot, err.Code);
        Assert.Equal(ValidationErrorCode.StaleSnapshot, err.CodeKind);

        // Forward compat: an unknown wire code parses to Unknown without throwing.
        var unknown = new ValidationError("future-code", "x");
        Assert.Equal(ValidationErrorCode.Unknown, unknown.CodeKind);
    }

    [Fact]
    public void AddOfficeAgent_without_format_module_throws_at_resolve_time()
    {
        var services = new ServiceCollection();
        services.AddOfficeAgent();
        using var provider = services.BuildServiceProvider();

        var ex = Assert.Throws<InvalidOperationException>(
            () => provider.GetRequiredService<OfficeAgentClient>());
        Assert.Contains("IFormatModule", ex.Message, StringComparison.Ordinal);
        Assert.Contains("AddWordFormat", ex.Message, StringComparison.Ordinal);
    }

    private sealed class Workspace : IDisposable
    {
        private readonly ServiceProvider _serviceProvider;

        public Workspace()
        {
            Root = Path.Combine(Path.GetTempPath(), $"officeagent-dx-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Root);

            var services = new ServiceCollection();
            services.AddWordFormat();
            services.AddFileSystemDocumentProvider("workspace", Root);
            services.AddOfficeAgent();
            _serviceProvider = services.BuildServiceProvider();
        }

        public string Root { get; }

        public OfficeAgentClient Client() => _serviceProvider.GetRequiredService<OfficeAgentClient>();

        public void Dispose()
        {
            _serviceProvider.Dispose();
            if (Directory.Exists(Root)) Directory.Delete(Root, recursive: true);
        }
    }
}
