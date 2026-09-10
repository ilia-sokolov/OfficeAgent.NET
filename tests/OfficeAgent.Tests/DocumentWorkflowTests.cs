using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.DependencyInjection;
using System.Text.Json;
using OfficeAgent.AgentFramework;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

public sealed class DocumentWorkflowTests
{
    [Fact]
    public async Task Agent_template_tool_creates_an_output_with_a_receipt()
    {
        using var workspace = new WorkflowWorkspace();
        var client = workspace.Client();
        var source = await client.RegisterBytesAsync("workspace", workspace.Root, QuoteTemplate(), "agent-template.docx");
        var tools = new OfficeAgentTools(client);

        using var response = JsonDocument.Parse(await tools.PopulateTemplateBatch(
            "workspace",
            source.ItemId,
            """
            {
              "items": [{
                "outputName": "agent-quote.docx",
                "binding": {
                  "values": { "CustomerName": "Agent Customer" },
                  "repeatingTables": [{
                    "tablePath": "table#0",
                    "templateRowIndex": 1,
                    "records": [{ "Description": "Review", "Amount": "500.00" }]
                  }]
                }
              }]
            }
            """));

        Assert.True(response.RootElement.GetProperty("committed").GetBoolean());
        var item = Assert.Single(response.RootElement.GetProperty("items").EnumerateArray());
        Assert.True(item.GetProperty("committed").GetBoolean());
        Assert.Equal("Committed", item.GetProperty("receipt").GetProperty("outcome").GetString());
    }

    [Fact]
    public async Task Agent_comparison_tool_returns_a_plan_that_previews_against_the_original()
    {
        using var workspace = new WorkflowWorkspace();
        var client = workspace.Client();
        var original = await client.RegisterBytesAsync("workspace", workspace.Root,
            ParagraphDocument("One", "Old"), "agent-original.docx");
        var revised = await client.RegisterBytesAsync("workspace", workspace.Root,
            ParagraphDocument("One", "New"), "agent-revised.docx");
        var tools = new OfficeAgentTools(client);

        using var comparison = JsonDocument.Parse(await tools.CompareDocuments(
            "workspace", original.ItemId, "workspace", revised.ItemId, "Agent Compare"));

        Assert.True(comparison.RootElement.GetProperty("isComplete").GetBoolean());
        var planJson = comparison.RootElement.GetProperty("plan").GetRawText();
        using var preview = JsonDocument.Parse(await tools.PreviewPlan("workspace", original.ItemId, planJson));
        Assert.True(preview.RootElement.GetProperty("isValid").GetBoolean());
    }

    [Fact]
    public async Task Template_batch_populates_scalar_and_repeating_rows_without_changing_source()
    {
        using var workspace = new WorkflowWorkspace();
        var client = workspace.Client();
        var source = await client.RegisterBytesAsync("workspace", workspace.Root, QuoteTemplate(), "quote-template.docx");

        var result = await client.PopulateTemplateBatchAsync(source, new TemplateBatchRequest
        {
            Items = new[]
            {
                new TemplateBatchItem
                {
                    OutputName = "quote-acme.docx",
                    Binding = new TemplateBinding
                    {
                        Values = new Dictionary<string, string?> { ["CustomerName"] = "Acme BV" },
                        RepeatingTables = new[]
                        {
                            new RepeatingTableBinding
                            {
                                TablePath = "table#0",
                                TemplateRowIndex = 1,
                                Records = new IReadOnlyDictionary<string, string?>[]
                                {
                                    new Dictionary<string, string?> { ["Description"] = "Consulting", ["Amount"] = "1200.00" },
                                    new Dictionary<string, string?> { ["Description"] = "Support", ["Amount"] = "300.00" }
                                }
                            }
                        }
                    }
                }
            }
        });

        var item = Assert.Single(result.Items);
        Assert.True(item.Committed, string.Join("; ", item.Diagnostics.Select(d => d.Message)));
        Assert.NotNull(item.Receipt);
        using (var output = await client.OpenReadAsync(item.Document!))
        using (var document = WordprocessingDocument.Open(output.Stream, false))
        {
            var body = document.MainDocumentPart!.Document.Body!;
            Assert.Contains("Acme BV", body.InnerText);
            Assert.DoesNotContain("{{Description}}", body.InnerText);
            Assert.Equal(3, body.Descendants<TableRow>().Count());
            Assert.Contains("Consulting", body.InnerText);
            Assert.Contains("Support", body.InnerText);
            var ids = body.Descendants<Paragraph>().Select(p => p.ParagraphId?.Value).Where(id => id is not null).ToList();
            Assert.Equal(ids.Count, ids.Distinct(StringComparer.Ordinal).Count());
        }

        using (var original = await client.OpenReadAsync(source))
        using (var document = WordprocessingDocument.Open(original.Stream, false))
        {
            Assert.Contains("PLACEHOLDER", document.MainDocumentPart!.Document.InnerText);
            Assert.Contains("{{Description}}", document.MainDocumentPart.Document.InnerText);
        }
    }

    [Fact]
    public async Task Template_plan_rejects_duplicate_scalar_tags()
    {
        using var workspace = new WorkflowWorkspace();
        var client = workspace.Client();
        var source = await client.RegisterBytesAsync("workspace", workspace.Root, DuplicateTagTemplate(), "duplicate.docx");

        var result = await client.BuildTemplatePlanAsync(source, new TemplateBinding
        {
            Values = new Dictionary<string, string?> { ["CustomerName"] = "Acme" }
        });

        Assert.False(result.IsValid);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "ambiguous-template-slot");
    }

    [Fact]
    public async Task Template_plan_rejects_duplicate_tags_even_when_the_binding_omits_them()
    {
        using var workspace = new WorkflowWorkspace();
        var client = workspace.Client();
        var source = await client.RegisterBytesAsync("workspace", workspace.Root, DuplicateTagTemplate(), "duplicate-omitted.docx");

        var result = await client.BuildTemplatePlanAsync(source, new TemplateBinding
        {
            MissingValueBehavior = MissingTemplateValueBehavior.Ignore
        });

        Assert.False(result.IsValid);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "ambiguous-template-slot");
    }

    [Fact]
    public void Empty_repeating_collection_removes_the_template_row()
    {
        var client = new OfficeAgentClient(new WordModule());
        var inspect = client.Inspect(QuoteTemplate());
        using var applied = client.Commit(new StreamHandle(new MemoryStream(QuoteTemplate())), new DocumentPlan
        {
            Snapshot = inspect.Snapshot,
            Operations = new PlanOperation[]
            {
                new RepeatTableRowOp
                {
                    Target = new NodeAnchor { Kind = "table", Path = "table#0" },
                    TemplateRowIndex = 1,
                    Mode = ChangeMode.Direct
                }
            }
        });

        Assert.True(applied.Committed, string.Join("; ", applied.Report.Errors.Select(error => error.Message)));
        using var document = WordprocessingDocument.Open(new MemoryStream(applied.ToBytes()), false);
        Assert.Single(document.MainDocumentPart!.Document.Body!.Descendants<TableRow>());
    }

    [Fact]
    public void Repeating_row_with_existing_revisions_is_rejected_without_mutation()
    {
        var input = TemplateWithRevisionInRow();
        var client = new OfficeAgentClient(new WordModule());
        using var result = client.Commit(new StreamHandle(new MemoryStream(input)), new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new RepeatTableRowOp
                {
                    Target = new NodeAnchor { Kind = "table", Path = "table#0" },
                    TemplateRowIndex = 1,
                    Records = new[]
                    {
                        (IReadOnlyDictionary<string, string?>)new Dictionary<string, string?>
                        {
                            ["Description"] = "Consulting", ["Amount"] = "1200.00"
                        }
                    },
                    Mode = ChangeMode.Direct
                }
            }
        });

        Assert.False(result.Committed);
        Assert.Contains(result.Report.Errors, error =>
            error.Code == ValidationErrorCodes.InvalidOperation && error.Message.Contains("tracked revisions"));
    }

    [Fact]
    public void Tracked_repeating_rows_accept_to_population_and_reject_to_template()
    {
        var original = QuoteTemplate();
        var client = new OfficeAgentClient(new WordModule());
        var redline = Apply(client, original, new RepeatTableRowOp
        {
            Target = new NodeAnchor { Kind = "table", Path = "table#0" },
            TemplateRowIndex = 1,
            Records = new[]
            {
                (IReadOnlyDictionary<string, string?>)new Dictionary<string, string?>
                {
                    ["Description"] = "Consulting", ["Amount"] = "1200.00"
                }
            },
            Mode = ChangeMode.Tracked
        });

        Assert.Contains(client.Inspect(redline).Nodes, node => node.Path.StartsWith("rowIns#", StringComparison.Ordinal));
        Assert.Contains(client.Inspect(redline).Nodes, node => node.Path.StartsWith("rowDel#", StringComparison.Ordinal));
        var accepted = Apply(client, redline, ResolveAll(RevisionAction.Accept));
        var rejected = Apply(client, redline, ResolveAll(RevisionAction.Reject));
        Assert.Contains("Consulting", BodyText(accepted));
        Assert.DoesNotContain("{{Description}}", BodyText(accepted));
        Assert.Contains("{{Description}}", BodyText(rejected));
        Assert.DoesNotContain("Consulting", BodyText(rejected));
    }

    [Fact]
    public void Comparison_plan_accepts_to_revised_and_rejects_to_original()
    {
        var original = ParagraphDocument("Introduction", "Old clause", "Closing");
        var revised = ParagraphDocument("Introduction", "New clause", "Additional clause", "Closing");
        var client = new OfficeAgentClient(new WordModule());

        var comparison = client.CompareDocuments(original, revised, new DocumentComparisonOptions
        {
            Revision = new RevisionMetadata { Author = "Comparison Bot" }
        });

        Assert.True(comparison.IsComplete, string.Join("; ", comparison.Diagnostics.Select(d => d.Message)));
        Assert.NotNull(comparison.Plan);
        Assert.Collection(comparison.Differences,
            difference => Assert.Equal(DocumentDifferenceKind.Changed, difference.Kind),
            difference => Assert.Equal(DocumentDifferenceKind.Added, difference.Kind));

        var redline = Apply(client, original, comparison.Plan!);
        var accepted = Apply(client, redline, ResolveAll(RevisionAction.Accept));
        var rejected = Apply(client, redline, ResolveAll(RevisionAction.Reject));

        Assert.Equal(Texts(client, revised), Texts(client, accepted));
        Assert.Equal(Texts(client, original), Texts(client, rejected));
        Assert.DoesNotContain(client.Inspect(accepted).Nodes, node => node.Kind == "revision");
    }

    [Fact]
    public void Comparison_reports_unsupported_table_changes_and_returns_no_plan()
    {
        var original = QuoteTemplate();
        var revised = ReplacePackageText(original, "{{Amount}}", "Changed");
        var client = new OfficeAgentClient(new WordModule());

        var comparison = client.CompareDocuments(original, revised);

        Assert.False(comparison.IsComplete);
        Assert.Null(comparison.Plan);
        Assert.Contains(comparison.Diagnostics, diagnostic => diagnostic.Code == "unsupported-non-body-change");
    }

    [Fact]
    public void Comparison_plan_is_rejected_after_original_drift()
    {
        var original = ParagraphDocument("One", "Two");
        var revised = ParagraphDocument("One", "Changed");
        var drifted = ParagraphDocument("Drifted", "Two");
        var client = new OfficeAgentClient(new WordModule());
        var comparison = client.CompareDocuments(original, revised);

        using var applied = client.Commit(new StreamHandle(new MemoryStream(drifted)), comparison.Plan!);

        Assert.False(applied.Committed);
        Assert.Contains(applied.Report.Errors, error => error.Code == ValidationErrorCodes.StaleSnapshot);
    }

    [Fact]
    public void Comparison_paragraph_removal_accepts_and_rejects_cleanly()
    {
        var original = ParagraphDocument("Keep", "Remove me", "Closing");
        var revised = ParagraphDocument("Keep", "Closing");
        var client = new OfficeAgentClient(new WordModule());
        var comparison = client.CompareDocuments(original, revised);

        Assert.True(comparison.IsComplete);
        Assert.Equal(DocumentDifferenceKind.Removed, Assert.Single(comparison.Differences).Kind);
        var redline = Apply(client, original, comparison.Plan!);

        Assert.Equal(Texts(client, revised), Texts(client, Apply(client, redline, ResolveAll(RevisionAction.Accept))));
        Assert.Equal(Texts(client, original), Texts(client, Apply(client, redline, ResolveAll(RevisionAction.Reject))));
    }

    [Fact]
    public void Comparison_honors_cancellation_and_resource_limits()
    {
        var client = new OfficeAgentClient(new WordModule());
        var document = ParagraphDocument("One", "Two");
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();

        Assert.Throws<OperationCanceledException>(() => client.CompareDocuments(
            document, document, new DocumentComparisonOptions(), cancelled.Token));
        var limited = client.CompareDocuments(document, document, new DocumentComparisonOptions
        {
            MaximumDocumentBytes = 1
        });
        Assert.Contains(limited.Diagnostics, diagnostic => diagnostic.Code == "comparison-input-limit-exceeded");
        Assert.Throws<ArgumentOutOfRangeException>(() => client.CompareDocuments(
            document, document, new DocumentComparisonOptions { MaximumParagraphs = 0 }));
        var exactlyOne = client.CompareDocuments(
            ParagraphDocument("Before"), ParagraphDocument("After"),
            new DocumentComparisonOptions { MaximumDifferences = 1 });
        Assert.True(exactlyOne.IsComplete);
    }

    [Fact]
    public async Task Provider_comparison_returns_a_stable_diagnostic_when_the_byte_limit_is_exceeded()
    {
        using var workspace = new WorkflowWorkspace();
        var client = workspace.Client();
        var original = await client.RegisterBytesAsync("workspace", workspace.Root, ParagraphDocument("One"), "limited-original.docx");
        var revised = await client.RegisterBytesAsync("workspace", workspace.Root, ParagraphDocument("Two"), "limited-revised.docx");

        var comparison = await client.CompareDocumentsAsync(original, revised,
            new DocumentComparisonOptions { MaximumDocumentBytes = 1 });

        Assert.False(comparison.IsComplete);
        Assert.Contains(comparison.Diagnostics, diagnostic => diagnostic.Code == "comparison-input-limit-exceeded");
    }

    [Fact]
    public void Comparison_with_changed_text_and_style_withholds_the_plan()
    {
        var original = ParagraphDocument("Old");
        var revised = ParagraphDocument("New");
        using (var stream = new MemoryStream())
        {
            stream.Write(revised, 0, revised.Length);
            stream.Position = 0;
            using (var document = WordprocessingDocument.Open(stream, true))
            {
                document.MainDocumentPart!.Document.Body!.Elements<Paragraph>().Single()
                    .ParagraphProperties = new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" });
                document.MainDocumentPart.Document.Save();
            }
            revised = stream.ToArray();
        }

        var comparison = new OfficeAgentClient(new WordModule()).CompareDocuments(original, revised);

        Assert.False(comparison.IsComplete);
        Assert.Null(comparison.Plan);
        Assert.Contains(comparison.Diagnostics, diagnostic => diagnostic.Code == "unsupported-style-change");
    }

    [Fact]
    public void Comparison_reports_when_original_has_no_paragraph_anchor_for_an_addition()
    {
        var original = TableOnlyDocument();
        var revised = ParagraphDocument("Added paragraph");
        var client = new OfficeAgentClient(new WordModule());

        var comparison = client.CompareDocuments(original, revised);

        Assert.False(comparison.IsComplete);
        Assert.Null(comparison.Plan);
        Assert.Contains(comparison.Diagnostics, diagnostic => diagnostic.Code == "comparison-anchor-unavailable");
    }

    private static RevisionOp ResolveAll(RevisionAction action) => new()
    {
        Target = new NodeAnchor { Kind = "revision", Path = "all" },
        Action = action
    };

    private static byte[] Apply(OfficeAgentClient client, byte[] input, PlanOperation operation) =>
        Apply(client, input, new DocumentPlan { Operations = new[] { operation } });

    private static byte[] Apply(OfficeAgentClient client, byte[] input, DocumentPlan plan)
    {
        using var applied = client.Commit(new StreamHandle(new MemoryStream(input)), plan);
        Assert.True(applied.Committed, string.Join("; ", applied.Report.Errors.Select(error => error.Message)));
        return applied.ToBytes();
    }

    private static IReadOnlyList<string> Texts(OfficeAgentClient client, byte[] document) =>
        client.Inspect(document).Paragraphs
            .Where(paragraph => paragraph.Location == "body" && paragraph.In is null)
            .Select(paragraph => paragraph.Text).ToArray();

    private static string BodyText(byte[] document)
    {
        using var package = WordprocessingDocument.Open(new MemoryStream(document), false);
        return package.MainDocumentPart!.Document.Body!.InnerText;
    }

    private static byte[] ParagraphDocument(params string[] texts)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            var body = new Body();
            for (var i = 0; i < texts.Length; i++)
                body.AppendChild(new Paragraph(new Run(new Text(texts[i])))
                {
                    ParagraphId = (i + 1).ToString("X8")
                });
            main.Document = new Document(body);
            main.Document.Save();
        }
        return stream.ToArray();
    }

    private static byte[] QuoteTemplate()
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            var customer = new SdtRun(
                new SdtProperties(new Tag { Val = "CustomerName" }, new SdtId { Val = 10 }),
                new SdtContentRun(new Run(new Text("PLACEHOLDER"))));
            var table = new Table(
                new TableProperties(new TableStyle { Val = "TableGrid" }),
                Row("Description", "Amount"),
                Row("{{Description}}", "{{Amount}}"));
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Quote for ")), customer) { ParagraphId = "00000001" },
                table,
                new Paragraph { ParagraphId = "00000004" }));
            main.Document.Save();
        }
        return stream.ToArray();
    }

    private static byte[] TableOnlyDocument()
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Table(Row("Only table"))));
            main.Document.Save();
        }
        return stream.ToArray();
    }

    private static byte[] DuplicateTagTemplate()
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(
                TaggedParagraph("CustomerName", 1, "One"),
                TaggedParagraph("CustomerName", 2, "Two")));
            main.Document.Save();
        }
        return stream.ToArray();
    }

    private static byte[] TemplateWithRevisionInRow()
    {
        var bytes = QuoteTemplate();
        using var stream = new MemoryStream();
        stream.Write(bytes, 0, bytes.Length);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var row = document.MainDocumentPart!.Document.Body!.Descendants<TableRow>().ElementAt(1);
            var run = row.Descendants<Run>().First();
            run.InsertAfterSelf(new InsertedRun((Run)run.CloneNode(true)));
            run.Remove();
            document.MainDocumentPart.Document.Save();
        }
        return stream.ToArray();
    }

    private static Paragraph TaggedParagraph(string tag, int id, string value) => new(
        new SdtRun(
            new SdtProperties(new Tag { Val = tag }, new SdtId { Val = id }),
            new SdtContentRun(new Run(new Text(value)))));

    private static TableRow Row(params string[] values) => new(
        values.Select(value => new TableCell(
            new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "2400" }),
            new Paragraph(new Run(new Text(value))))));

    private static byte[] ReplacePackageText(byte[] input, string before, string after)
    {
        using var stream = new MemoryStream();
        stream.Write(input, 0, input.Length);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var text = document.MainDocumentPart!.Document.Descendants<Text>().Single(node => node.Text == before);
            text.Text = after;
            document.MainDocumentPart.Document.Save();
        }
        return stream.ToArray();
    }

    private sealed class WorkflowWorkspace : IDisposable
    {
        private readonly ServiceProvider _provider;
        public WorkflowWorkspace()
        {
            Root = Path.Combine(Path.GetTempPath(), "officeagent-workflows-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Root);
            var services = new ServiceCollection();
            services.AddWordFormat();
            services.AddFileSystemDocumentProvider("workspace", Root);
            services.AddOfficeAgent();
            _provider = services.BuildServiceProvider();
        }
        public string Root { get; }
        public OfficeAgentClient Client() => _provider.GetRequiredService<OfficeAgentClient>();
        public void Dispose()
        {
            _provider.Dispose();
            if (Directory.Exists(Root)) Directory.Delete(Root, true);
        }
    }
}
