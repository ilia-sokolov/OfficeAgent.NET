using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeAgent.Abstractions;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

public class TableDiscoveryTests
{
    private static OfficeAgentClient Office() => new(new WordModule());
    private static StreamHandle Handle(byte[] bytes) => new(new MemoryStream(bytes), "doc.docx");

    [Fact]
    public async Task InspectDocument_tool_now_surfaces_nodes_with_table_paths()
    {
        // The previous gap: TableNodeProvider produced NodeInfo correctly, but the
        // inspect_document tool dropped the entire `nodes` array from its JSON so the
        // LLM could not discover any table#N paths.
        using var workspace = new ToolsWorkspace();
        var tools = new OfficeAgent.AgentFramework.OfficeAgentTools(workspace.Client);
        var added = await workspace.Client.RegisterBytesAsync("workspace", workspace.Root, BuildCountriesDocument(), "doc.docx");

        using var report = JsonDocument.Parse(await tools.InspectDocument("workspace", added.ItemId));

        Assert.True(report.RootElement.TryGetProperty("nodes", out var nodes));
        var tableNode = nodes.EnumerateArray().FirstOrDefault(n => n.GetProperty("Kind").GetString() == "table");
        Assert.NotEqual(JsonValueKind.Undefined, tableNode.ValueKind);
        Assert.Equal("table#0", tableNode.GetProperty("Path").GetString());
        Assert.Contains("row", tableNode.GetProperty("Summary").GetString());
    }

    private sealed class ToolsWorkspace : IDisposable
    {
        private readonly Microsoft.Extensions.DependencyInjection.ServiceProvider _serviceProvider;

        public ToolsWorkspace()
        {
            Root = Path.Combine(Path.GetTempPath(), $"officeagent-tables-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Root);

            var services = new Microsoft.Extensions.DependencyInjection.ServiceCollection();
            services.AddWordFormat();
            services.AddFileSystemDocumentProvider("workspace", Root);
            services.AddOfficeAgent();
            _serviceProvider = Microsoft.Extensions.DependencyInjection.ServiceCollectionContainerBuilderExtensions.BuildServiceProvider(services);
            Client = Microsoft.Extensions.DependencyInjection.ServiceProviderServiceExtensions.GetRequiredService<OfficeAgentClient>(_serviceProvider);
        }

        public string Root { get; }
        public OfficeAgentClient Client { get; }

        public void Dispose()
        {
            _serviceProvider.Dispose();
            if (Directory.Exists(Root)) Directory.Delete(Root, recursive: true);
        }
    }

    [Fact]
    public void Paragraphs_inside_a_table_advertise_their_container_via_In()
    {
        var bytes = BuildCountriesDocument();
        var inspect = Office().Inspect(Handle(bytes));

        var inTable = inspect.Paragraphs.Where(p => p.In is not null).ToList();
        Assert.NotEmpty(inTable);
        Assert.All(inTable, p => Assert.Equal("table#0", p.In));

        var freeFlowing = inspect.Paragraphs.Where(p => p.In is null).ToList();
        Assert.Contains(freeFlowing, p => p.Text.Contains("Intro"));
    }

    [Fact]
    public void RemoveTableRows_with_explicit_indices_drops_the_specified_rows()
    {
        var bytes = BuildCountriesDocument();
        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new RemoveTableRowsOp
                {
                    Target = new NodeAnchor { Kind = "table", Path = "table#0" },
                    RowIndices = new[] { -1 } // last row
                }
            }
        };

        var result = Office().Commit(Handle(bytes), plan);
        Assert.True(result.Committed);

        using var doc = WordprocessingDocument.Open(new MemoryStream(OfficeAgentClient.ToBytes(result)), false);
        var rows = doc.MainDocumentPart!.Document.Body!.Descendants<Table>().Single().Elements<TableRow>().ToList();
        Assert.Equal(3, rows.Count); // header + RU + US (UK was last and removed)
        Assert.Contains("US", rows[2].InnerText);
    }

    [Fact]
    public void RemoveTableRows_onlyIfEmpty_drops_blank_rows_only()
    {
        var bytes = BuildCountriesDocumentWithBlankRows();
        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new RemoveTableRowsOp
                {
                    Target = new NodeAnchor { Kind = "table", Path = "table#0" },
                    OnlyIfEmpty = true
                }
            }
        };

        var result = Office().Commit(Handle(bytes), plan);
        Assert.True(result.Committed);

        using var doc = WordprocessingDocument.Open(new MemoryStream(OfficeAgentClient.ToBytes(result)), false);
        var rows = doc.MainDocumentPart!.Document.Body!.Descendants<Table>().Single().Elements<TableRow>().ToList();

        // Original: header + RU + blank + US + blank + blank + UK = 7 rows; expect 4 after onlyIfEmpty
        Assert.Equal(4, rows.Count);
        foreach (var row in rows)
            Assert.False(string.IsNullOrWhiteSpace(row.InnerText));
    }

    [Fact]
    public void ChangeText_with_empty_expect_returns_actionable_error()
    {
        // Used to surface as expect-mismatch (because IndexOfOccurrence rejects an empty
        // needle), leaving the LLM to guess what was wrong. Now it is invalid-operation
        // with a sentence that tells the agent exactly how to express "delete a row".
        var bytes = BuildCountriesDocument();
        var paraId = Office().Inspect(Handle(bytes)).Paragraphs.First().ParaId;

        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new ChangeTextOp
                {
                    Target = new TextSpanAnchor { ParaId = paraId, Expect = "", Occurrence = 0 },
                    With = "anything"
                }
            }
        };

        var report = Office().Preview(Handle(bytes), plan);
        Assert.False(report.IsValid);
        var err = Assert.Single(report.Errors);
        Assert.Equal(ValidationErrorCodes.InvalidOperation, err.Code);
        Assert.Contains("non-empty", err.Message, StringComparison.OrdinalIgnoreCase);
    }

    private static byte[] BuildCountriesDocument()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var table = new Table(
                new TableProperties(new TableStyle { Val = "TableGrid" }),
                Row("Country", "Population mil", "Surface sq km"),
                Row("RU", "146", "17098246"),
                Row("US", "332", "9833520"),
                Row("UK", "68",  "243610"));
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Intro"))),
                table,
                new Paragraph(new Run(new Text("Outro")))));
            main.Document.Save();
        }
        return ms.ToArray();
    }

    private static byte[] BuildCountriesDocumentWithBlankRows()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var table = new Table(
                new TableProperties(new TableStyle { Val = "TableGrid" }),
                Row("Country", "Population mil", "Surface sq km"),
                Row("RU", "146", "17098246"),
                Row("", "", ""),
                Row("US", "332", "9833520"),
                Row("", "", ""),
                Row("", "", ""),
                Row("UK", "68", "243610"));
            main.Document = new Document(new Body(table));
            main.Document.Save();
        }
        return ms.ToArray();
    }

    private static TableRow Row(params string[] cells)
    {
        var row = new TableRow();
        foreach (var c in cells)
            row.AppendChild(new TableCell(new Paragraph(new Run(new Text(c) { Space = SpaceProcessingModeValues.Preserve }))));
        return row;
    }
}
