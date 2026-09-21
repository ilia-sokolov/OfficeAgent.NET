using System.IO.Compression;
using System.Security.Cryptography;
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
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

public sealed class DocumentMergeTests
{
    [Fact]
    public void Assembly_preserves_order_styles_lists_tables_sections_and_control_tags()
    {
        var inputs = new[] { Fixture("Proposal", "FF0000"), Fixture("Appendix", "0000FF", landscape: true) };
        var client = new OfficeAgentClient(new WordModule());
        var preview = client.PreviewMerge(inputs, new() { Title = "Packet", Author = "Publisher" });
        Assert.True(preview.IsValid, Errors(preview.Diagnostics));
        Assert.Equal(2, preview.Sources.Count);
        Assert.Contains(preview.Sources[1].Decisions, d => d.Contains("renamed"));
        var result = client.CommitMerge(preview.Plan!, inputs);
        Assert.True(result.Committed, Errors(result.Diagnostics));
        Assert.Equal(Convert.ToHexString(SHA256.HashData(inputs[0])).ToLowerInvariant(), result.Receipt!.Inputs[0].Sha256);
        Assert.Equal(Convert.ToHexString(SHA256.HashData(result.Content!)).ToLowerInvariant(), result.Receipt.OutputSha256);
        Assert.Equal(DocumentMergePlan.CurrentVersion, preview.Plan!.Version);
        Assert.Equal(DocumentMergeReceipt.CurrentReceiptVersion, result.Receipt.ReceiptVersion);
        var receiptRoundTrip = JsonSerializer.Deserialize<DocumentMergeReceipt>(
            JsonSerializer.Serialize(result.Receipt))!;
        Assert.Equal(result.Receipt.ReceiptVersion, receiptRoundTrip.ReceiptVersion);
        Assert.Equal(result.Receipt.PlanSha256, receiptRoundTrip.PlanSha256);
        Assert.Equal(result.Receipt.Inputs.Select(input => input.Sha256),
            receiptRoundTrip.Inputs.Select(input => input.Sha256));
        Assert.Equal(result.Receipt.OutputSha256, receiptRoundTrip.OutputSha256);
        using var doc = Open(result.Content!);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc));
        var body = doc.MainDocumentPart!.Document.Body!;
        Assert.Equal(new[] { "Proposal", "", "Appendix" }, body.Elements<Paragraph>().Select(p => p.InnerText));
        Assert.Contains(preview.Sources[0].Decisions, d => d.Contains("section-boundary paragraph"));
        Assert.Equal(2, body.Elements<Table>().Count());
        var sections = body.Descendants<SectionProperties>().ToArray();
        Assert.Equal(2, sections.Length);
        Assert.Equal(PageOrientationValues.Landscape, sections[1].GetFirstChild<PageSize>()!.Orient!.Value);
        Assert.Equal(SectionMarkValues.NextPage, sections[0].GetFirstChild<SectionType>()!.Val!.Value);
        Assert.Equal(6, sections[1].ChildElements.Count(e => e is HeaderReference or FooterReference));
        Assert.Equal(new[] { "Customer", "m2_Customer" }, body.Descendants<Tag>().Select(t => t.Val!.Value));
        var styles = doc.MainDocumentPart.StyleDefinitionsPart!.Styles;
        Assert.Equal("FF0000", styles.Elements<Style>().Single(s => s.StyleId == "m1_Clause").Descendants<Color>().Single().Val!.Value);
        Assert.Equal("0000FF", styles.Elements<Style>().Single(s => s.StyleId == "m2_Clause").Descendants<Color>().Single().Val!.Value);
        var numbers = body.Descendants<NumberingId>().Select(n => n.Val!.Value).ToArray();
        Assert.Equal(2, numbers.Distinct().Count());
        Assert.Equal("Packet", doc.PackageProperties.Title);
        Assert.Equal("Publisher", doc.PackageProperties.Creator);
        // Same inputs and options yield exact deterministic candidate bytes.
        Assert.Equal(result.Content, client.CommitMerge(preview.Plan!, inputs).Content);
    }

    [Fact]
    public void Imported_headers_and_hyperlinks_resolve_and_do_not_leak_into_blank_source()
    {
        var left = Fixture("First", "FF0000", header: "Company A");
        var right = Fixture("Second", "FF0000");
        var client = new OfficeAgentClient(new WordModule());
        var preview = client.PreviewMerge(new[] { left, right });
        Assert.True(preview.IsValid, Errors(preview.Diagnostics));
        using var doc = Open(client.CommitMerge(preview.Plan!, new[] { left, right }).Content!);
        var sections = doc.MainDocumentPart!.Document.Descendants<SectionProperties>().ToArray();
        var firstId = sections[0].Elements<HeaderReference>().Single(r => r.Type == HeaderFooterValues.Default).Id!;
        var lastId = sections[1].Elements<HeaderReference>().Single(r => r.Type == HeaderFooterValues.Default).Id!;
        Assert.Contains("Company A", ((HeaderPart)doc.MainDocumentPart.GetPartById(firstId!)).Header.InnerText);
        Assert.Equal("", ((HeaderPart)doc.MainDocumentPart.GetPartById(lastId!)).Header.InnerText);
        foreach (var link in doc.MainDocumentPart.Document.Descendants<Hyperlink>().Where(h => h.Id is not null))
            Assert.Equal("https://example.com/", doc.MainDocumentPart.HyperlinkRelationships.Single(r => r.Id == link.Id).Uri.ToString());
        Assert.Equal(2, doc.MainDocumentPart.Document.Descendants<BookmarkStart>().Select(b => b.Name!.Value).Distinct().Count());
        foreach (var link in doc.MainDocumentPart.Document.Descendants<Hyperlink>().Where(h => h.Anchor is not null))
            Assert.Contains(doc.MainDocumentPart.Document.Descendants<BookmarkStart>(), b => b.Name == link.Anchor);
    }

    [Fact]
    public void Source_boundary_overrides_a_continuous_final_break_without_changing_the_next_source()
    {
        var first = Edit(Fixture("First", "FF0000"), doc =>
            doc.MainDocumentPart!.Document.Body!.GetFirstChild<SectionProperties>()!
                .PrependChild(new SectionType { Val = SectionMarkValues.Continuous }));
        var second = Edit(Fixture("Second", "0000FF"), doc =>
            doc.MainDocumentPart!.Document.Body!.GetFirstChild<SectionProperties>()!
                .PrependChild(new SectionType { Val = SectionMarkValues.EvenPage }));
        var client = new OfficeAgentClient(new WordModule());
        var preview = client.PreviewMerge(new[] { first, second });
        Assert.True(preview.IsValid, Errors(preview.Diagnostics));
        using var document = Open(client.CommitMerge(preview.Plan!, new[] { first, second }).Content!);
        var sections = document.MainDocumentPart!.Document.Descendants<SectionProperties>().ToArray();
        Assert.Equal(SectionMarkValues.NextPage, sections[0].GetFirstChild<SectionType>()!.Val!.Value);
        Assert.Equal(SectionMarkValues.EvenPage, sections[1].GetFirstChild<SectionType>()!.Val!.Value);
    }

    [Fact]
    public void Embedded_raster_images_are_imported_with_distinct_resolvable_relationships()
    {
        var first = WithImage(Fixture("First", "FF0000"));
        var second = WithImage(Fixture("Second", "0000FF"));
        var client = new OfficeAgentClient(new WordModule());
        var preview = client.PreviewMerge(new[] { first, second });
        Assert.True(preview.IsValid, Errors(preview.Diagnostics));
        Assert.All(preview.Sources, source => Assert.Equal(1, source.Images));
        using var document = Open(client.CommitMerge(preview.Plan!, new[] { first, second }).Content!);
        var main = document.MainDocumentPart!;
        Assert.Equal(2, main.ImageParts.Count());
        var embeds = main.Document.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().Select(b => b.Embed!.Value).ToArray();
        Assert.Equal(2, embeds.Length);
        Assert.Equal(2, embeds.Distinct().Count());
        Assert.All(embeds, id => Assert.IsAssignableFrom<ImagePart>(main.GetPartById(id)));
    }

    [Theory]
    [InlineData("revision", "unsupported-merge-content")]
    [InlineData("comment", "unsupported-merge-relationship")]
    [InlineData("field", "unsupported-merge-field")]
    [InlineData("external", "unsupported-merge-external-content")]
    [InlineData("defaults", "merge-incompatible-formatting")]
    public void Unsupported_sources_block_preview(string change, string expected)
    {
        var one = Fixture("One", "FF0000");
        var two = Edit(Fixture("Two", "0000FF"), doc =>
        {
            var main = doc.MainDocumentPart!;
            if (change == "revision") main.Document.Body!.Elements<Paragraph>().First().Append(new InsertedRun(new Run(new Text("Changed"))) { Id = "9", Author = "A" });
            if (change == "comment") main.AddNewPart<WordprocessingCommentsPart>().Comments = new Comments();
            if (change == "field") main.Document.Body!.Elements<Paragraph>().First().Append(new SimpleField(new Run(new Text("cached"))) { Instruction = "TOC" });
            if (change == "external") main.AddExternalRelationship("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", new Uri("https://example.com/image.png"));
            if (change == "defaults") main.StyleDefinitionsPart!.Styles.PrependChild(new DocDefaults(new RunPropertiesDefault(new RunPropertiesBaseStyle(new Color { Val = "123456" }))));
        });
        var preview = new OfficeAgentClient(new WordModule()).PreviewMerge(new[] { one, two });
        Assert.Null(preview.Plan);
        Assert.Contains(preview.Diagnostics, d => d.Code == expected);
    }

    [Fact]
    public void Stale_input_or_changed_options_cannot_reuse_a_plan()
    {
        var client = new OfficeAgentClient(new WordModule());
        var inputs = new[] { Fixture("One", "FF0000"), Fixture("Two", "FF0000") };
        var preview = client.PreviewMerge(inputs);
        Assert.True(preview.IsValid, Errors(preview.Diagnostics));
        var changed = new[] { inputs[0], Fixture("Changed", "FF0000") };
        var result = client.CommitMerge(preview.Plan!, changed);
        Assert.False(result.Committed);
        Assert.Null(result.Content);
        Assert.Contains(result.Diagnostics, d => d.Code == "stale-merge-source");
        Assert.Throws<ArgumentException>(() => client.CommitMerge(new DocumentMergePlan
        {
            Inputs = preview.Plan!.Inputs,
            PlanSha256 = preview.Plan.PlanSha256,
            Options = new() { Title = "Changed" }
        }, inputs));
    }

    [Fact]
    public void Host_limits_and_cancellation_apply_before_output()
    {
        var inputs = new[] { Fixture("One", "FF0000"), Fixture("Two", "FF0000") };
        var client = new OfficeAgentClient(new WordModule()) { MergeLimits = new() { MaximumExpandedBytes = 100 } };
        Assert.Contains(client.PreviewMerge(inputs).Diagnostics, d => d.Code == "merge-resource-limit");
        var normal = new OfficeAgentClient(new WordModule());
        Assert.Throws<OperationCanceledException>(() => normal.PreviewMerge(inputs, cancellationToken: new CancellationToken(true)));
        Assert.Throws<ArgumentException>(() => normal.PreviewMerge(new[] { inputs[0] }));
    }

    [Fact]
    public async Task Provider_commit_creates_a_new_output_and_round_trips_serialized_plan()
    {
        var store = new MemoryDocumentProvider("docs");
        var left = store.Add("a.docx", Fixture("One", "FF0000"));
        var right = store.Add("b.docx", Fixture("Two", "0000FF"));
        var client = new OfficeAgentClient(new DocumentProviderRegistry(new[] { store }), new WordModule());
        var preview = await client.PreviewMergeAsync(new() { Sources = new[] { left, right } });
        Assert.True(preview.IsValid, Errors(preview.Diagnostics));
        var plan = JsonSerializer.Deserialize<DocumentMergePlan>(JsonSerializer.Serialize(preview.Plan))!;
        var result = await client.CommitMergeAsync(plan, "docs", "packet.docx");
        Assert.True(result.Committed, Errors(result.Diagnostics));
        Assert.Equal(3, store.Count);
        Assert.Null(result.Content);
        Assert.Equal(result.Document!.ItemId, result.Receipt!.OutputDocument!.ItemId);
        Assert.Equal(new[] { left.ItemId, right.ItemId }, result.Receipt.Inputs.Select(i => i.Document!.ItemId));
        await Assert.ThrowsAsync<DocumentProviderException>(() => client.CommitMergeAsync(plan, "docs", "packet.docx"));
    }

    [Fact]
    public async Task Merge_plan_version_is_separate_and_rejected_before_source_read_or_output()
    {
        var omitted = JsonSerializer.Deserialize<DocumentMergePlan>("""
            {
              "inputs": [],
              "options": {},
              "planSha256": ""
            }
            """)!;
        Assert.Equal(DocumentMergePlan.CurrentVersion, omitted.Version);

        var store = new MemoryDocumentProvider("docs");
        var client = new OfficeAgentClient(new DocumentProviderRegistry(new[] { store }), new WordModule());
        var invalid = new DocumentMergePlan
        {
            Version = "2",
            Inputs = new[]
            {
                new DocumentMergeInput { Index = 0, Sha256 = new string('0', 64) },
                new DocumentMergeInput { Index = 1, Sha256 = new string('1', 64) }
            },
            PlanSha256 = new string('2', 64)
        };

        var error = await Assert.ThrowsAsync<ArgumentException>(() =>
            client.CommitMergeAsync(invalid, "docs", "packet.docx"));

        Assert.Contains(DocumentMergePlan.CurrentVersion, error.Message);
        Assert.Equal(0, store.Count);
    }

    [Fact]
    public void Documents_the_library_creates_can_be_merged()
    {
        // The blank Word document gives its heading styles an outline level, and so do
        // defineStyle and format, but the assembler did not list w:outlineLvl, so two documents
        // create_document had just produced were refused as unsupported content.
        var client = new OfficeAgentClient(new WordModule());
        var inputs = new[] { client.CreateBlank("a.docx"), client.CreateBlank("b.docx") };
        using (var blank = Open(inputs[0]))
            Assert.NotEmpty(blank.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Descendants<OutlineLevel>());

        var preview = client.PreviewMerge(inputs);
        Assert.True(preview.IsValid, Errors(preview.Diagnostics));
        var result = client.CommitMerge(preview.Plan!, inputs);
        Assert.True(result.Committed, Errors(result.Diagnostics));

        using var merged = Open(result.Content!);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2019).Validate(merged));
        Assert.Contains(merged.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Elements<Style>(),
            style => style.StyleParagraphProperties?.OutlineLevel?.Val?.Value == 0);
    }

    [Fact]
    public async Task Tool_round_trip_and_creation_gating()
    {
        var store = new MemoryDocumentProvider("docs");
        var left = store.Add("a.docx", Fixture("One", "FF0000"));
        var right = store.Add("b.docx", Fixture("Two", "0000FF"));
        var client = new OfficeAgentClient(new DocumentProviderRegistry(new[] { store }), new WordModule());
        var tools = new OfficeAgentTools(client);
        Assert.Contains(tools.AsAIFunctions(), f => f.Name == "preview_document_merge");
        Assert.DoesNotContain(tools.AsAIFunctions(), f => f.Name == "merge_documents");
        Assert.Contains(tools.AsAIFunctions(new() { AllowCreation = true }), f => f.Name == "merge_documents");
        using var preview = JsonDocument.Parse(await tools.PreviewDocumentMerge(JsonSerializer.Serialize(new
        {
            sources = new[]
            {
                new { connectionId = left.ConnectionId, documentId = left.ItemId },
                new { connectionId = right.ConnectionId, documentId = right.ItemId }
            }
        })));
        Assert.True(preview.RootElement.GetProperty("isValid").GetBoolean(), preview.RootElement.ToString());
        using var result = JsonDocument.Parse(await tools.MergeDocuments(preview.RootElement.GetProperty("plan").GetRawText(), "docs", "packet.docx"));
        Assert.True(result.RootElement.GetProperty("committed").GetBoolean(), result.RootElement.ToString());
        Assert.Equal(JsonValueKind.Null, result.RootElement.GetProperty("content").ValueKind);
    }

    [Fact]
    public void Dependency_injection_resolves_the_assembler()
    {
        var services = new ServiceCollection();
        services.AddWordFormat();
        services.AddOfficeAgent();
        using var provider = services.BuildServiceProvider();
        var preview = provider.GetRequiredService<OfficeAgentClient>().PreviewMerge(new[] { Fixture("A", "FF0000"), Fixture("B", "0000FF") });
        Assert.True(preview.IsValid, Errors(preview.Diagnostics));
    }

    internal static byte[] Fixture(string text, string color, bool landscape = false, string? header = null)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var stylePart = main.AddNewPart<StyleDefinitionsPart>();
            stylePart.Styles = new Styles(new Style(new StyleName { Val = "Clause" }, new StyleRunProperties(new Color { Val = color })) { Type = StyleValues.Paragraph, StyleId = "Clause", Default = true });
            main.AddNewPart<NumberingDefinitionsPart>().Numbering = new Numbering(
                new AbstractNum(new Level(new StartNumberingValue { Val = 1 }, new NumberingFormat { Val = NumberFormatValues.Decimal }, new LevelText { Val = "%1." }) { LevelIndex = 0 }) { AbstractNumberId = 0 },
                new NumberingInstance(new AbstractNumId { Val = 0 }) { NumberID = 1 });
            var paragraph = new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "Clause" }, new NumberingProperties(new NumberingLevelReference { Val = 0 }, new NumberingId { Val = 1 })),
                new BookmarkStart { Id = "1", Name = "Start" }, new Run(new Text(text)), new BookmarkEnd { Id = "1" });
            var table = new Table(new TableProperties(new TableWidth { Width = "4000", Type = TableWidthUnitValues.Dxa }), new TableGrid(new GridColumn { Width = "2000" }, new GridColumn { Width = "2000" }),
                new TableRow(new TableCell(new TableCellProperties(new GridSpan { Val = 2 }), new Paragraph(new Run(new Text("Table content"))))));
            var control = new SdtBlock(new SdtProperties(new SdtAlias { Val = "Customer" }, new Tag { Val = "Customer" }, new SdtId { Val = 7 }, new SdtContentText()),
                new SdtContentBlock(new Paragraph(new Hyperlink(new Run(new Text("Web"))) { Id = "link" }, new Hyperlink(new Run(new Text("Internal"))) { Anchor = "Start" })));
            main.AddHyperlinkRelationship(new Uri("https://example.com/"), true, "link");
            var section = new SectionProperties();
            if (header is not null)
            {
                var part = main.AddNewPart<HeaderPart>();
                part.Header = new Header(new Paragraph(new Run(new Text(header))));
                section.Append(new HeaderReference { Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(part) });
            }
            section.Append(new PageSize { Width = landscape ? 16838U : 11906U, Height = landscape ? 11906U : 16838U, Orient = landscape ? PageOrientationValues.Landscape : PageOrientationValues.Portrait });
            main.Document = new Document(new Body(paragraph, table, control, section));
        }
        return stream.ToArray();
    }

    private static WordprocessingDocument Open(byte[] bytes) => WordprocessingDocument.Open(new MemoryStream(bytes), false);
    private static byte[] WithImage(byte[] bytes)
    {
        var client = new OfficeAgentClient(new WordModule());
        var paragraph = client.Inspect(bytes).Paragraphs.First();
        using var applied = client.Commit(new StreamHandle(new MemoryStream(bytes, writable: false)), new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new InsertImageOp
                {
                    Target = new TextSpanAnchor { ParaId = paragraph.ParaId, Expect = paragraph.Text },
                    Base64Bytes = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=",
                    ImageType = "png", WidthPx = 16, HeightPx = 16, Position = InsertPosition.After, Mode = ChangeMode.Direct
                }
            }
        });
        Assert.True(applied.Committed);
        return OfficeAgentClient.ToBytes(applied);
    }
    private static byte[] Edit(byte[] bytes, Action<WordprocessingDocument> edit)
    {
        using var stream = new MemoryStream(); stream.Write(bytes); stream.Position = 0;
        using (var doc = WordprocessingDocument.Open(stream, true)) edit(doc);
        return stream.ToArray();
    }
    private static string Errors(IEnumerable<WorkflowDiagnostic> diagnostics) => string.Join("; ", diagnostics.Select(d => $"{d.Code}: {d.Path}: {d.Message}"));
}
