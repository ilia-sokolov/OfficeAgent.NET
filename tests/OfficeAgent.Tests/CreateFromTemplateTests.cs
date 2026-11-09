using System.IO.Compression;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeAgent.Abstractions;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Excel;
using OfficeAgent.PowerPoint;
using OfficeAgent.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.Tests;

/// <summary>
/// Creating a document from the company's template: the output starts as the template's bytes,
/// becomes an ordinary document when the template is a .dotx, .potx or .xltx, and keeps every
/// creation guarantee a blank document has.
/// </summary>
public sealed class CreateFromTemplateTests
{
    private const string WordDocument = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
    private const string WordTemplate = "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml";
    private const string Presentation = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
    private const string Workbook = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";

    // ── Templates become documents ───────────────────────────────────────

    [Fact]
    public async Task A_dotx_becomes_a_docx_with_the_plan_applied_and_the_template_untouched()
    {
        using var workspace = new Workspace();
        var templateBytes = Retyped(DocxFactory.Contract(), WordprocessingDocumentType.Template);
        var template = await workspace.RegisterAsync("letterhead.dotx", templateBytes);
        // An agent inspects the template itself for anchors; they and its snapshot hold for the copy.
        var inspected = workspace.Client.Inspect(templateBytes);
        var paragraph = inspected.Paragraphs.First(p => p.Text.Length > 0);

        var result = await workspace.Client.CreateFromTemplateAsync("workspace", "letter.docx", template, new DocumentPlan
        {
            Snapshot = inspected.Snapshot,
            Operations = new PlanOperation[]
            {
                new InsertOp
                {
                    Target = new TextSpanAnchor { ParaId = paragraph.ParaId, Expect = "" },
                    Position = InsertPosition.After, Text = "Dear Northwind,", Mode = ChangeMode.Direct
                }
            }
        });

        Assert.True(result.Committed, string.Join("; ", result.Report.Errors.Select(e => e.Message)));
        Assert.Equal("letter.docx", result.Document!.Name);
        var output = File.ReadAllBytes(Path.Combine(workspace.Root, "letter.docx"));
        Assert.Equal(WordDocument, MainContentType(output));
        Assert.Contains("Dear Northwind,", workspace.Client.Inspect(output).Paragraphs.Select(p => p.Text));
        Assert.Contains(paragraph.Text, workspace.Client.Inspect(output).Paragraphs.Select(p => p.Text));
        Assert.Equal(templateBytes, File.ReadAllBytes(Path.Combine(workspace.Root, "letterhead.dotx")));
        Assert.Equal(WordTemplate, MainContentType(templateBytes));
        Assert.Empty(SchemaErrors(output));

        // The receipt describes the document the plan was applied to, and the stored output.
        Assert.NotNull(result.Receipt);
        Assert.Equal(NativeCorpus.Sha256(output), result.Receipt!.OutputSha256);
        Assert.Equal(result.Document.ItemId, result.Receipt.OutputDocument!.ItemId);
    }

    [Fact]
    public async Task A_potx_becomes_a_pptx()
    {
        using var workspace = new Workspace();
        var templateBytes = Retyped(new PowerPointModule().CreateBlank(), PresentationDocumentType.Template);
        var template = await workspace.RegisterAsync("brand.potx", templateBytes);

        var result = await workspace.Client.CreateFromTemplateAsync("workspace", "pitch.pptx", template, new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new InsertSlideOp { Slide = new SlideData { Layout = "titleAndContent", Title = "Why Northwind", Body = new[] { "Twelve years" } } }
            }
        });

        Assert.True(result.Committed, string.Join("; ", result.Report.Errors.Select(e => e.Message)));
        var output = File.ReadAllBytes(Path.Combine(workspace.Root, "pitch.pptx"));
        Assert.Equal(Presentation, MainContentType(output));
        using var deck = PresentationDocument.Open(new MemoryStream(output), false);
        Assert.Equal(PresentationDocumentType.Presentation, deck.DocumentType);
        Assert.Equal(2, deck.PresentationPart!.Presentation!.SlideIdList!.Count());
        Assert.Equal(templateBytes, File.ReadAllBytes(Path.Combine(workspace.Root, "brand.potx")));
    }

    [Fact]
    public async Task An_xltx_becomes_an_xlsx_and_no_plan_copies_it_as_a_document()
    {
        using var workspace = new Workspace();
        var templateBytes = Retyped(new ExcelModule().CreateBlank(), SpreadsheetDocumentType.Template);
        var template = await workspace.RegisterAsync("budget.xltx", templateBytes);

        var result = await workspace.Client.CreateFromTemplateAsync("workspace", "budget-2027.xlsx", template);

        Assert.True(result.Committed);
        Assert.Null(result.Receipt);
        var output = File.ReadAllBytes(Path.Combine(workspace.Root, "budget-2027.xlsx"));
        Assert.Equal(Workbook, MainContentType(output));
        Assert.Equal(templateBytes, File.ReadAllBytes(Path.Combine(workspace.Root, "budget.xltx")));
    }

    [Fact]
    public async Task An_ordinary_document_is_copied_as_it_is()
    {
        using var workspace = new Workspace();
        var templateBytes = DocxFactory.Contract();
        var template = await workspace.RegisterAsync("master.docx", templateBytes);

        var result = await workspace.Client.CreateFromTemplateAsync("workspace", "copy.docx", template);

        Assert.True(result.Committed);
        Assert.Equal(templateBytes, File.ReadAllBytes(Path.Combine(workspace.Root, "copy.docx")));
    }

    [Fact]
    public async Task A_pptx_template_takes_the_plan()
    {
        using var workspace = new Workspace();
        var template = await workspace.RegisterAsync("deck.pptx", PptxFactory.Deck());

        var result = await workspace.Client.CreateFromTemplateAsync("workspace", "copy.pptx", template, new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new ChangeTextOp { Target = new TextSpanAnchor { ParaId = "slide256/shape2/p0", Expect = PptxFactory.TitleText }, With = "Annual Review", Mode = ChangeMode.Direct }
            }
        });

        Assert.True(result.Committed, string.Join("; ", result.Report.Errors.Select(e => e.Message)));
        var output = File.ReadAllBytes(Path.Combine(workspace.Root, "copy.pptx"));
        Assert.Contains("Annual Review", workspace.Client.Inspect(output).Paragraphs.Select(p => p.Text));
        Assert.Equal(Presentation, MainContentType(output));
    }

    // ── Refusals write nothing ───────────────────────────────────────────

    [Theory]
    [InlineData(WordprocessingDocumentType.MacroEnabledTemplate, "letterhead.dotm")]
    [InlineData(WordprocessingDocumentType.MacroEnabledDocument, "letterhead.docm")]
    public async Task A_macro_enabled_template_is_refused(WordprocessingDocumentType type, string name)
    {
        using var workspace = new Workspace(".dotm", ".docm");
        var template = await workspace.RegisterAsync(name, Retyped(DocxFactory.Contract(), type));

        await Assert.ThrowsAsync<MacroEnabledDocumentException>(() =>
            workspace.Client.CreateFromTemplateAsync("workspace", "letter.docx", template));

        Assert.False(File.Exists(Path.Combine(workspace.Root, "letter.docx")));
    }

    [Fact]
    public async Task A_vba_project_in_a_document_typed_package_is_refused()
    {
        using var workspace = new Workspace();
        var template = await workspace.RegisterAsync("sneaky.docx", WithEntry(DocxFactory.Contract(), "word/vbaProject.bin"));

        await Assert.ThrowsAsync<MacroEnabledDocumentException>(() =>
            workspace.Client.CreateFromTemplateAsync("workspace", "letter.docx", template));
        Assert.False(File.Exists(Path.Combine(workspace.Root, "letter.docx")));
    }

    [Theory]
    [InlineData("letter.pptx")]
    [InlineData("letter.dotx")]
    [InlineData("letter.xlsx")]
    public async Task A_name_in_another_format_is_refused_before_anything_is_written(string name)
    {
        using var workspace = new Workspace();
        var template = await workspace.RegisterAsync("letterhead.dotx", Retyped(DocxFactory.Contract(), WordprocessingDocumentType.Template));
        var before = Directory.GetFiles(workspace.Root).Order(StringComparer.Ordinal).ToArray();

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            workspace.Client.CreateFromTemplateAsync("workspace", name, template));

        Assert.Equal(ProviderErrorCode.InvalidArgument, error.Code);
        Assert.Contains(".docx", error.Message, StringComparison.Ordinal);
        Assert.Equal(before, Directory.GetFiles(workspace.Root).Order(StringComparer.Ordinal).ToArray());
    }

    [Fact]
    public async Task An_existing_name_is_never_overwritten()
    {
        using var workspace = new Workspace();
        var template = await workspace.RegisterAsync("letterhead.dotx", Retyped(DocxFactory.Contract(), WordprocessingDocumentType.Template));
        File.WriteAllBytes(Path.Combine(workspace.Root, "letter.docx"), new byte[] { 1, 2, 3 });

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            workspace.Client.CreateFromTemplateAsync("workspace", "letter.docx", template));

        Assert.Equal(ProviderErrorCode.AlreadyExists, error.Code);
        Assert.Equal(new byte[] { 1, 2, 3 }, File.ReadAllBytes(Path.Combine(workspace.Root, "letter.docx")));
    }

    [Fact]
    public async Task A_failing_plan_leaves_nothing_behind()
    {
        using var workspace = new Workspace();
        var templateBytes = Retyped(DocxFactory.Contract(), WordprocessingDocumentType.Template);
        var template = await workspace.RegisterAsync("letterhead.dotx", templateBytes);
        var before = Directory.GetFiles(workspace.Root).Order(StringComparer.Ordinal).ToArray();

        var result = await workspace.Client.CreateFromTemplateAsync("workspace", "letter.docx", template, new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new InsertOp { Target = new TextSpanAnchor { ParaId = "w14:DEADBEEF", Expect = "" }, Text = "Never written" }
            }
        });

        Assert.False(result.Committed);
        Assert.Null(result.Document);
        Assert.Contains(result.Report.Errors, e => e.Code == ValidationErrorCodes.AnchorNotFound);
        Assert.Equal(before, Directory.GetFiles(workspace.Root).Order(StringComparer.Ordinal).ToArray());
        Assert.Equal(templateBytes, File.ReadAllBytes(Path.Combine(workspace.Root, "letterhead.dotx")));
    }

    [Fact]
    public async Task The_template_can_come_from_another_connection()
    {
        using var workspace = new Workspace();
        var templates = workspace.Root + "-templates";
        Directory.CreateDirectory(templates);
        try
        {
            var library = new FileSystemDocumentProvider(new FileSystemDocumentProviderOptions
            {
                ConnectionId = "templates", RootPath = templates, AllowedExtensions = new[] { ".dotx" }
            });
            var client = new OfficeAgentClient(
                new DocumentProviderRegistry(new IDocumentProvider[] { workspace.Provider, library }), new WordModule());
            var path = Path.Combine(templates, "letterhead.dotx");
            File.WriteAllBytes(path, Retyped(DocxFactory.Contract(), WordprocessingDocumentType.Template));
            var template = await client.RegisterAsync("templates", path);

            var result = await client.CreateFromTemplateAsync("workspace", "letter.docx", template);

            Assert.True(result.Committed);
            Assert.Equal(WordDocument, MainContentType(File.ReadAllBytes(Path.Combine(workspace.Root, "letter.docx"))));
            Assert.Single(Directory.GetFiles(templates));
        }
        finally
        {
            Directory.Delete(templates, recursive: true);
        }
    }

    // ── The create_document tool ─────────────────────────────────────────

    [Fact]
    public async Task The_tool_creates_from_a_template_source()
    {
        using var workspace = new Workspace();
        File.WriteAllBytes(Path.Combine(workspace.Root, "letterhead.dotx"), Retyped(DocxFactory.Contract(), WordprocessingDocumentType.Template));
        var tools = new OfficeAgentTools(workspace.Client);

        using var created = JsonDocument.Parse(await tools.CreateDocument(
            "workspace", "letter.docx", "", templateSource: Path.Combine(workspace.Root, "letterhead.dotx")));

        Assert.True(created.RootElement.GetProperty("committed").GetBoolean(), created.RootElement.ToString());
        Assert.Equal("letter.docx", created.RootElement.GetProperty("outputName").GetString());
        Assert.Equal(WordDocument, MainContentType(File.ReadAllBytes(Path.Combine(workspace.Root, "letter.docx"))));
    }

    [Fact]
    public async Task The_tool_reports_a_mismatched_name_and_a_macro_template_by_code()
    {
        using var workspace = new Workspace(".dotm");
        File.WriteAllBytes(Path.Combine(workspace.Root, "letterhead.dotx"), Retyped(DocxFactory.Contract(), WordprocessingDocumentType.Template));
        File.WriteAllBytes(Path.Combine(workspace.Root, "macro.dotm"), Retyped(DocxFactory.Contract(), WordprocessingDocumentType.MacroEnabledTemplate));
        var tools = new OfficeAgentTools(workspace.Client);

        using var mismatched = JsonDocument.Parse(await tools.CreateDocument(
            "workspace", "letter.pptx", "", templateSource: Path.Combine(workspace.Root, "letterhead.dotx")));
        using var macro = JsonDocument.Parse(await tools.CreateDocument(
            "workspace", "letter.docx", "", templateSource: Path.Combine(workspace.Root, "macro.dotm")));
        using var connectionOnly = JsonDocument.Parse(await tools.CreateDocument(
            "workspace", "letter.docx", "", templateConnectionId: "workspace"));

        Assert.Equal(ToolErrorCodes.InvalidArgument, mismatched.RootElement.GetProperty("errors")[0].GetProperty("code").GetString());
        Assert.Contains(".docx", mismatched.RootElement.GetProperty("errors")[0].GetProperty("message").GetString());
        Assert.Equal(MacroEnabledDocumentException.Code, macro.RootElement.GetProperty("errors")[0].GetProperty("code").GetString());
        Assert.Equal(ToolErrorCodes.InvalidArgument, connectionOnly.RootElement.GetProperty("errors")[0].GetProperty("code").GetString());
        Assert.False(File.Exists(Path.Combine(workspace.Root, "letter.docx")));
        Assert.False(File.Exists(Path.Combine(workspace.Root, "letter.pptx")));
    }

    // ── Template batches from .dotx and .potx ────────────────────────────

    [Fact]
    public async Task A_dotx_template_batch_produces_docx_documents()
    {
        using var workspace = new Workspace();
        var templateBytes = Retyped(QuoteTemplate(), WordprocessingDocumentType.Template);
        var template = await workspace.RegisterAsync("quote.dotx", templateBytes);
        var request = new TemplateBatchRequest
        {
            Items = new[]
            {
                new TemplateBatchItem
                {
                    OutputName = "quote-fabrikam.docx",
                    Binding = new TemplateBinding { Values = new Dictionary<string, string?> { ["CustomerName"] = "Fabrikam" }, Mode = ChangeMode.Direct }
                }
            }
        };

        var discovered = await workspace.Client.DiscoverTemplateAsync(template);
        var preview = await workspace.Client.PreviewTemplateBatchAsync(template, request);
        var result = await workspace.Client.PopulateTemplateBatchAsync(template, request, preview.Token);

        Assert.Contains(discovered.Slots, slot => slot.Name == "CustomerName");
        Assert.Equal(NativeCorpus.Sha256(templateBytes), discovered.TemplateSha256);
        Assert.True(preview.IsValid);
        Assert.True(result.Committed, string.Join("; ", result.Items.SelectMany(i => i.Diagnostics).Select(d => d.Message)));
        var output = File.ReadAllBytes(Path.Combine(workspace.Root, "quote-fabrikam.docx"));
        Assert.Equal(WordDocument, MainContentType(output));
        Assert.Contains(workspace.Client.Inspect(output).Paragraphs, p => p.Text.Contains("Fabrikam", StringComparison.Ordinal));
        Assert.Equal(templateBytes, File.ReadAllBytes(Path.Combine(workspace.Root, "quote.dotx")));
    }

    [Fact]
    public async Task A_potx_template_batch_produces_pptx_documents()
    {
        using var workspace = new Workspace();
        var templateBytes = Retyped(NamedShapeDeck(), PresentationDocumentType.Template);
        var template = await workspace.RegisterAsync("pitch.potx", templateBytes);

        var result = await workspace.Client.PopulateTemplateBatchAsync(template, new TemplateBatchRequest
        {
            Items = new[]
            {
                new TemplateBatchItem
                {
                    OutputName = "pitch-contoso.pptx",
                    Binding = new TemplateBinding { Values = new Dictionary<string, string?> { ["ClientName"] = "Contoso" }, MissingValueBehavior = MissingTemplateValueBehavior.Ignore, Mode = ChangeMode.Direct }
                }
            }
        });

        Assert.True(result.Committed, string.Join("; ", result.Items.SelectMany(i => i.Diagnostics).Select(d => d.Message)));
        var output = File.ReadAllBytes(Path.Combine(workspace.Root, "pitch-contoso.pptx"));
        Assert.Equal(Presentation, MainContentType(output));
        Assert.Contains(workspace.Client.Inspect(output).Paragraphs, p => p.Text.Contains("Contoso", StringComparison.Ordinal));
    }

    [Fact]
    public async Task A_batch_output_named_in_another_format_is_refused_in_preflight()
    {
        using var workspace = new Workspace();
        var template = await workspace.RegisterAsync("quote.dotx", Retyped(QuoteTemplate(), WordprocessingDocumentType.Template));
        var before = Directory.GetFiles(workspace.Root).Order(StringComparer.Ordinal).ToArray();

        var result = await workspace.Client.PopulateTemplateBatchAsync(template, new TemplateBatchRequest
        {
            Items = new[]
            {
                new TemplateBatchItem
                {
                    OutputName = "quote.dotx.pptx",
                    Binding = new TemplateBinding { Values = new Dictionary<string, string?> { ["CustomerName"] = "Fabrikam" } }
                }
            }
        });

        Assert.False(result.Committed);
        Assert.Contains(result.Items.Single().Diagnostics, d => d.Code == TemplateDiagnosticCodes.OutputFormatMismatch);
        Assert.Equal(before, Directory.GetFiles(workspace.Root).Order(StringComparer.Ordinal).ToArray());
    }

    [Fact]
    public async Task A_macro_enabled_template_is_refused_by_discovery()
    {
        using var workspace = new Workspace(".dotm");
        var template = await workspace.RegisterAsync("quote.dotm", Retyped(QuoteTemplate(), WordprocessingDocumentType.MacroEnabledTemplate));

        await Assert.ThrowsAsync<MacroEnabledDocumentException>(() => workspace.Client.DiscoverTemplateAsync(template));
    }

    // ── Fixtures ─────────────────────────────────────────────────────────

    /// <summary>The package's own type changed, as Word or PowerPoint does on Save As.</summary>
    internal static byte[] Retyped(byte[] document, WordprocessingDocumentType type)
    {
        using var stream = new MemoryStream();
        stream.Write(document);
        using (var package = WordprocessingDocument.Open(stream, true)) package.ChangeDocumentType(type);
        return stream.ToArray();
    }

    internal static byte[] Retyped(byte[] document, PresentationDocumentType type)
    {
        using var stream = new MemoryStream();
        stream.Write(document);
        using (var package = PresentationDocument.Open(stream, true)) package.ChangeDocumentType(type);
        return stream.ToArray();
    }

    private static byte[] Retyped(byte[] document, SpreadsheetDocumentType type)
    {
        using var stream = new MemoryStream();
        stream.Write(document);
        using (var package = SpreadsheetDocument.Open(stream, true)) package.ChangeDocumentType(type);
        return stream.ToArray();
    }

    /// <summary>
    /// The content type <c>[Content_Types].xml</c> gives the part the package relationship
    /// names as the office document: what Office reads to decide what the file is.
    /// </summary>
    internal static string MainContentType(byte[] package)
    {
        using var zip = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        System.Xml.Linq.XDocument Load(string entry)
        {
            using var reader = new StreamReader(zip.GetEntry(entry)!.Open());
            return System.Xml.Linq.XDocument.Parse(reader.ReadToEnd());
        }

        var main = "/" + Load("_rels/.rels").Root!.Elements()
            .Single(r => ((string)r.Attribute("Type")!).EndsWith("/officeDocument", StringComparison.Ordinal))
            .Attribute("Target")!.Value.TrimStart('/');
        var types = Load("[Content_Types].xml").Root!.Elements().ToList();
        var overridden = types.SingleOrDefault(t => string.Equals((string?)t.Attribute("PartName"), main, StringComparison.OrdinalIgnoreCase));
        return (string)(overridden ?? types.Single(t => string.Equals(
            (string?)t.Attribute("Extension"), Path.GetExtension(main).TrimStart('.'), StringComparison.OrdinalIgnoreCase))).Attribute("ContentType")!;
    }

    private static byte[] WithEntry(byte[] package, string entry)
    {
        using var stream = new MemoryStream();
        stream.Write(package);
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        using (var writer = zip.CreateEntry(entry).Open())
            writer.Write(new byte[] { 0xD0, 0xCF, 0x11, 0xE0 });
        return stream.ToArray();
    }

    private static string[] SchemaErrors(byte[] docx)
    {
        using var document = WordprocessingDocument.Open(new MemoryStream(docx), false);
        return new OpenXmlValidator(FileFormatVersions.Office2019).Validate(document).Select(e => e.Description).ToArray();
    }

    private static byte[] QuoteTemplate()
    {
        using var ms = new MemoryStream();
        using (var document = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new W.Document(new W.Body(
                new W.Paragraph(new W.Run(new W.Text("Quote for ") { Space = SpaceProcessingModeValues.Preserve }),
                    new W.SdtRun(
                        new W.SdtProperties(new W.Tag { Val = "CustomerName" }, new W.SdtId { Val = 41 }),
                        new W.SdtContentRun(new W.Run(new W.Text("CUSTOMER")))))));
            main.Document.Save();
        }
        return ms.ToArray();
    }

    /// <summary>A real deck with a shape named as a template slot.</summary>
    private static byte[] NamedShapeDeck()
    {
        var client = new OfficeAgentClient(new PowerPointModule());
        var inserted = client.Commit(new StreamHandle(new MemoryStream(new PowerPointModule().CreateBlank()), "t.pptx"), new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new InsertShapeOp { Target = new NodeAnchor { Kind = "slide", Path = "slide#256" }, Text = new[] { "[CLIENT]" }, XPx = 40, YPx = 300, WidthPx = 400, HeightPx = 60 }
            }
        });
        Assert.True(inserted.Committed);
        // A template names its slot shapes in the Selection Pane; this does the same.
        using var stream = new MemoryStream();
        stream.Write(inserted.ToBytes());
        using (var deck = PresentationDocument.Open(stream, true))
            foreach (var shape in deck.PresentationPart!.SlideParts.SelectMany(s => s.Slide!.Descendants<DocumentFormat.OpenXml.Presentation.Shape>()))
                if (shape.InnerText.Contains("[CLIENT]", StringComparison.Ordinal))
                    shape.NonVisualShapeProperties!.NonVisualDrawingProperties!.Name = "ClientName";
        return stream.ToArray();
    }

    private sealed class Workspace : IDisposable
    {
        public Workspace(params string[] extraExtensions)
        {
            Root = Path.Combine(Path.GetTempPath(), $"officeagent-template-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Root);
            Provider = new FileSystemDocumentProvider(new FileSystemDocumentProviderOptions
            {
                ConnectionId = "workspace",
                RootPath = Root,
                AllowedExtensions = new[] { ".docx", ".pptx", ".xlsx", ".dotx", ".potx", ".xltx" }.Concat(extraExtensions).ToArray()
            });
            Client = new OfficeAgentClient(
                new DocumentProviderRegistry(new IDocumentProvider[] { Provider }),
                new WordModule(), new PowerPointModule(), new ExcelModule());
        }

        public string Root { get; }
        public FileSystemDocumentProvider Provider { get; }
        public OfficeAgentClient Client { get; }

        public async Task<DocumentReference> RegisterAsync(string name, byte[] content)
        {
            var path = Path.Combine(Root, name);
            await File.WriteAllBytesAsync(path, content);
            return await Client.RegisterAsync("workspace", path);
        }

        public void Dispose()
        {
            if (Directory.Exists(Root)) Directory.Delete(Root, recursive: true);
        }
    }
}
