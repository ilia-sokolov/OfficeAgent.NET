using System.Security.Cryptography;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Word;

var root = Path.Combine(Path.GetTempPath(), $"officeagent-integration-{Guid.NewGuid():N}");
Directory.CreateDirectory(root);

try
{
    await VerifyTrackedEditAndRefusalAsync(root);
    await VerifyTemplatePopulationAsync(root);
    await VerifyCompleteComparisonAsync(root);
    Console.WriteLine("all-recipes=passed");
    return 0;
}
finally
{
    if (Directory.Exists(root))
        Directory.Delete(root, recursive: true);
}

static async Task VerifyTrackedEditAndRefusalAsync(string root)
{
    var path = Path.Combine(root, "proposal.docx");
    CreateDocument(path, "Prepared for Northwind Labs.");
    var bytes = await File.ReadAllBytesAsync(path);
    using var input = new MemoryStream(bytes, writable: false);
    var handle = new StreamHandle(input, "proposal.docx");
    var client = new OfficeAgentClient(new WordModule());

    var inspection = await client.InspectAsync(handle);
    var hit = (await client.FindAsync(handle, new FindQuery("Northwind Labs"))).Single();
    var plan = new DocumentPlan
    {
        Snapshot = inspection.Snapshot,
        Operations = new PlanOperation[]
        {
            new ChangeTextOp
            {
                Target = hit.Anchor,
                With = "Contoso Research",
                Mode = ChangeMode.Tracked
            }
        }
    };
    var preview = await client.PreviewAsync(handle, plan);
    Require(preview.IsValid, "tracked edit preview must be valid");
    Require(preview.Changes.Single().Before == "Northwind Labs", "tracked preview before text");
    Require(preview.Changes.Single().After == "Contoso Research", "tracked preview after text");

    using var result = await client.CommitAsync(handle, plan);
    Require(result.Committed, FormatErrors(result.Report.Errors));
    var output = Path.Combine(root, "proposal-edited.docx");
    await result.SaveAsync(output);
    using (var document = WordprocessingDocument.Open(output, isEditable: false))
    {
        var body = document.MainDocumentPart?.Document?.Body
            ?? throw new InvalidOperationException("tracked output has no body");
        Require(VisibleText(body).Contains("Contoso Research", StringComparison.Ordinal), "tracked output visible text");
        Require(string.Concat(body.Descendants<DeletedText>().Select(text => text.Text)).Contains("Northwind Labs", StringComparison.Ordinal), "tracked output deleted text");
        ValidateSchema(document, "tracked output");
    }
    Console.WriteLine("tracked-edit=passed");

    var beforeHash = Convert.ToHexString(SHA256.HashData(bytes));
    var unsupported = await client.PreviewAsync(handle, new DocumentPlan
    {
        Snapshot = inspection.Snapshot,
        Operations = new PlanOperation[]
        {
            new SetCellOp
            {
                Target = new CellAnchor { SheetId = 1, Address = "A1" },
                Value = "Not valid for Word"
            }
        }
    });
    var code = unsupported.Errors.Single().Code;
    var afterHash = Convert.ToHexString(SHA256.HashData(bytes));
    Require(!unsupported.IsValid, "unsupported operation must fail preview");
    Require(code == "unsupported-operation", "unsupported operation code");
    Require(beforeHash == afterHash, "refused preview must not mutate input bytes");
    Console.WriteLine($"refusal={code} bytes-unchanged={beforeHash == afterHash}");
}

static async Task VerifyTemplatePopulationAsync(string root)
{
    var storage = Path.Combine(root, "template-storage");
    Directory.CreateDirectory(storage);
    var templatePath = Path.Combine(storage, "quote-template.docx");
    CreateQuoteTemplate(templatePath);
    var client = CreateStoredClient(storage);
    var template = await client.RegisterAsync("workspace", templatePath);
    var result = await client.PopulateTemplateBatchAsync(template, new TemplateBatchRequest
    {
        Items = new[]
        {
            new TemplateBatchItem
            {
                OutputName = "fabrikam-quote.docx",
                Binding = new TemplateBinding
                {
                    Values = new Dictionary<string, string?> { ["CustomerName"] = "Fabrikam Services" },
                    RepeatingTables = new[]
                    {
                        new RepeatingTableBinding
                        {
                            TablePath = "table#0",
                            TemplateRowIndex = 1,
                            Records = new IReadOnlyDictionary<string, string?>[]
                            {
                                new Dictionary<string, string?>
                                {
                                    ["Description"] = "Architecture workshop",
                                    ["Quantity"] = "1",
                                    ["Price"] = "1200.00"
                                },
                                new Dictionary<string, string?>
                                {
                                    ["Description"] = "Implementation review",
                                    ["Quantity"] = "2",
                                    ["Price"] = "600.00"
                                }
                            }
                        }
                    }
                }
            }
        }
    });
    var item = result.Items.Single();
    Require(result.Committed && item.Committed && item.Document is not null,
        string.Join("; ", item.Diagnostics.Select(diagnostic => $"{diagnostic.Code}: {diagnostic.Message}")));
    var output = Path.Combine(root, "fabrikam-quote.docx");
    using (var saved = await client.OpenReadAsync(item.Document!))
    await using (var file = File.Create(output))
        await saved.Stream.CopyToAsync(file);
    using (var document = WordprocessingDocument.Open(output, isEditable: false))
    {
        var body = document.MainDocumentPart?.Document?.Body
            ?? throw new InvalidOperationException("template output has no body");
        var text = VisibleText(body);
        Require(text.Contains("Fabrikam Services", StringComparison.Ordinal), "template scalar value");
        Require(text.Contains("Architecture workshop", StringComparison.Ordinal), "first repeated record");
        Require(text.Contains("Implementation review", StringComparison.Ordinal), "second repeated record");
        Require(!text.Contains("{{", StringComparison.Ordinal), "template markers must be resolved");
        ValidateSchema(document, "template output");
    }
    Console.WriteLine("template-population=passed");
}

static async Task VerifyCompleteComparisonAsync(string root)
{
    var storage = Path.Combine(root, "comparison-storage");
    Directory.CreateDirectory(storage);
    var originalPath = Path.Combine(storage, "original.docx");
    var revisedPath = Path.Combine(storage, "revised.docx");
    CreateDocument(originalPath, "Services agreement", "Invoices are due in 30 days.", "Signed for Adventure Works.");
    CreateDocument(revisedPath, "Services agreement", "Invoices are due in 45 days.", "Late payments require written notice.", "Signed for Adventure Works.");
    var client = CreateStoredClient(storage);
    var original = await client.RegisterAsync("workspace", originalPath);
    var revised = await client.RegisterAsync("workspace", revisedPath);
    var comparison = await client.CompareDocumentsAsync(original, revised, new DocumentComparisonOptions
    {
        Revision = new RevisionMetadata { Author = "OfficeAgent Integration Recipe" }
    });
    Require(comparison.IsComplete && comparison.Plan is not null,
        string.Join("; ", comparison.Diagnostics.Select(diagnostic => $"{diagnostic.Code}: {diagnostic.Message}")));
    var preview = await client.PreviewAsync(original, comparison.Plan!);
    Require(preview.IsValid, FormatErrors(preview.Errors));
    var commit = await client.CommitAsync(original, comparison.Plan!, new SaveDocumentOptions
    {
        Mode = SaveMode.NewDocument,
        NewName = "comparison-redline.docx"
    });
    Require(commit.Committed && commit.Document is not null, FormatErrors(commit.Report.Errors));
    var output = Path.Combine(root, "comparison-redline.docx");
    using (var saved = await client.OpenReadAsync(commit.Document!))
    await using (var file = File.Create(output))
        await saved.Stream.CopyToAsync(file);
    using (var document = WordprocessingDocument.Open(output, isEditable: false))
    {
        var body = document.MainDocumentPart?.Document?.Body
            ?? throw new InvalidOperationException("comparison output has no body");
        Require(VisibleText(body).Contains("Invoices are due in 45 days.", StringComparison.Ordinal), "comparison revised text");
        Require(string.Concat(body.Descendants<DeletedText>().Select(text => text.Text)).Contains("Invoices are due in 30 days.", StringComparison.Ordinal), "comparison deleted text");
        Require(body.Descendants<InsertedRun>().Any(), "comparison must contain inserted revisions");
        ValidateSchema(document, "comparison output");
    }
    Console.WriteLine("complete-comparison=passed");
}

static OfficeAgentClient CreateStoredClient(string root) => new(
    new DocumentProviderRegistry(new[]
    {
        new FileSystemDocumentProvider(new FileSystemDocumentProviderOptions
        {
            ConnectionId = "workspace",
            RootPath = root,
            DefaultChangeMode = ChangeMode.Direct
        })
    }),
    new WordModule());

static void CreateDocument(string path, params string[] paragraphs)
{
    using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
    var main = document.AddMainDocumentPart();
    main.Document = new Document(new Body(paragraphs.Select(text => new Paragraph(new Run(new Text(text))))));
    main.Document.Save();
}

static void CreateQuoteTemplate(string path)
{
    using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
    var main = document.AddMainDocumentPart();
    var customer = new SdtRun(
        new SdtProperties(new Tag { Val = "CustomerName" }, new SdtId { Val = 10 }),
        new SdtContentRun(new Run(new Text("CUSTOMER"))));
    main.Document = new Document(new Body(
        new Paragraph(new Run(new Text("Quote for ")), customer),
        new Table(
            new TableProperties(new TableStyle { Val = "TableGrid" }),
            new TableGrid(
                new GridColumn { Width = "3600" },
                new GridColumn { Width = "1200" },
                new GridColumn { Width = "1800" }),
            Row("Description", "Quantity", "Price"),
            Row("{{Description}}", "{{Quantity}}", "{{Price}}")),
        new Paragraph()));
    main.Document.Save();
}

static TableRow Row(params string[] values) => new(values.Select(value =>
    new TableCell(new Paragraph(new Run(new Text(value))))));

static string VisibleText(OpenXmlElement body) => string.Concat(body.Descendants<Text>().Select(text => text.Text));

static void ValidateSchema(WordprocessingDocument document, string label)
{
    var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(document).ToList();
    Require(errors.Count == 0, $"{label} schema errors: {string.Join("; ", errors.Select(error => error.Description))}");
}

static string FormatErrors(IEnumerable<ValidationError> errors) =>
    string.Join("; ", errors.Select(error => $"{error.Code}: {error.Message}"));

static void Require(bool condition, string message)
{
    if (!condition)
        throw new InvalidOperationException(message);
}
