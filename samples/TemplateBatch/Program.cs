using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Word;

var outputDirectory = Path.GetFullPath(args.Length == 0 ? "generated-quotes" : args[0]);
Directory.CreateDirectory(outputDirectory);
var storageRoot = Path.Combine(Path.GetTempPath(), $"officeagent-template-{Guid.NewGuid():N}");
Directory.CreateDirectory(storageRoot);

try
{
    var templatePath = Path.Combine(storageRoot, "quote-template.docx");
    CreateQuoteTemplate(templatePath);
    var provider = new FileSystemDocumentProvider(new FileSystemDocumentProviderOptions
    {
        ConnectionId = "workspace",
        RootPath = storageRoot,
        DefaultChangeMode = ChangeMode.Direct
    });
    var client = new OfficeAgentClient(new DocumentProviderRegistry(new[] { provider }), new WordModule());
    var template = await client.RegisterAsync("workspace", templatePath);
    var result = await client.PopulateTemplateBatchAsync(template, new TemplateBatchRequest
    {
        Items = new[]
        {
            Quote("quote-acme.docx", "Acme BV", ("Discovery workshop", "1", "1200.00"), ("Support", "4", "150.00")),
            Quote("quote-contoso.docx", "Contoso Ltd", ("Migration assessment", "2", "950.00"))
        }
    });

    foreach (var item in result.Items)
    {
        if (!item.Committed || item.Document is null)
        {
            Console.Error.WriteLine($"{item.OutputName}: {string.Join("; ", item.Diagnostics.Select(d => $"{d.Code}: {d.Message}"))}");
            continue;
        }
        using var saved = await client.OpenReadAsync(item.Document);
        var outputPath = Path.Combine(outputDirectory, item.OutputName);
        await using var output = File.Create(outputPath);
        await saved.Stream.CopyToAsync(output);
        var outputHash = item.Receipt?.OutputSha256 ?? string.Empty;
        Console.WriteLine($"Wrote {outputPath} (receipt {item.Receipt?.Outcome}, {outputHash[..Math.Min(12, outputHash.Length)]}...).");
    }
    return result.Committed ? 0 : 2;
}
finally
{
    if (Directory.Exists(storageRoot)) Directory.Delete(storageRoot, recursive: true);
}

static TemplateBatchItem Quote(string outputName, string customer, params (string Description, string Quantity, string Price)[] lines) => new()
{
    OutputName = outputName,
    Binding = new TemplateBinding
    {
        Values = new Dictionary<string, string?> { ["CustomerName"] = customer },
        RepeatingTables = new[]
        {
            new RepeatingTableBinding
            {
                TablePath = "table#0",
                TemplateRowIndex = 1,
                Records = lines.Select(line => (IReadOnlyDictionary<string, string?>)new Dictionary<string, string?>
                {
                    ["Description"] = line.Description,
                    ["Quantity"] = line.Quantity,
                    ["Price"] = line.Price
                }).ToArray()
            }
        }
    }
};

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
            Row("Description", "Quantity", "Price"),
            Row("{{Description}}", "{{Quantity}}", "{{Price}}")),
        new Paragraph()));
    main.Document.Save();
}

static TableRow Row(params string[] values) => new(values.Select(value =>
    new TableCell(new Paragraph(new Run(new Text(value))))));
