using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Word;

if (args.Length != 3 && !(args.Length == 2 && args[0] == "--demo"))
{
    Console.Error.WriteLine("Usage: DocumentComparison <original.docx> <revised.docx> <redline.docx>");
    Console.Error.WriteLine("   or: DocumentComparison --demo <redline.docx>");
    return 1;
}

var storageRoot = Path.Combine(Path.GetTempPath(), $"officeagent-compare-{Guid.NewGuid():N}");
Directory.CreateDirectory(storageRoot);
var demo = args.Length == 2;
var outputPath = Path.GetFullPath(demo ? args[1] : args[2]);
var originalPath = demo ? Path.Combine(storageRoot, "demo-original.docx") : Path.GetFullPath(args[0]);
var revisedPath = demo ? Path.Combine(storageRoot, "demo-revised.docx") : Path.GetFullPath(args[1]);

try
{
    if (demo)
    {
        CreateDocument(originalPath, "Services agreement", "Invoices are due in 30 days.", "Signed for Acme BV.");
        CreateDocument(revisedPath, "Services agreement", "Invoices are due in 45 days.",
            "Late payments require written notice.", "Signed for Acme BV.");
    }
    var stagedOriginal = Path.Combine(storageRoot, "original.docx");
    var stagedRevised = Path.Combine(storageRoot, "revised.docx");
    File.Copy(originalPath, stagedOriginal);
    File.Copy(revisedPath, stagedRevised);
    var provider = new FileSystemDocumentProvider(new FileSystemDocumentProviderOptions
    {
        ConnectionId = "workspace",
        RootPath = storageRoot
    });
    var client = new OfficeAgentClient(new DocumentProviderRegistry(new[] { provider }), new WordModule());
    var original = await client.RegisterAsync("workspace", stagedOriginal);
    var revised = await client.RegisterAsync("workspace", stagedRevised);
    var comparison = await client.CompareDocumentsAsync(original, revised, new DocumentComparisonOptions
    {
        Revision = new RevisionMetadata { Author = "OfficeAgent Compare" }
    });

    Console.WriteLine($"Original SHA-256: {comparison.OriginalSha256}");
    Console.WriteLine($"Revised SHA-256:  {comparison.RevisedSha256}");
    foreach (var difference in comparison.Differences)
        Console.WriteLine($"{difference.Kind}: {difference.Before ?? "<none>"} -> {difference.After ?? "<none>"}");
    if (!comparison.IsComplete || comparison.Plan is null)
    {
        foreach (var diagnostic in comparison.Diagnostics)
            Console.Error.WriteLine($"{diagnostic.Code}: {diagnostic.Message}");
        return 2;
    }

    var preview = await client.PreviewAsync(original, comparison.Plan);
    if (!preview.IsValid)
    {
        foreach (var error in preview.Errors) Console.Error.WriteLine($"{error.Code}: {error.Message}");
        return 3;
    }
    var commit = await client.CommitAsync(original, comparison.Plan, new SaveDocumentOptions
    {
        Mode = SaveMode.NewDocument,
        NewName = "comparison-redline.docx"
    });
    if (!commit.Committed || commit.Document is null) return 4;
    Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);
    using var saved = await client.OpenReadAsync(commit.Document);
    await using var output = File.Create(outputPath);
    await saved.Stream.CopyToAsync(output);
    Console.WriteLine($"Wrote tracked comparison to {outputPath}.");
    return 0;
}
finally
{
    if (Directory.Exists(storageRoot)) Directory.Delete(storageRoot, recursive: true);
}

static void CreateDocument(string path, params string[] paragraphs)
{
    using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
    var main = document.AddMainDocumentPart();
    main.Document = new Document(new Body(paragraphs.Select(text =>
        new Paragraph(new Run(new Text(text))))));
    main.Document.Save();
}
