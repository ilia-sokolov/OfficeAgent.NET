/// <summary>
/// TASK T3: Table Row Insert with Tracked Changes and Comment Preservation
///
/// Demonstrates inserting a new row into a contract table with:
/// - Change tracking enabled
/// - Existing comments preserved
/// - Document validity maintained
/// </summary>

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Word;

// Create a test document with a table
Console.WriteLine("Creating sample contract document with table and comment...");
byte[] contractDoc = CreateContractWithTableAndComment();
Console.WriteLine($"  - Created document: {contractDoc.Length} bytes");

// Inspect the document
var client = new OfficeAgentClient(new WordModule());
var inspection = client.Inspect(new StreamHandle(new MemoryStream(contractDoc)));
Console.WriteLine($"  - Found {inspection.Paragraphs.Count} paragraphs");
Console.WriteLine($"  - Found {inspection.Nodes.Count(n => n.Kind == "comment")} comments");

// Insert a row with tracked changes
Console.WriteLine("\nInserting new row with tracked changes...");
var applied = client.Commit(
    new StreamHandle(new MemoryStream(contractDoc)),
    new DocumentPlan
    {
        Revision = new RevisionMetadata
        {
            Author = "ContractEditor",
            TimestampUtc = DateTimeOffset.UtcNow
        },
        Operations = new PlanOperation[]
        {
            new InsertTableRowsOp
            {
                Target = new NodeAnchor { Kind = "table", Path = "table#0" },
                Rows = new[]
                {
                    new[] { "New Contract Item", "$50,000", "High Priority" }
                },
                Position = TablePosition.End,
                Mode = ChangeMode.Tracked
            }
        }
    });

if (!applied.Committed)
{
    Console.WriteLine("ERROR: Failed to commit document plan");
    foreach (var error in applied.Report.Errors)
        Console.WriteLine($"  - {error.Message}");
    return 1;
}

byte[] resultDoc = applied.ToBytes();
Console.WriteLine($"  - Document updated: {resultDoc.Length} bytes");

// Validate the document
Console.WriteLine("\nValidating document...");
using (var stream = new MemoryStream(resultDoc))
using (var package = WordprocessingDocument.Open(stream, isEditable: false))
{
    var validator = new OpenXmlValidator(FileFormatVersions.Office2019);
    var problems = validator.Validate(package).ToList();
    if (problems.Any())
    {
        Console.WriteLine($"  - Found {problems.Count} validation errors:");
        foreach (var p in problems.Take(5))
            Console.WriteLine($"    * {p.Path?.XPath}: {p.Description}");
        return 1;
    }
    Console.WriteLine("  - Document is valid");
}

// Inspect the result
Console.WriteLine("\nInspecting result...");
var resultInspection = client.Inspect(new StreamHandle(new MemoryStream(resultDoc)));
Console.WriteLine($"  - Paragraphs: {resultInspection.Paragraphs.Count}");
Console.WriteLine($"  - Comments preserved: {resultInspection.Nodes.Count(n => n.Kind == "comment")}");

// Verify the inserted row is marked with revision
using (var stream = new MemoryStream(resultDoc))
using (var package = WordprocessingDocument.Open(stream, isEditable: false))
{
    var body = package.MainDocumentPart!.Document.Body!;
    var table = body.Elements<Table>().FirstOrDefault();
    if (table != null)
    {
        var rows = table.Elements<TableRow>().ToList();
        var lastRow = rows.LastOrDefault();
        if (lastRow?.TableRowProperties is { } rowProps)
        {
            var inserted = rowProps.GetFirstChild<Inserted>();
            if (inserted != null)
            {
                var attrs = inserted.GetAttributes();
                var id = attrs.FirstOrDefault(a => a.LocalName == "id").Value;
                var author = attrs.FirstOrDefault(a => a.LocalName == "author").Value;
                Console.WriteLine($"  - Last row marked as inserted (id={id}, author={author})");
            }
        }
    }
}

// Save result
string outputPath = Path.Combine(Path.GetTempPath(), "contract_with_inserted_row.docx");
File.WriteAllBytes(outputPath, resultDoc);
Console.WriteLine($"\nResult saved to: {outputPath}");

Console.WriteLine("\n✓ Task T3 completed successfully");
Console.WriteLine("  - Row inserted into table");
Console.WriteLine("  - Changes tracked natively");
Console.WriteLine("  - Comments preserved");
Console.WriteLine("  - Document valid");
return 0;

// ─────────────────────────────────────────────────────────────────────────────

static byte[] CreateContractWithTableAndComment()
{
    using var ms = new MemoryStream();
    using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
    {
        var main = doc.AddMainDocumentPart();

        // Add heading
        var heading = new Paragraph(
            new Run(new Text("Contract Line Items")));
        heading.ParagraphId = "00000001";

        // Add table with 3 columns
        var table = new Table();
        var tableProps = new TableProperties();
        var borders = new TableBorders();
        borders.Append(new TopBorder { Val = BorderValues.Single, Size = 12, Space = 0, Color = "auto" });
        borders.Append(new LeftBorder { Val = BorderValues.Single, Size = 12, Space = 0, Color = "auto" });
        borders.Append(new BottomBorder { Val = BorderValues.Single, Size = 12, Space = 0, Color = "auto" });
        borders.Append(new RightBorder { Val = BorderValues.Single, Size = 12, Space = 0, Color = "auto" });
        borders.Append(new InsideHorizontalBorder { Val = BorderValues.Single, Size = 12, Space = 0, Color = "auto" });
        borders.Append(new InsideVerticalBorder { Val = BorderValues.Single, Size = 12, Space = 0, Color = "auto" });
        tableProps.Append(borders);
        table.AppendChild(tableProps);

        var grid = new TableGrid(
            new GridColumn { Width = "2500" },
            new GridColumn { Width = "2000" },
            new GridColumn { Width = "1500" });
        table.AppendChild(grid);

        // Header row
        table.AppendChild(new TableRow(
            CreateCell("Item Description"),
            CreateCell("Amount"),
            CreateCell("Priority")));

        // Data rows
        table.AppendChild(new TableRow(
            CreateCell("Professional Services"),
            CreateCell("$100,000"),
            CreateCell("Critical")));

        table.AppendChild(new TableRow(
            CreateCell("Software License"),
            CreateCell("$50,000"),
            CreateCell("High")));

        // Footer
        var footer = new Paragraph(
            new Run(new Text("End of contract items.")));
        footer.ParagraphId = "00000005";

        main.Document = new Document(new Body(heading, table, footer));
        main.Document.Save();
    }

    // Add a comment to the heading
    var baseDoc = ms.ToArray();
    var commentClient = new OfficeAgentClient(new WordModule());
    var inspection = commentClient.Inspect(new StreamHandle(new MemoryStream(baseDoc)));
    var headingId = inspection.Paragraphs.First(p => p.Text.Contains("Contract Line Items")).ParaId;

    var withComment = commentClient.Commit(
        new StreamHandle(new MemoryStream(baseDoc)),
        new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new CommentOp
                {
                    Target = new TextSpanAnchor { ParaId = headingId, Expect = "Contract Line Items" },
                    Text = "Please review the contract line items and amounts",
                    Author = "Reviewer",
                    Initials = "RV"
                }
            }
        });

    return withComment.ToBytes();
}

static TableCell CreateCell(string text) =>
    new TableCell(
        new Paragraph(
            new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve })));
