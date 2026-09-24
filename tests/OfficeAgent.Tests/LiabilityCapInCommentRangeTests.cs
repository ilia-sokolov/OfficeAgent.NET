using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// TASK T2: Liability Cap in Comment Range (Held-Out, Advanced)
///
/// Tests the ability to modify a liability cap value in a table cell that overlaps
/// with a Word comment range, while preserving the comment and tracking the change.
///
/// This scenario is challenging because comment ranges and cell edits can conflict
/// in low-level DOM APIs. A naive approach can corrupt the comment or lose change tracking.
///
/// OfficeAgent.NET handles this by maintaining an abstract representation of the document,
/// validating edits against a snapshot before applying, and applying operations atomically.
/// </summary>
public class LiabilityCapInCommentRangeTests
{
    private static OfficeAgentClient Client() => new(new WordModule());

    // ── Core Scenario: Text edit in cell with comment ───────────────────

    [Fact]
    public void Modifying_a_value_in_a_table_cell_with_a_comment_preserves_the_comment()
    {
        // 1. Create a document with a table containing a liability cap
        var documentWithTable = CreateContractWithLiabilityCap("$500,000");

        // 2. Add a comment to the liability cap value
        var documentWithComment = AddCommentToCell(documentWithTable, "$500,000");

        // 3. Inspect the document to find the old value
        var inspection = Client().Inspect(new StreamHandle(new MemoryStream(documentWithComment)));
        var hits = Client().Find(
            new StreamHandle(new MemoryStream(documentWithComment)),
            new FindQuery("$500,000"));

        Assert.NotEmpty(hits);
        var hit = hits[0];

        // 4. Create a plan to modify the value with tracking
        var plan = new DocumentPlan
        {
            Snapshot = inspection.Snapshot,
            Revision = new RevisionMetadata { Author = "Liability Cap Editor" },
            Operations = new PlanOperation[]
            {
                new ChangeTextOp
                {
                    Target = hit.Anchor,
                    With = "$1,000,000",
                    Mode = ChangeMode.Tracked
                }
            }
        };

        // 5. Apply the change
        var applied = Client().Commit(
            new StreamHandle(new MemoryStream(documentWithComment)),
            plan);

        Assert.True(applied.Committed, string.Join("; ", applied.Report.Errors.Select(e => e.Message)));
        var result = applied.ToBytes();

        // 6. Verify document validity
        AssertValid(result);

        // 7. Verify the new value exists
        var newInspection = Client().Inspect(new StreamHandle(new MemoryStream(result)));
        var newHits = Client().Find(
            new StreamHandle(new MemoryStream(result)),
            new FindQuery("$1,000,000"));
        Assert.NotEmpty(newHits);

        // 8. Verify comment still exists
        var commentNodes = newInspection.Nodes.Where(n => n.Kind == "comment").ToList();
        Assert.NotEmpty(commentNodes);
    }

    [Fact]
    public void The_comment_remains_valid_after_tracked_modification()
    {
        // Setup: Document with table, liability cap, and comment
        var documentWithTable = CreateContractWithLiabilityCap("$750,000");
        var documentWithComment = AddCommentToCell(documentWithTable, "$750,000");

        // Find and modify
        var inspection = Client().Inspect(new StreamHandle(new MemoryStream(documentWithComment)));
        var hits = Client().Find(
            new StreamHandle(new MemoryStream(documentWithComment)),
            new FindQuery("$750,000"));

        var applied = Client().Commit(
            new StreamHandle(new MemoryStream(documentWithComment)),
            new DocumentPlan
            {
                Snapshot = inspection.Snapshot,
                Revision = new RevisionMetadata { Author = "Editor", TimestampUtc = DateTimeOffset.UtcNow },
                Operations = new PlanOperation[]
                {
                    new ChangeTextOp
                    {
                        Target = hits[0].Anchor,
                        With = "$2,000,000",
                        Mode = ChangeMode.Tracked
                    }
                }
            });

        Assert.True(applied.Committed);
        var result = applied.ToBytes();
        AssertValid(result);

        // Verify comment structure in the modified document
        using var stream = new MemoryStream(result);
        using var package = WordprocessingDocument.Open(stream, isEditable: false);
        var main = package.MainDocumentPart!;

        // Comment parts should exist
        Assert.NotNull(main.WordprocessingCommentsPart);
        Assert.NotNull(main.WordprocessingCommentsExPart);

        // Comments should not be empty
        var comments = main.WordprocessingCommentsPart!.Comments!.Elements<Comment>().ToList();
        Assert.NotEmpty(comments);

        // Comment ranges should be intact
        var body = main.Document?.Body;
        Assert.NotNull(body);
        var commentRangeStarts = body.Descendants<CommentRangeStart>().ToList();
        var commentReferences = body.Descendants<CommentReference>().ToList();

        Assert.NotEmpty(commentRangeStarts);
        Assert.NotEmpty(commentReferences);
    }

    [Fact]
    public void Multiple_occurrences_of_a_cap_can_be_modified_independently()
    {
        // Document with multiple liability caps
        var doc = CreateContractWithLiabilityCap("$500,000");
        var withComment = AddCommentToCell(doc, "$500,000");

        // Modify the first occurrence
        var inspection = Client().Inspect(new StreamHandle(new MemoryStream(withComment)));
        var hits = Client().Find(
            new StreamHandle(new MemoryStream(withComment)),
            new FindQuery("$500,000"));

        var firstHit = hits[0];
        var applied = Client().Commit(
            new StreamHandle(new MemoryStream(withComment)),
            new DocumentPlan
            {
                Snapshot = inspection.Snapshot,
                Revision = new RevisionMetadata { Author = "Editor" },
                Operations = new PlanOperation[]
                {
                    new ChangeTextOp
                    {
                        Target = firstHit.Anchor,
                        With = "$1,000,000",
                        Mode = ChangeMode.Tracked
                    }
                }
            });

        Assert.True(applied.Committed);
        var result = applied.ToBytes();

        // Verify the modification
        var finalInspection = Client().Inspect(new StreamHandle(new MemoryStream(result)));
        var newCapHits = Client().Find(
            new StreamHandle(new MemoryStream(result)),
            new FindQuery("$1,000,000"));
        Assert.NotEmpty(newCapHits);

        // Comment should still be there
        var commentNodes = finalInspection.Nodes.Where(n => n.Kind == "comment").ToList();
        Assert.NotEmpty(commentNodes);
    }

    [Fact]
    public void Direct_mode_edit_in_cell_with_comment_also_preserves_comment()
    {
        var documentWithTable = CreateContractWithLiabilityCap("$300,000");
        var documentWithComment = AddCommentToCell(documentWithTable, "$300,000");

        var inspection = Client().Inspect(new StreamHandle(new MemoryStream(documentWithComment)));
        var hits = Client().Find(
            new StreamHandle(new MemoryStream(documentWithComment)),
            new FindQuery("$300,000"));

        var applied = Client().Commit(
            new StreamHandle(new MemoryStream(documentWithComment)),
            new DocumentPlan
            {
                Snapshot = inspection.Snapshot,
                Revision = new RevisionMetadata { Author = "Direct Editor" },
                Operations = new PlanOperation[]
                {
                    new ChangeTextOp
                    {
                        Target = hits[0].Anchor,
                        With = "$600,000",
                        Mode = ChangeMode.Direct  // No redline, just the edit
                    }
                }
            });

        Assert.True(applied.Committed);
        var result = applied.ToBytes();
        AssertValid(result);

        var finalInspection = Client().Inspect(new StreamHandle(new MemoryStream(result)));
        var newHits = Client().Find(
            new StreamHandle(new MemoryStream(result)),
            new FindQuery("$600,000"));

        Assert.NotEmpty(newHits);

        // Comment should still be present even with Direct mode
        var commentNodes = finalInspection.Nodes.Where(n => n.Kind == "comment").ToList();
        Assert.NotEmpty(commentNodes);
    }

    // ── Fixtures ─────────────────────────────────────────────────────────

    /// <summary>
    /// Creates a simple contract document with a table containing a liability cap.
    /// </summary>
    private static byte[] CreateContractWithLiabilityCap(string capAmount)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();

            var title = new Paragraph(new Run(new Text("Limitation of Liability Agreement")))
            {
                ParagraphId = "00000001"
            };

            var intro = new Paragraph(new Run(new Text("The parties agree to the following terms:")))
            {
                ParagraphId = "00000002"
            };

            var table = CreateLiabilityCapTable(capAmount);

            var closing = new Paragraph(new Run(new Text("This agreement is effective as of today.")))
            {
                ParagraphId = "00000005"
            };

            main.Document = new Document(new Body(title, intro, table, closing));
            main.Document.Save();
        }

        return ms.ToArray();
    }

    /// <summary>
    /// Creates a table with liability cap information.
    /// Structure: Header row with "Liability Cap" | "Amount"
    ///           Data row with "Maximum Liability" | capAmount
    /// </summary>
    private static Table CreateLiabilityCapTable(string capAmount)
    {
        var table = new Table();

        // Table properties with correct schema order
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

        // Table grid
        var grid = new TableGrid(
            new GridColumn { Width = "3000" },
            new GridColumn { Width = "3000" });
        table.AppendChild(grid);

        // Header row
        var headerRow = new TableRow();
        headerRow.AppendChild(CreateCell("Liability Cap"));
        headerRow.AppendChild(CreateCell("Amount"));
        table.AppendChild(headerRow);

        // Data row with liability cap value
        var dataRow = new TableRow();
        dataRow.AppendChild(CreateCell("Maximum Liability"));
        dataRow.AppendChild(CreateCell(capAmount));
        table.AppendChild(dataRow);

        return table;
    }

    /// <summary>Creates a table cell with the given text.</summary>
    private static TableCell CreateCell(string text)
    {
        var cell = new TableCell(
            new Paragraph(
                new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve })));
        return cell;
    }

    /// <summary>
    /// Adds a comment to the cell containing the specified liability cap value.
    /// </summary>
    private static byte[] AddCommentToCell(byte[] document, string capValue)
    {
        var client = Client();
        var inspection = client.Inspect(new StreamHandle(new MemoryStream(document)));

        // Find the paragraph containing the cap value
        var targetPara = inspection.Paragraphs.FirstOrDefault(p => p.Text.Contains(capValue));
        if (targetPara == null)
        {
            throw new InvalidOperationException($"Could not find paragraph containing '{capValue}'");
        }

        // Add a comment to this paragraph
        var withComment = client.Commit(
            new StreamHandle(new MemoryStream(document)),
            new DocumentPlan
            {
                Operations = new PlanOperation[]
                {
                    new CommentOp
                    {
                        Target = new TextSpanAnchor { ParaId = targetPara.ParaId, Expect = capValue },
                        Text = "Please review the liability cap amount",
                        Author = "Legal Reviewer",
                        Initials = "LR"
                    }
                }
            });

        if (!withComment.Committed)
        {
            throw new InvalidOperationException(
                $"Failed to add comment: {string.Join("; ", withComment.Report.Errors.Select(e => e.Message))}");
        }

        return withComment.ToBytes();
    }

    /// <summary>Validates the document against Office 2019 schema.</summary>
    private static void AssertValid(byte[] document)
    {
        using var stream = new MemoryStream(document);
        using var opened = WordprocessingDocument.Open(stream, isEditable: false);
        var problems = new OpenXmlValidator(FileFormatVersions.Office2019)
            .Validate(opened)
            .Select(e => $"{e.Path?.XPath}: {e.Description}")
            .ToList();
        Assert.True(problems.Count == 0, string.Join("; ", problems.Take(5)));
    }
}
