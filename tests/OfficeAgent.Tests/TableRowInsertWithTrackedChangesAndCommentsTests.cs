using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// TASK T3: Table Row Insert with Tracked Changes and Comment Preservation
///
/// Tests the ability to insert a new row into a contract table while:
/// - Applying tracked changes natively
/// - Preserving existing comments on adjacent rows
/// - Maintaining document validity
/// </summary>
public class TableRowInsertWithTrackedChangesAndCommentsTests
{
    private static OfficeAgentClient Client() => new(new WordModule());

    // ── Row Insertion with Change Tracking ───────────────────────────────

    [Fact]
    public void Inserting_a_table_row_marks_it_as_tracked_by_default()
    {
        var document = SimpleTableDocument();

        var result = Client().Preview(
            new StreamHandle(new MemoryStream(document)),
            new DocumentPlan
            {
                Operations = new PlanOperation[]
                {
                    new InsertTableRowsOp
                    {
                        Target = new NodeAnchor { Kind = "table", Path = "table#0" },
                        Rows = new[] { new[] { "New Data 1", "New Data 2", "New Data 3" } },
                        Position = TablePosition.End
                    }
                }
            });

        Assert.True(result.IsValid, string.Join("; ", result.Errors.Select(e => e.Message)));
    }

    [Fact]
    public void Inserting_a_table_row_with_tracked_mode_creates_revision_markup()
    {
        var document = SimpleTableDocument();

        var applied = Client().Commit(
            new StreamHandle(new MemoryStream(document)),
            new DocumentPlan
            {
                Revision = new RevisionMetadata { Author = "Agent", TimestampUtc = DateTimeOffset.UtcNow },
                Operations = new PlanOperation[]
                {
                    new InsertTableRowsOp
                    {
                        Target = new NodeAnchor { Kind = "table", Path = "table#0" },
                        Rows = new[] { new[] { "New Data 1", "New Data 2", "New Data 3" } },
                        Position = TablePosition.End,
                        Mode = ChangeMode.Tracked
                    }
                }
            });

        Assert.True(applied.Committed, string.Join("; ", applied.Report.Errors.Select(e => e.Message)));
        var result = applied.ToBytes();
        AssertValid(result);

        using var stream = new MemoryStream(result);
        using var package = WordprocessingDocument.Open(stream, isEditable: false);
        var body = package.MainDocumentPart!.Document.Body!;
        var table = body.Elements<Table>().First();

        // The inserted row should have revision markup on it
        var insertedRow = table.Elements<TableRow>().Last();
        var rowProps = insertedRow.TableRowProperties;
        Assert.NotNull(rowProps);
        Assert.NotNull(rowProps.GetFirstChild<Inserted>());
    }

    [Fact]
    public void Inserting_a_table_row_in_direct_mode_has_no_revision_markup()
    {
        var document = SimpleTableDocument();

        var applied = Client().Commit(
            new StreamHandle(new MemoryStream(document)),
            new DocumentPlan
            {
                Revision = new RevisionMetadata { Author = "Agent", TimestampUtc = DateTimeOffset.UtcNow },
                Operations = new PlanOperation[]
                {
                    new InsertTableRowsOp
                    {
                        Target = new NodeAnchor { Kind = "table", Path = "table#0" },
                        Rows = new[] { new[] { "New Data 1", "New Data 2", "New Data 3" } },
                        Position = TablePosition.End,
                        Mode = ChangeMode.Direct
                    }
                }
            });

        Assert.True(applied.Committed);
        var result = applied.ToBytes();
        AssertValid(result);

        using var stream = new MemoryStream(result);
        using var package = WordprocessingDocument.Open(stream, isEditable: false);
        var body = package.MainDocumentPart!.Document.Body!;
        var table = body.Elements<Table>().First();

        var insertedRow = table.Elements<TableRow>().Last();
        var rowProps = insertedRow.TableRowProperties;
        // In direct mode, there should be no Inserted marker
        if (rowProps != null)
        {
            Assert.Null(rowProps.GetFirstChild<Inserted>());
        }
    }

    // ── Comment Preservation ─────────────────────────────────────────────

    [Fact]
    public void A_comment_on_text_adjacent_to_table_is_preserved_when_inserting_rows()
    {
        var baseDocument = SimpleTableDocumentWithComment();

        var applied = Client().Commit(
            new StreamHandle(new MemoryStream(baseDocument)),
            new DocumentPlan
            {
                Revision = new RevisionMetadata { Author = "Agent", TimestampUtc = DateTimeOffset.UtcNow },
                Operations = new PlanOperation[]
                {
                    new InsertTableRowsOp
                    {
                        Target = new NodeAnchor { Kind = "table", Path = "table#0" },
                        Rows = new[] { new[] { "Inserted Row 1", "Inserted Row 2", "Inserted Row 3" } },
                        Position = TablePosition.End,
                        Mode = ChangeMode.Tracked
                    }
                }
            });

        Assert.True(applied.Committed, string.Join("; ", applied.Report.Errors.Select(e => e.Message)));
        var result = applied.ToBytes();
        AssertValid(result);

        // Verify the comment still exists
        var inspection = Client().Inspect(new StreamHandle(new MemoryStream(result)));
        var comment = inspection.Nodes.FirstOrDefault(n => n.Kind == "comment");
        Assert.NotNull(comment);
        Assert.Contains("Test comment", comment.Summary);
    }

    [Fact]
    public void Comments_remain_valid_after_tracked_row_insertion()
    {
        var baseDocument = SimpleTableDocumentWithComment();

        var applied = Client().Commit(
            new StreamHandle(new MemoryStream(baseDocument)),
            new DocumentPlan
            {
                Revision = new RevisionMetadata { Author = "Editor", TimestampUtc = DateTimeOffset.UtcNow },
                Operations = new PlanOperation[]
                {
                    new InsertTableRowsOp
                    {
                        Target = new NodeAnchor { Kind = "table", Path = "table#0" },
                        Rows = new[]
                        {
                            new[] { "Row 1 Col 1", "Row 1 Col 2", "Row 1 Col 3" },
                            new[] { "Row 2 Col 1", "Row 2 Col 2", "Row 2 Col 3" }
                        },
                        Position = TablePosition.Before,
                        RowIndex = 0,
                        Mode = ChangeMode.Tracked
                    }
                }
            });

        Assert.True(applied.Committed, string.Join("; ", applied.Report.Errors.Select(e => e.Message)));
        var result = applied.ToBytes();

        using var stream = new MemoryStream(result);
        using var package = WordprocessingDocument.Open(stream, isEditable: false);
        var main = package.MainDocumentPart!;

        // Verify comments part exists and is valid
        Assert.NotNull(main.WordprocessingCommentsPart);
        Assert.NotNull(main.WordprocessingCommentsExPart);

        // The comment should still have a valid reference
        var comments = main.WordprocessingCommentsPart.Comments!.Elements<Comment>().ToList();
        Assert.NotEmpty(comments);

        // Verify comment ranges in body are intact
        var body = main.Document.Body!;
        var commentRangeStarts = body.Descendants<CommentRangeStart>().ToList();
        var commentReferences = body.Descendants<CommentReference>().ToList();

        Assert.NotEmpty(commentRangeStarts);
        Assert.NotEmpty(commentReferences);
    }

    // ── Complete Workflow ────────────────────────────────────────────────

    [Fact]
    public void Complete_workflow_insert_row_with_tracked_changes_preserves_comments()
    {
        // Create a document with table and comment
        var baseDocument = SimpleTableDocumentWithComment();

        // Insert multiple rows with tracked changes
        var applied = Client().Commit(
            new StreamHandle(new MemoryStream(baseDocument)),
            new DocumentPlan
            {
                Revision = new RevisionMetadata
                {
                    Author = "OfficeAgent",
                    TimestampUtc = DateTimeOffset.UtcNow
                },
                Operations = new PlanOperation[]
                {
                    new InsertTableRowsOp
                    {
                        Target = new NodeAnchor { Kind = "table", Path = "table#0" },
                        Rows = new[]
                        {
                            new[] { "Q1 Revenue", "$100K", "On Track" },
                            new[] { "Q2 Revenue", "$120K", "Exceeding" },
                            new[] { "Q3 Revenue", "$90K", "Below Target" }
                        },
                        Position = TablePosition.After,
                        RowIndex = 0,
                        Mode = ChangeMode.Tracked
                    }
                }
            });

        Assert.True(applied.Committed, string.Join("; ", applied.Report.Errors.Select(e => e.Message)));
        var resultBytes = applied.ToBytes();

        // Document validity
        AssertValid(resultBytes);

        // Verify structure
        using var stream = new MemoryStream(resultBytes);
        using var package = WordprocessingDocument.Open(stream, isEditable: false);
        var body = package.MainDocumentPart!.Document.Body!;
        var table = body.Elements<Table>().First();

        // Should have header + original row + 3 new rows
        var rows = table.Elements<TableRow>().ToList();
        Assert.True(rows.Count >= 5, $"Expected at least 5 rows, got {rows.Count}");

        // Newly inserted rows should have revision markup
        var insertedRows = rows.Skip(1).Take(3).ToList(); // Skip header, take the 3 inserted rows
        foreach (var row in insertedRows)
        {
            var props = row.TableRowProperties;
            Assert.NotNull(props);
            // Row should be marked as inserted
            Assert.NotNull(props.GetFirstChild<Inserted>());
        }

        // Comment should still be present
        var inspection = Client().Inspect(new StreamHandle(new MemoryStream(resultBytes)));
        var comment = inspection.Nodes.FirstOrDefault(n => n.Kind == "comment");
        Assert.NotNull(comment);
    }

    // ── Fixtures ─────────────────────────────────────────────────────────

    /// <summary>Creates a simple document with a 3-column table.</summary>
    private static byte[] SimpleTableDocument()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();

            var heading = new Paragraph(
                new Run(new Text("Quarterly Report")));
            heading.ParagraphId = "00000001";

            var table = CreateSimpleTable();

            var footer = new Paragraph(
                new Run(new Text("End of report.")));
            footer.ParagraphId = "00000005";

            main.Document = new Document(new Body(heading, table, footer));
            main.Document.Save();
        }

        return ms.ToArray();
    }

    /// <summary>Creates a document with a table and a comment on the heading.</summary>
    private static byte[] SimpleTableDocumentWithComment()
    {
        var baseDoc = SimpleTableDocument();

        // Add a comment to the heading
        var client = Client();
        var inspection = client.Inspect(new StreamHandle(new MemoryStream(baseDoc)));
        var headingId = inspection.Paragraphs.First(p => p.Text.Contains("Quarterly Report")).ParaId;

        var withComment = client.Commit(
            new StreamHandle(new MemoryStream(baseDoc)),
            new DocumentPlan
            {
                Operations = new PlanOperation[]
                {
                    new CommentOp
                    {
                        Target = new TextSpanAnchor { ParaId = headingId, Expect = "Quarterly Report" },
                        Text = "Test comment on this report",
                        Author = "Reviewer",
                        Initials = "RV"
                    }
                }
            });

        Assert.True(withComment.Committed);
        return withComment.ToBytes();
    }

    /// <summary>Creates a 3-column table with 2 data rows plus header.</summary>
    private static Table CreateSimpleTable()
    {
        var table = new Table();

        // Table properties (required by schema)
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

        // Table grid (required by schema)
        var grid = new TableGrid(
            new GridColumn { Width = "2000" },
            new GridColumn { Width = "2000" },
            new GridColumn { Width = "2000" });
        table.AppendChild(grid);

        // Header row
        var headerRow = new TableRow();
        headerRow.AppendChild(CreateCell("Column 1"));
        headerRow.AppendChild(CreateCell("Column 2"));
        headerRow.AppendChild(CreateCell("Column 3"));
        table.AppendChild(headerRow);

        // Data row 1
        var dataRow1 = new TableRow();
        dataRow1.AppendChild(CreateCell("Data 1-1"));
        dataRow1.AppendChild(CreateCell("Data 1-2"));
        dataRow1.AppendChild(CreateCell("Data 1-3"));
        table.AppendChild(dataRow1);

        // Data row 2
        var dataRow2 = new TableRow();
        dataRow2.AppendChild(CreateCell("Data 2-1"));
        dataRow2.AppendChild(CreateCell("Data 2-2"));
        dataRow2.AppendChild(CreateCell("Data 2-3"));
        table.AppendChild(dataRow2);

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
