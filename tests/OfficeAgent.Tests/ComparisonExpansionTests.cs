using System.Diagnostics;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// The two differences this comparison learned to understand: text split across runs that
/// carry the same formatting, and cell words edited inside a table whose shape did not
/// move. Everything it still cannot represent stays refused.
/// </summary>
public sealed class ComparisonExpansionTests
{
    // ── Equivalent run segmentation is not a difference ──────────────────

    /// <summary>
    /// The same sentence stored as one run and as three identically formatted runs is the
    /// same sentence. Word re-segments constantly, and reporting that as a formatting
    /// change withheld the plan for documents nobody would call different.
    /// </summary>
    [Fact]
    public void Equivalent_run_segmentation_produces_no_difference()
    {
        var original = Document(Paragraph(Run("Payment is due in thirty days.")));
        var revised = Document(Paragraph(
            Run("Payment is due "), Run("in thirty "), Run("days.")));

        var comparison = Client().CompareDocuments(original, revised);

        Assert.Empty(comparison.Differences);
        Assert.True(comparison.IsComplete, Explain(comparison));
        Assert.NotNull(comparison.Plan);
        Assert.Empty(comparison.Plan!.Operations);
    }

    /// <summary>
    /// Re-segmentation around a genuine edit still yields exactly one text difference,
    /// not a formatting refusal.
    /// </summary>
    [Fact]
    public void A_text_edit_across_resegmented_runs_is_one_ordinary_difference()
    {
        var original = Document(Paragraph(Run("Payment is due in thirty days.")));
        var revised = Document(Paragraph(
            Run("Payment is due "), Run("in forty-five "), Run("days.")));

        var comparison = Client().CompareDocuments(original, revised);

        var difference = Assert.Single(comparison.Differences);
        Assert.Equal("Payment is due in thirty days.", difference.Before);
        Assert.Equal("Payment is due in forty-five days.", difference.After);
        Assert.True(comparison.IsComplete, Explain(comparison));
    }

    /// <summary>
    /// Merging only applies to runs that agree on formatting. A word that became bold is
    /// a formatting change and is still refused, because this comparison cannot express
    /// one.
    /// </summary>
    [Fact]
    public void Differently_formatted_runs_are_still_a_refusal()
    {
        var original = Document(Paragraph(Run("Payment is due in thirty days.")));
        var revised = Document(Paragraph(
            Run("Payment is due in "), BoldRun("thirty"), Run(" days.")));

        var comparison = Client().CompareDocuments(original, revised);

        Assert.False(comparison.IsComplete);
        Assert.Null(comparison.Plan);
        Assert.Contains(comparison.Diagnostics, d => d.Code == "unsupported-paragraph-markup-change");
    }

    /// <summary>
    /// Non-text content keeps its boundary. A run holding a line break is not merged into
    /// its neighbour, so inserting one is still seen.
    /// </summary>
    [Fact]
    public void A_run_carrying_a_break_is_not_merged_away()
    {
        var original = Document(Paragraph(Run("First part second part")));
        var revised = Document(Paragraph(
            Run("First part"), BreakRun(), Run("second part")));

        var comparison = Client().CompareDocuments(original, revised);

        Assert.False(comparison.IsComplete);
        Assert.Null(comparison.Plan);
    }

    /// <summary>
    /// Whitespace is text, not noise. A paragraph whose spacing changed is a real
    /// difference and is reported as one.
    /// </summary>
    [Fact]
    public void Whitespace_differences_are_preserved_as_text_differences()
    {
        var original = Document(Paragraph(Run("Alpha beta")));
        var revised = Document(Paragraph(Run("Alpha  beta")));

        var comparison = Client().CompareDocuments(original, revised);

        var difference = Assert.Single(comparison.Differences);
        Assert.Equal("Alpha beta", difference.Before);
        Assert.Equal("Alpha  beta", difference.After);
    }

    // ── Cells with unchanged geometry ────────────────────────────────────

    /// <summary>
    /// A cell whose words changed inside a table whose shape did not produces a tracked
    /// revision that accepts to the new text and rejects back to the old one.
    /// </summary>
    [Fact]
    public void A_cell_text_change_accepts_and_rejects_to_the_right_semantics()
    {
        var original = TableDocument("Region", "Q1", "North", "41850");
        var revised = TableDocument("Region", "Q1", "North", "58200");
        var client = Client();

        var comparison = client.CompareDocuments(original, revised);

        Assert.True(comparison.IsComplete, Explain(comparison));
        Assert.NotNull(comparison.Plan);
        Assert.Contains(comparison.Differences, d => d.Before == "41850" && d.After == "58200");

        using var applied = client.Commit(Handle(original), comparison.Plan!);
        Assert.True(applied.Committed, string.Join("; ",
            applied.Report.Errors.Select(error => $"{error.Code}: {error.Message}")));
        var redlined = applied.ToBytes();

        // The redline is a real Word revision: accepting gives the revised document's
        // text, rejecting gives the original's.
        Assert.Contains("58200", VisibleText(Resolve(client, redlined, RevisionAction.Accept)),
            StringComparison.Ordinal);
        Assert.Contains("41850", VisibleText(Resolve(client, redlined, RevisionAction.Reject)),
            StringComparison.Ordinal);
        Assert.DoesNotContain("41850", VisibleText(Resolve(client, redlined, RevisionAction.Accept)),
            StringComparison.Ordinal);

        AssertSchemaValid(redlined);
    }

    /// <summary>Several cells may change at once, each reported separately.</summary>
    [Fact]
    public void Several_cell_changes_are_reported_individually()
    {
        var original = TableDocument("Region", "Q1", "North", "41850");
        var revised = TableDocument("Region", "Q2", "South", "58200");

        var comparison = Client().CompareDocuments(original, revised);

        Assert.True(comparison.IsComplete, Explain(comparison));
        Assert.Equal(3, comparison.Differences.Count);
    }

    /// <summary>
    /// A table whose geometry moved is still refused. Positional alignment is only sound
    /// while the shape is identical, so a changed grid blocks the plan.
    /// </summary>
    [Fact]
    public void A_changed_table_geometry_is_still_refused()
    {
        var original = TableDocument("Region", "Q1", "North", "41850");
        var revised = Edit(original, document =>
        {
            var table = document.MainDocumentPart!.Document.Body!.Elements<Table>().Single();
            table.Append(new TableRow(
                new TableCell(new Paragraph(new Run(new Text("South")))),
                new TableCell(new Paragraph(new Run(new Text("12300"))))));
        });

        var comparison = Client().CompareDocuments(original, revised);

        Assert.False(comparison.IsComplete);
        Assert.Null(comparison.Plan);
        Assert.Contains(comparison.Coverage, area =>
            area.Name == "tables" && area.State == ComparisonAreaState.Blocked);
    }

    /// <summary>
    /// A cell whose formatting changed alongside its text is refused rather than flattened
    /// to plain text, because reproducing it would lose the formatting being edited.
    /// </summary>
    [Fact]
    public void A_cell_whose_formatting_changed_is_refused_not_flattened()
    {
        var original = TableDocument("Region", "Q1", "North", "41850");
        var revised = Edit(original, document =>
        {
            var cell = document.MainDocumentPart!.Document.Body!
                .Descendants<TableCell>().Last();
            cell.RemoveAllChildren<Paragraph>();
            cell.Append(new Paragraph(new Run(
                new RunProperties(new Bold()), new Text("58200"))));
        });

        var comparison = Client().CompareDocuments(original, revised);

        Assert.False(comparison.IsComplete);
        Assert.Null(comparison.Plan);
        Assert.Contains(comparison.Diagnostics, d =>
            d.Code is "unsupported-table-markup-change" or "unsupported-table-change");
    }

    /// <summary>Body and cell changes in one comparison are both reported.</summary>
    [Fact]
    public void A_body_change_and_a_cell_change_are_reported_together()
    {
        var original = TableDocument("Region", "Q1", "North", "41850", lead: "Regional results.");
        var revised = TableDocument("Region", "Q1", "North", "58200", lead: "Regional summary.");

        var comparison = Client().CompareDocuments(original, revised);

        Assert.True(comparison.IsComplete, Explain(comparison));
        Assert.Contains(comparison.Differences, d => d.Before == "Regional results.");
        Assert.Contains(comparison.Differences, d => d.Before == "41850");
    }

    // ── Bounded ──────────────────────────────────────────────────────────

    /// <summary>
    /// The difference ceiling applies to cell changes as well, so a large table cannot
    /// produce an unbounded result, and a truncated run never claims completeness.
    /// </summary>
    [Fact]
    public void Cell_changes_respect_the_difference_ceiling()
    {
        var originalCells = Enumerable.Range(0, 40).Select(i => $"Old {i}").ToArray();
        var revisedCells = Enumerable.Range(0, 40).Select(i => $"New {i}").ToArray();

        var comparison = Client().CompareDocuments(
            TableDocument(originalCells),
            TableDocument(revisedCells),
            new DocumentComparisonOptions { MaximumDifferences = 5 });

        Assert.False(comparison.IsComplete);
        Assert.Null(comparison.Plan);
        Assert.True(comparison.Differences.Count <= 5,
            $"expected at most 5 differences, got {comparison.Differences.Count}");
        Assert.Contains(comparison.Diagnostics, d => d.Code == "comparison-difference-limit-exceeded");
    }

    /// <summary>
    /// A long paragraph and a wide table finish quickly. The comparison walks structure
    /// rather than doing anything quadratic in the text.
    /// </summary>
    [Fact]
    public void Long_content_compares_within_a_sane_budget()
    {
        var longText = string.Join(" ", Enumerable.Range(0, 4000).Select(i => $"word{i}"));
        var original = Document(Paragraph(Run(longText)), Paragraph(Run("Tail")));
        var revised = Document(Paragraph(Run(longText)), Paragraph(Run("Tail changed")));

        var stopwatch = Stopwatch.StartNew();
        var comparison = Client().CompareDocuments(original, revised);
        stopwatch.Stop();

        Assert.True(comparison.IsComplete, Explain(comparison));
        Assert.Single(comparison.Differences);
        Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(20),
            $"comparison took {stopwatch.Elapsed}");
    }

    // ── Helpers ──────────────────────────────────────────────────────────

    private static OfficeAgentClient Client() => new(new WordModule());

    private static string Explain(DocumentComparisonResult comparison) => string.Join("; ",
        comparison.Diagnostics.Select(d => $"{d.Code}: {d.Message}"));

    private static StreamHandle Handle(byte[] bytes) => new(new MemoryStream(bytes, writable: false));

    private static byte[] Resolve(OfficeAgentClient client, byte[] bytes, RevisionAction action)
    {
        using var result = client.Commit(Handle(bytes), new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new RevisionOp { Target = new NodeAnchor { Kind = "revision", Path = "all" }, Action = action }
            }
        });
        Assert.True(result.Committed, string.Join("; ",
            result.Report.Errors.Select(error => $"{error.Code}: {error.Message}")));
        return result.ToBytes();
    }

    private static string VisibleText(byte[] bytes)
    {
        using var document = WordprocessingDocument.Open(new MemoryStream(bytes, writable: false), false);
        return string.Concat(document.MainDocumentPart!.Document!.Body!
            .Descendants<Text>().Select(text => text.Text));
    }

    private static void AssertSchemaValid(byte[] bytes)
    {
        using var document = WordprocessingDocument.Open(new MemoryStream(bytes, writable: false), false);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2019).Validate(document)
            .Select(error => error.Description));
    }

    private static Run Run(string text) =>
        new(new Text(text) { Space = SpaceProcessingModeValues.Preserve });

    private static Run BoldRun(string text) =>
        new(new RunProperties(new Bold()), new Text(text) { Space = SpaceProcessingModeValues.Preserve });

    private static Run BreakRun() => new(new Break());

    private static Paragraph Paragraph(params Run[] runs) => new(runs);

    private static byte[] Document(params Paragraph[] paragraphs)
    {
        using var ms = new MemoryStream();
        using (var document = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(paragraphs.Cast<OpenXmlElement>()));
            main.Document.Save();
        }

        return ms.ToArray();
    }

    /// <summary>A document with one lead paragraph and a two-column table of the cells.</summary>
    private static byte[] TableDocument(params string[] cells) => TableDocument(cells, "Regional results.");

    private static byte[] TableDocument(string[] cells, string lead)
    {
        using var ms = new MemoryStream();
        using (var document = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            var rows = cells.Chunk(2).Select(pair => new TableRow(
                pair.Select(value => new TableCell(new Paragraph(new Run(new Text(value)))))
                    .Cast<OpenXmlElement>()));

            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text(lead))),
                new Table(new[]
                {
                    (OpenXmlElement)new TableProperties(new TableStyle { Val = "TableGrid" }),
                    new TableGrid(new GridColumn { Width = "4000" }, new GridColumn { Width = "4000" })
                }.Concat(rows)),
                new Paragraph()));
            main.Document.Save();
        }

        return ms.ToArray();
    }

    private static byte[] TableDocument(string headerLeft, string headerRight, string cellLeft, string cellRight,
        string lead = "Regional results.") =>
        TableDocument(new[] { headerLeft, headerRight, cellLeft, cellRight }, lead);

    private static byte[] Edit(byte[] source, Action<WordprocessingDocument> mutate)
    {
        using var ms = new MemoryStream();
        ms.Write(source, 0, source.Length);
        ms.Position = 0;
        using (var document = WordprocessingDocument.Open(ms, isEditable: true))
        {
            mutate(document);
            document.MainDocumentPart!.Document.Save();
        }

        return ms.ToArray();
    }
}
