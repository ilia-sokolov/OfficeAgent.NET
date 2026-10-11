using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.PowerPoint;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// The Word verbs the docx-platform-tests conformance suite found missing: <c>link</c>
/// (#42), <c>mergeTableCells</c> (#43), and attaching a paragraph to an existing list with
/// <c>format</c>'s <c>numberingId</c> (#45). Each is checked against its recorded decision,
/// refusals included, and every output must pass the Open XML validator.
/// </summary>
public class ConformanceVerbTests
{
    private static OfficeAgentClient Client() => new(new WordModule());

    // ── link (#42) ───────────────────────────────────────────────────────

    [Fact]
    public void A_link_wraps_the_matched_runs_and_points_at_the_url()
    {
        var document = Commit(Paragraphs(("See the ", false), ("ECMA", true), (" home page.", false)), d => new LinkOp { Target = Span(d, "ECMA home"), Url = "https://ecma-international.org/" });

        using var opened = Open(document);
        var hyperlink = Assert.Single(opened.MainDocumentPart!.Document.Body!.Descendants<Hyperlink>());
        Assert.Equal("ECMA home", hyperlink.InnerText);
        var relationship = Assert.Single(opened.MainDocumentPart.HyperlinkRelationships, r => r.Id == hyperlink.Id);
        Assert.Equal(new Uri("https://ecma-international.org/"), relationship.Uri);
        Assert.True(relationship.IsExternal);
        // The runs keep their formatting: "ECMA" stays bold inside the link.
        Assert.Contains(hyperlink.Descendants<Run>(), run => run.InnerText == "ECMA" && run.RunProperties?.Bold is not null);
        Assert.Equal("See the ECMA home page.", opened.MainDocumentPart.Document.Body.InnerText);
        AssertValid(document);
    }

    [Fact]
    public void A_mailto_link_is_accepted()
    {
        var document = Commit(Paragraphs(("Write to legal@contoso.example today.", false)), d => new LinkOp { Target = Span(d, "legal@contoso.example"), Url = "mailto:legal@contoso.example" });

        using var opened = Open(document);
        Assert.Single(opened.MainDocumentPart!.Document.Body!.Descendants<Hyperlink>());
    }

    [Theory]
    [InlineData("javascript:alert(1)")]
    [InlineData("file:///C:/secrets.txt")]
    [InlineData("docs/readme.html")]
    [InlineData("")]
    public void A_url_outside_http_https_and_mailto_is_refused(string url)
    {
        var error = Refused(Paragraphs(("See the ECMA home page.", false)), d => new LinkOp { Target = Span(d, "ECMA home"), Url = url });

        Assert.Equal(ValidationErrorCodes.InvalidOperation, error.Code);
    }

    [Fact]
    public void A_tracked_link_is_refused_with_the_way_forward()
    {
        var error = Refused(Paragraphs(("See the ECMA home page.", false)), d => new LinkOp { Target = Span(d, "ECMA home"), Url = "https://ecma-international.org/", Mode = ChangeMode.Tracked });

        Assert.Equal(ValidationErrorCodes.InvalidOperation, error.Code);
        Assert.Contains("Direct", error.Message);
    }

    [Fact]
    public void Text_already_inside_a_link_is_not_linked_again()
    {
        var linked = Commit(Paragraphs(("See the ECMA home page.", false)), d => new LinkOp { Target = Span(d, "ECMA home"), Url = "https://ecma-international.org/" });

        var error = Refused(linked, d => new LinkOp { Target = Span(d, "ECMA"), Url = "https://example.com/" });

        Assert.Equal(ValidationErrorCodes.InvalidOperation, error.Code);
        Assert.Contains("already inside a hyperlink", error.Message);
    }

    [Fact]
    public void A_link_reads_from_plan_json()
    {
        var plan = JsonSerializer.Deserialize<DocumentPlan>(
            """{ "operations": [ { "op": "link", "target": { "paraId": "p1", "expect": "ECMA" }, "url": "https://ecma-international.org/" } ] }""",
            new JsonSerializerOptions { PropertyNameCaseInsensitive = true })!;

        var link = Assert.IsType<LinkOp>(Assert.Single(plan.Operations));
        Assert.Equal("https://ecma-international.org/", link.Url);
    }

    // ── mergeTableCells (#43) ────────────────────────────────────────────

    [Fact]
    public void Merging_two_cells_spans_the_grid_and_keeps_both_texts()
    {
        var document = Commit(Table(new[] { "Merge left", "and middle", "Right stable" }),
            new MergeTableCellsOp { Target = Cell(0, 0), ColumnSpan = 2 });

        using var opened = Open(document);
        var table = opened.MainDocumentPart!.Document.Body!.Descendants<Table>().Single();
        Assert.Equal(3, table.GetFirstChild<TableGrid>()!.Elements<GridColumn>().Count());
        var cells = table.Elements<TableRow>().First().Elements<TableCell>().ToList();
        Assert.Equal(2, cells.Count);
        Assert.Equal(2, cells[0].TableCellProperties!.GridSpan!.Val!.Value);
        Assert.Equal(2400, int.Parse(cells[0].TableCellProperties!.TableCellWidth!.Width!.Value!));
        Assert.Equal(new[] { "Merge left", "and middle" }, cells[0].Elements<Paragraph>().Select(p => p.InnerText));
        Assert.Equal("Right stable", cells[1].InnerText);
        AssertValid(document);
    }

    [Fact]
    public void An_empty_absorbed_cell_adds_no_paragraph()
    {
        var document = Commit(Table(new[] { "Merge left", "", "Right stable" }),
            new MergeTableCellsOp { Target = Cell(0, 0), ColumnSpan = 2 });

        using var opened = Open(document);
        var first = opened.MainDocumentPart!.Document.Body!.Descendants<TableCell>().First();
        Assert.Single(first.Elements<Paragraph>());
    }

    [Fact]
    public void A_span_past_the_end_of_the_row_is_refused()
    {
        var error = Refused(Table(new[] { "A", "B", "C" }), new MergeTableCellsOp { Target = Cell(0, 1), ColumnSpan = 3 });

        Assert.Equal(ValidationErrorCodes.InvalidOperation, error.Code);
        Assert.Contains("fewer than columnSpan", error.Message);
    }

    [Fact]
    public void A_span_that_would_split_an_existing_merge_is_refused()
    {
        var merged = Commit(Table(new[] { "A", "B", "C" }), new MergeTableCellsOp { Target = Cell(0, 1), ColumnSpan = 2 });

        var error = Refused(merged, new MergeTableCellsOp { Target = Cell(0, 0), ColumnSpan = 2 });

        Assert.Contains("use columnSpan 3", error.Message);
    }

    [Fact]
    public void A_span_no_wider_than_the_cell_is_refused()
    {
        var error = Refused(Table(new[] { "A", "B" }), new MergeTableCellsOp { Target = Cell(0, 0), ColumnSpan = 1 });

        Assert.Contains("must be greater than", error.Message);
    }

    [Fact]
    public void A_tracked_merge_is_refused()
    {
        var error = Refused(Table(new[] { "A", "B" }),
            new MergeTableCellsOp { Target = Cell(0, 0), ColumnSpan = 2, Mode = ChangeMode.Tracked });

        Assert.Equal(ValidationErrorCodes.InvalidOperation, error.Code);
        Assert.Contains("Direct", error.Message);
    }

    [Fact]
    public void A_vertically_merged_cell_is_refused()
    {
        var error = Refused(Table(new[] { "A", "B" }, verticalMerge: true),
            new MergeTableCellsOp { Target = Cell(0, 0), ColumnSpan = 2 });

        Assert.Contains("vertical merge", error.Message);
    }

    // ── numberingId (#45) ────────────────────────────────────────────────

    [Fact]
    public void Inspection_lists_the_documents_lists_by_what_they_look_like()
    {
        var node = Assert.Single(Client().Inspect(Numbered()).Nodes, n => n.Kind == "list");

        Assert.Equal("list#42", node.Path);
        Assert.Contains("Contract clauses", node.Summary);
        Assert.Contains("level 0 decimal \"%1.\" style ClauseHeading", node.Summary);
        Assert.Contains("level 1 lowerLetter \"(%2)\"", node.Summary);
    }

    [Fact]
    public void Attaching_continues_the_existing_list_and_takes_the_level_style()
    {
        var document = Commit(Numbered(), d => new FormatOp
        {
            Target = Span(d, "Number this paragraph"), NumberingId = 42, ListLevel = 0, Mode = ChangeMode.Direct
        });

        using var opened = Open(document);
        var paragraph = opened.MainDocumentPart!.Document.Body!.Elements<Paragraph>().Last();
        var numbering = paragraph.ParagraphProperties!.NumberingProperties!;
        Assert.Equal(42, numbering.NumberingId!.Val!.Value);
        Assert.Equal(0, numbering.NumberingLevelReference!.Val!.Value);
        Assert.Equal("ClauseHeading", paragraph.ParagraphProperties.ParagraphStyleId!.Val!.Value);
        // No new definition: the paragraph joined the list the document had.
        Assert.Single(opened.MainDocumentPart.NumberingDefinitionsPart!.Numbering!.Elements<NumberingInstance>());
        AssertValid(document);
    }

    [Fact]
    public void A_level_without_a_style_keeps_the_paragraph_style()
    {
        var document = Commit(Numbered(), d => new FormatOp
        {
            Target = Span(d, "Number this paragraph"), NumberingId = 42, ListLevel = 1, Mode = ChangeMode.Direct
        });

        using var opened = Open(document);
        var paragraph = opened.MainDocumentPart!.Document.Body!.Elements<Paragraph>().Last();
        Assert.Equal(1, paragraph.ParagraphProperties!.NumberingProperties!.NumberingLevelReference!.Val!.Value);
        Assert.Equal("BodyText", paragraph.ParagraphProperties.ParagraphStyleId!.Val!.Value);
    }

    [Fact]
    public void A_tracked_attach_records_the_paragraph_format_change()
    {
        var document = Commit(Numbered(), d => new FormatOp
        {
            Target = Span(d, "Number this paragraph"), NumberingId = 42, Mode = ChangeMode.Tracked
        });

        using var opened = Open(document);
        var paragraph = opened.MainDocumentPart!.Document.Body!.Elements<Paragraph>().Last();
        Assert.NotNull(paragraph.ParagraphProperties!.ParagraphPropertiesChange);
        AssertValid(document);
    }

    [Fact]
    public void An_unknown_numbering_id_is_not_found()
    {
        var error = Refused(Numbered(), d => new FormatOp { Target = Span(d, "Number this paragraph"), NumberingId = 7 });

        Assert.Equal(ValidationErrorCodes.AnchorNotFound, error.Code);
    }

    [Fact]
    public void Numbering_id_and_list_style_are_mutually_exclusive()
    {
        var error = Refused(Numbered(), d => new FormatOp { Target = Span(d, "Number this paragraph"), NumberingId = 42, ListStyle = "decimal" });

        Assert.Equal(ValidationErrorCodes.InvalidOperation, error.Code);
        Assert.Contains("mutually exclusive", error.Message);
    }

    [Fact]
    public void Numbering_id_with_a_list_id_is_refused()
    {
        var error = Refused(Numbered(), d => new FormatOp { Target = Span(d, "Number this paragraph"), NumberingId = 42, ListId = 1 });

        Assert.Equal(ValidationErrorCodes.InvalidOperation, error.Code);
    }

    // ── discovery ────────────────────────────────────────────────────────

    [Fact]
    public void Discovery_offers_the_new_verbs_for_word_and_not_for_decks()
    {
        var capabilities = CapabilityDiscovery.Describe(new IFormatModule[] { new WordModule(), new PowerPointModule() });
        var word = capabilities.Formats.Single(f => f.Format == OfficeAgent.Abstractions.DocumentFormat.Word);
        var deck = capabilities.Formats.Single(f => f.Format == OfficeAgent.Abstractions.DocumentFormat.PowerPoint);

        Assert.Contains("link", word.Operations);
        Assert.Contains("mergeTableCells", word.Operations);
        Assert.Contains("list", word.NodeKinds);
        Assert.DoesNotContain("link", deck.Operations);
        Assert.DoesNotContain("mergeTableCells", deck.Operations);
    }

    // ── fixtures ─────────────────────────────────────────────────────────

    private const string ParaId = "1A2B3C4D";

    /// <summary>A text anchor whose paragraph id comes from inspecting the document, as an agent gets it.</summary>
    private static TextSpanAnchor Span(byte[] document, string expect) => new()
    {
        ParaId = Client().Inspect(document).Paragraphs.First(p => p.Text.Contains(expect, StringComparison.Ordinal)).ParaId,
        Expect = expect
    };

    private static NodeAnchor Cell(int row, int column) => new() { Kind = "tableCell", Path = $"table#0/cell#{row}/{column}" };

    /// <summary>One paragraph built from runs, each optionally bold.</summary>
    private static byte[] Paragraphs(params (string Text, bool Bold)[] runs) => Build(body =>
    {
        var paragraph = new Paragraph { ParagraphId = ParaId };
        foreach (var (text, bold) in runs)
        {
            var run = new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
            if (bold) run.PrependChild(new RunProperties(new Bold()));
            paragraph.AppendChild(run);
        }
        body.AppendChild(paragraph);
    });

    /// <summary>A one-row table, every cell 1200 twips wide, on a three-column grid when it has three cells.</summary>
    private static byte[] Table(string[] cells, bool verticalMerge = false) => Build(body =>
    {
        var grid = new TableGrid(cells.Select(_ => new GridColumn { Width = "1200" }));
        var row = new TableRow(cells.Select((text, i) =>
        {
            var properties = new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "1200" });
            if (verticalMerge && i == 0) properties.AppendChild(new VerticalMerge { Val = MergedCellValues.Restart });
            return new TableCell(properties, new Paragraph(text.Length == 0 ? new Run() : new Run(new Text(text))));
        }));
        body.AppendChild(new Table(new TableProperties(new TableWidth { Type = TableWidthUnitValues.Auto, Width = "0" }), grid, row));
        body.AppendChild(new Paragraph());
    });

    /// <summary>
    /// A document with one list, numbering instance 42 over abstract definition 7: level 0
    /// decimal "%1." in style ClauseHeading, level 1 lowerLetter "(%2)". The paragraph to
    /// number is in style BodyText.
    /// </summary>
    private static byte[] Numbered() => Build(body =>
    {
        body.AppendChild(new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "BodyText" }),
            new Run(new Text("Number this paragraph"))) { ParagraphId = ParaId });
    }, main =>
    {
        var part = main.AddNewPart<NumberingDefinitionsPart>();
        part.Numbering = new Numbering(
            new AbstractNum(
                new MultiLevelType { Val = MultiLevelValues.HybridMultilevel },
                new AbstractNumDefinitionName { Val = "Contract clauses" },
                new Level(new StartNumberingValue { Val = 1 }, new NumberingFormat { Val = NumberFormatValues.Decimal },
                    new ParagraphStyleIdInLevel { Val = "ClauseHeading" }, new LevelText { Val = "%1." }) { LevelIndex = 0 },
                new Level(new StartNumberingValue { Val = 1 }, new NumberingFormat { Val = NumberFormatValues.LowerLetter },
                    new LevelText { Val = "(%2)" }) { LevelIndex = 1 })
            { AbstractNumberId = 7 },
            new NumberingInstance(new AbstractNumId { Val = 7 }) { NumberID = 42 });
    });

    private static byte[] Build(Action<Body> fill, Action<MainDocumentPart>? parts = null)
    {
        var buffer = new MemoryStream();
        using (var package = WordprocessingDocument.Create(buffer, WordprocessingDocumentType.Document))
        {
            var main = package.AddMainDocumentPart();
            var body = new Body();
            fill(body);
            main.Document = new Document(body);
            parts?.Invoke(main);
        }
        return buffer.ToArray();
    }

    private static byte[] Commit(byte[] document, PlanOperation operation)
    {
        using var applied = Client().Commit(new StreamHandle(new MemoryStream(document)), new DocumentPlan { Operations = new[] { operation } });
        Assert.True(applied.Committed, string.Join("; ", applied.Report.Errors.Select(e => $"{e.Code}: {e.Message}")));
        return applied.ToBytes();
    }

    private static byte[] Commit(byte[] document, Func<byte[], PlanOperation> operation) => Commit(document, operation(document));

    private static ValidationError Refused(byte[] document, Func<byte[], PlanOperation> operation) => Refused(document, operation(document));

    private static ValidationError Refused(byte[] document, PlanOperation operation)
    {
        var report = Client().Preview(new StreamHandle(new MemoryStream(document)), new DocumentPlan { Operations = new[] { operation } });
        Assert.False(report.IsValid);
        return Assert.Single(report.Errors);
    }

    private static WordprocessingDocument Open(byte[] bytes) => WordprocessingDocument.Open(new MemoryStream(bytes), isEditable: false);

    private static void AssertValid(byte[] document)
    {
        using var opened = Open(document);
        var problems = new OpenXmlValidator(FileFormatVersions.Office2019)
            .Validate(opened)
            .Select(e => $"{e.Path?.XPath}: {e.Description}")
            .ToList();
        Assert.True(problems.Count == 0, string.Join("; ", problems.Take(3)));
    }
}
