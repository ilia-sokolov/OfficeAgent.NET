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
/// Every verb that changes Word content honours <c>mode</c>, and the revision engine can
/// resolve every marker WordprocessingML defines - not only the two run wrappers.
/// A structural edit written without markup under a tracked connection is the failure
/// these cover: the reviewer sees the text replacement redlined and the new clause
/// silently present.
/// </summary>
public class WordRedlineTests
{
    private static OfficeAgentClient Client() => new(new WordModule());

    // ── Structural verbs honour the mode ─────────────────────────────────

    [Fact]
    public void An_inserted_paragraph_is_a_redline_by_default()
    {
        var document = Apply(Blank(), new InsertOp
        {
            Target = new TextSpanAnchor { ParaId = "auto-0000", Expect = string.Empty },
            Position = InsertPosition.After,
            Text = "The Supplier shall indemnify the Client."
        });

        AssertValid(document);
        using var body = Body(document);

        var inserted = body.Root.Descendants<InsertedRun>().Single();
        Assert.Equal("The Supplier shall indemnify the Client.", inserted.InnerText);

        // The paragraph mark too: without it, rejecting the revision leaves a blank
        // paragraph where the clause was.
        var paragraph = inserted.Ancestors<Paragraph>().Single();
        Assert.NotNull(paragraph.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<Inserted>());
    }

    [Fact]
    public void An_inserted_paragraph_in_direct_mode_carries_no_markup()
    {
        var document = Apply(Blank(), new InsertOp
        {
            Target = new TextSpanAnchor { ParaId = "auto-0000", Expect = string.Empty },
            Position = InsertPosition.After,
            Text = "Plain.",
            Mode = ChangeMode.Direct
        });

        using var body = Body(document);
        Assert.Empty(body.Root.Descendants<InsertedRun>());
        Assert.Contains("Plain.", body.Root.InnerText);
    }

    [Fact]
    public void Added_table_rows_are_marked_inserted()
    {
        var withTable = Apply(Blank(), new InsertTableOp
        {
            Target = new TextSpanAnchor { ParaId = "auto-0000", Expect = string.Empty },
            Position = InsertPosition.After,
            Table = new TableData { Headers = new[] { "Region", "Q1" }, Rows = new[] { new[] { "NL", "41850" } } },
            Mode = ChangeMode.Direct
        });

        var document = Apply(withTable, new InsertTableRowsOp
        {
            Target = new NodeAnchor { Kind = "table", Path = "table#0" },
            Rows = new[] { new[] { "BE", "12300" } },
            Position = TablePosition.End
        });

        AssertValid(document);
        using var body = Body(document);

        var row = body.Root.Descendants<TableRow>().Single(r => r.InnerText.Contains("BE"));
        Assert.NotNull(row.TableRowProperties?.GetFirstChild<Inserted>());
        Assert.Contains(row.Descendants<InsertedRun>(), r => r.InnerText == "BE");
    }

    [Fact]
    public void A_removed_table_row_stays_struck_through_until_a_reviewer_accepts()
    {
        var withTable = Apply(Blank(), new InsertTableOp
        {
            Target = new TextSpanAnchor { ParaId = "auto-0000", Expect = string.Empty },
            Position = InsertPosition.After,
            Table = new TableData { Headers = new[] { "Region", "Q1" }, Rows = new[] { new[] { "NL", "41850" } } },
            Mode = ChangeMode.Direct
        });

        var document = Apply(withTable, new RemoveTableRowsOp
        {
            Target = new NodeAnchor { Kind = "table", Path = "table#0" },
            RowIndices = new[] { -1 }
        });

        AssertValid(document);
        using var body = Body(document);

        var rows = body.Root.Descendants<TableRow>().ToList();
        Assert.Equal(2, rows.Count);

        var marked = rows[1];
        Assert.NotNull(marked.TableRowProperties?.GetFirstChild<Deleted>());
        // The cell text moved into w:delText: a row whose cells still held live runs would
        // reject to a row of empty cells.
        Assert.Contains("NL", marked.Descendants<DeletedText>().Select(t => t.Text));
        Assert.DoesNotContain(marked.Descendants<Text>(), t => t.Text == "NL");
    }

    [Fact]
    public void A_removed_table_column_marks_its_cells_rather_than_dropping_them()
    {
        var withTable = Apply(Blank(), new InsertTableOp
        {
            Target = new TextSpanAnchor { ParaId = "auto-0000", Expect = string.Empty },
            Position = InsertPosition.After,
            Table = new TableData { Headers = new[] { "Region", "Q1" }, Rows = new[] { new[] { "NL", "41850" } } },
            Mode = ChangeMode.Direct
        });

        var document = Apply(withTable, new RemoveTableColumnsOp
        {
            Target = new NodeAnchor { Kind = "table", Path = "table#0" },
            ColumnIndices = new[] { 1 }
        });

        AssertValid(document);
        using var body = Body(document);

        // The grid still lines up while the revision is pending.
        Assert.All(body.Root.Descendants<TableRow>(), r => Assert.Equal(2, r.Elements<TableCell>().Count()));
        Assert.Equal(2, body.Root.Descendants<CellDeletion>().Count());
    }

    [Fact]
    public void Formatting_records_the_properties_it_replaced()
    {
        var seeded = Apply(Blank(), new ChangeTextOp
        {
            Target = new TextSpanAnchor { ParaId = "auto-0000", Expect = string.Empty },
            With = "Payment terms are net thirty days.",
            Mode = ChangeMode.Direct
        });

        var document = Apply(seeded, new FormatOp
        {
            Target = new TextSpanAnchor { ParaId = FirstParaId(seeded), Expect = "net thirty days" },
            Bold = true
        });

        AssertValid(document);
        using var body = Body(document);

        var run = body.Root.Descendants<Run>().Single(r => r.InnerText == "net thirty days");
        Assert.NotNull(run.RunProperties?.GetFirstChild<Bold>());

        var change = run.RunProperties!.GetFirstChild<RunPropertiesChange>();
        Assert.NotNull(change);
        // What "reject" restores: the run had no bold before.
        Assert.Null(change!.GetFirstChild<PreviousRunProperties>()?.GetFirstChild<Bold>());

        // w:rPrChange closes CT_RPr - anything appended after it is a document Word repairs.
        Assert.Equal("rPrChange", run.RunProperties!.ChildElements.Last().LocalName);
    }

    [Fact]
    public void Formatting_that_changes_nothing_records_no_revision()
    {
        var seeded = Apply(Blank(), new ChangeTextOp
        {
            Target = new TextSpanAnchor { ParaId = "auto-0000", Expect = string.Empty },
            With = "Already bold.",
            Mode = ChangeMode.Direct
        });

        var once = Apply(seeded, new FormatOp
        {
            Target = new TextSpanAnchor { ParaId = FirstParaId(seeded), Expect = string.Empty },
            Bold = true,
            Mode = ChangeMode.Direct
        });

        var twice = Apply(once, new FormatOp
        {
            Target = new TextSpanAnchor { ParaId = FirstParaId(once), Expect = string.Empty },
            Bold = true
        });

        using var body = Body(twice);
        Assert.Empty(body.Root.Descendants<RunPropertiesChange>());
    }

    [Fact]
    public void Filling_a_content_control_is_a_redline()
    {
        var document = Apply(DocxFactory.Contract(), new FillOp
        {
            Target = new StructuralAnchor { Tag = DocxFactory.ClientControlTag },
            Value = "Globex Inc."
        });

        AssertValid(document);
        using var body = Body(document);

        Assert.Contains(DocxFactory.ClientPlaceholder, body.Root.Descendants<DeletedText>().Select(t => t.Text));
        Assert.Contains(body.Root.Descendants<InsertedRun>(), r => r.InnerText == "Globex Inc.");
    }

    // ── Resolving every kind of revision ─────────────────────────────────

    [Fact]
    public void Accepting_everything_leaves_no_revision_behind()
    {
        var document = Redlined();

        var accepted = Apply(document, new RevisionOp
        {
            Target = new NodeAnchor { Kind = "revision", Path = "all" },
            Action = RevisionAction.Accept
        });

        AssertValid(accepted);
        Assert.Empty(RevisionPaths(accepted));

        using var body = Body(accepted);
        Assert.Contains("An inserted clause.", body.Root.InnerText);
        Assert.DoesNotContain("removed", body.Root.InnerText);
    }

    [Fact]
    public void Rejecting_an_inserted_paragraph_takes_the_paragraph_with_it()
    {
        var document = Apply(Blank(), new InsertOp
        {
            Target = new TextSpanAnchor { ParaId = "auto-0000", Expect = string.Empty },
            Position = InsertPosition.After,
            Text = "Second thoughts."
        });

        var rejected = Apply(document, new RevisionOp
        {
            Target = new NodeAnchor { Kind = "revision", Path = "all" },
            Action = RevisionAction.Reject
        });

        AssertValid(rejected);
        using var body = Body(rejected);

        Assert.DoesNotContain("Second thoughts.", body.Root.InnerText);
        // The blank document's single paragraph, and nothing left over from the insert.
        Assert.Single(body.Root.Elements<Paragraph>());
    }

    [Fact]
    public void Rejecting_a_formatting_revision_puts_the_old_properties_back()
    {
        var seeded = Apply(Blank(), new ChangeTextOp
        {
            Target = new TextSpanAnchor { ParaId = "auto-0000", Expect = string.Empty },
            With = "Restore me.",
            Mode = ChangeMode.Direct
        });

        var formatted = Apply(seeded, new FormatOp
        {
            Target = new TextSpanAnchor { ParaId = FirstParaId(seeded), Expect = string.Empty },
            Bold = true,
            Color = "FF0000"
        });

        var rejected = Apply(formatted, new RevisionOp
        {
            Target = new NodeAnchor { Kind = "revision", Path = "all" },
            Action = RevisionAction.Reject
        });

        AssertValid(rejected);
        using var body = Body(rejected);

        var run = body.Root.Descendants<Run>().Single(r => r.InnerText == "Restore me.");
        Assert.Null(run.RunProperties?.GetFirstChild<Bold>());
        Assert.Empty(body.Root.Descendants<RunPropertiesChange>());
    }

    [Fact]
    public void Rejecting_a_paragraph_format_leaves_an_unrelated_mark_revision_alone()
    {
        // w:pPrChange records CT_PPrBase, which excludes the paragraph mark's own run
        // properties. Clearing w:pPr wholesale on reject would take a pending paragraph-mark
        // insertion with it - a revision this operation was never addressing.
        var inserted = Apply(Blank(), new InsertOp
        {
            Target = new TextSpanAnchor { ParaId = "auto-0000", Expect = string.Empty },
            Position = InsertPosition.After,
            Text = "A new clause."
        });

        var paraId = Client()
            .Inspect(new StreamHandle(new MemoryStream(inserted)))
            .Paragraphs.Single(p => p.Text == "A new clause.").ParaId;

        var formatted = Apply(inserted, new FormatOp
        {
            Target = new TextSpanAnchor { ParaId = paraId, Expect = string.Empty },
            Alignment = "center"
        });

        var formatRevision = Client()
            .Inspect(new StreamHandle(new MemoryStream(formatted)))
            .Nodes.Single(n => n.Kind == "revision" && n.Path.StartsWith("paraFormat#"));

        var rejected = Apply(formatted, new RevisionOp
        {
            Target = new NodeAnchor { Kind = "revision", Path = formatRevision.Path },
            Action = RevisionAction.Reject
        });

        AssertValid(rejected);
        using var body = Body(rejected);

        var paragraph = body.Root.Elements<Paragraph>().Single(p => p.InnerText == "A new clause.");
        Assert.Null(paragraph.ParagraphProperties?.GetFirstChild<Justification>());
        Assert.NotNull(paragraph.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<Inserted>());
    }

    [Fact]
    public void A_deleted_row_is_removed_when_its_revision_is_accepted()
    {
        var withTable = Apply(Blank(), new InsertTableOp
        {
            Target = new TextSpanAnchor { ParaId = "auto-0000", Expect = string.Empty },
            Position = InsertPosition.After,
            Table = new TableData { Headers = new[] { "Region", "Q1" }, Rows = new[] { new[] { "NL", "41850" } } },
            Mode = ChangeMode.Direct
        });

        var marked = Apply(withTable, new RemoveTableRowsOp
        {
            Target = new NodeAnchor { Kind = "table", Path = "table#0" },
            RowIndices = new[] { -1 }
        });

        var accepted = Apply(marked, new RevisionOp
        {
            Target = new NodeAnchor { Kind = "revision", Path = "all" },
            Action = RevisionAction.Accept
        });

        AssertValid(accepted);
        using var body = Body(accepted);
        Assert.Single(body.Root.Descendants<TableRow>());
        Assert.Empty(RevisionPathsIn(body.Root));
    }

    [Fact]
    public void Revisions_can_be_resolved_one_author_at_a_time()
    {
        var mine = Apply(Blank(), new InsertOp
        {
            Target = new TextSpanAnchor { ParaId = "auto-0000", Expect = string.Empty },
            Position = InsertPosition.After,
            Text = "Mine."
        });

        var withTheirs = AddForeignRevision(mine, author: "Jane Doe", text: "Theirs.");

        var accepted = Apply(withTheirs, new RevisionOp
        {
            Target = new NodeAnchor { Kind = "revision", Path = "author:Jane Doe" },
            Action = RevisionAction.Accept
        });

        AssertValid(accepted);
        using var body = Body(accepted);

        // Jane's insertion is now ordinary text; the engine's own is still pending.
        Assert.Contains("Theirs.", body.Root.InnerText);
        Assert.DoesNotContain(body.Root.Descendants<InsertedRun>(), r => r.InnerText == "Theirs.");
        Assert.Contains(body.Root.Descendants<InsertedRun>(), r => r.InnerText == "Mine.");
    }

    [Fact]
    public void Inspect_names_every_kind_of_revision_with_its_author()
    {
        var nodes = Client()
            .Inspect(new StreamHandle(new MemoryStream(Redlined())))
            .Nodes.Where(n => n.Kind == "revision")
            .ToList();

        var tags = nodes.Select(n => n.Path.Split('#')[0]).Distinct().ToList();
        Assert.Contains("ins", tags);
        Assert.Contains("del", tags);
        Assert.Contains("markIns", tags);
        Assert.Contains("runFormat", tags);
        Assert.All(nodes, n => Assert.Contains("by OfficeAgent", n.Summary));
    }

    [Fact]
    public void An_id_that_matches_nothing_is_still_an_error()
    {
        var report = Client().Preview(
            new StreamHandle(new MemoryStream(Redlined())),
            new DocumentPlan
            {
                Operations = new PlanOperation[]
                {
                    new RevisionOp { Target = new NodeAnchor { Kind = "revision", Path = "ins#4242" } }
                }
            });

        Assert.False(report.IsValid);
        Assert.Contains(report.Errors, e => e.Code == ValidationErrorCodes.AnchorNotFound);
    }

    // ── The deck has no redline vocabulary ───────────────────────────────

    [Fact]
    public void A_deck_refuses_a_tracked_structural_edit_rather_than_writing_it_directly()
    {
        var deck = new PowerPointModule().CreateBlank();
        var client = new OfficeAgentClient(new PowerPointModule());

        var report = client.Preview(
            new StreamHandle(new MemoryStream(deck), "deck.pptx"),
            new DocumentPlan
            {
                Operations = new PlanOperation[]
                {
                    new InsertOp
                    {
                        Target = new TextSpanAnchor { ParaId = "slide256/shape2/p0", Expect = string.Empty },
                        Text = "A line.",
                        Mode = ChangeMode.Tracked
                    }
                }
            });

        Assert.False(report.IsValid);
        Assert.Contains(report.Errors, e => e.Message.Contains("no tracked-changes representation"));
    }

    [Fact]
    public void A_deck_takes_the_same_verb_without_a_mode()
    {
        var deck = new PowerPointModule().CreateBlank();
        var client = new OfficeAgentClient(new PowerPointModule());

        var report = client.Preview(
            new StreamHandle(new MemoryStream(deck), "deck.pptx"),
            new DocumentPlan
            {
                Operations = new PlanOperation[]
                {
                    new InsertOp
                    {
                        Target = new TextSpanAnchor { ParaId = "slide256/shape2/p0", Expect = string.Empty },
                        Text = "A line."
                    }
                }
            });

        Assert.True(report.IsValid, string.Join("; ", report.Errors.Select(e => e.Message)));
    }

    // ── Fixtures ─────────────────────────────────────────────────────────

    /// <summary>A document carrying one of every revision this engine writes.</summary>
    private static byte[] Redlined()
    {
        var seeded = Apply(Blank(), new ChangeTextOp
        {
            Target = new TextSpanAnchor { ParaId = "auto-0000", Expect = string.Empty },
            With = "A clause to be removed.",
            Mode = ChangeMode.Direct
        });

        var inserted = Apply(seeded, new InsertOp
        {
            Target = new TextSpanAnchor { ParaId = FirstParaId(seeded), Expect = string.Empty },
            Position = InsertPosition.After,
            Text = "An inserted clause."
        });

        var replaced = Apply(inserted, new ChangeTextOp
        {
            Target = new TextSpanAnchor { ParaId = FirstParaId(inserted), Expect = "to be removed" },
            With = "that stays"
        });

        return Apply(replaced, new FormatOp
        {
            Target = new TextSpanAnchor { ParaId = FirstParaId(replaced), Expect = "A clause" },
            Italic = true
        });
    }

    /// <summary>
    /// An insertion by someone else, written straight into the package - the author filter
    /// is only interesting when more than one author is present.
    /// </summary>
    private static byte[] AddForeignRevision(byte[] document, string author, string text)
    {
        var buffer = new MemoryStream();
        buffer.Write(document, 0, document.Length);
        buffer.Position = 0;

        using (var package = WordprocessingDocument.Open(buffer, isEditable: true))
        {
            var body = package.MainDocumentPart!.Document.Body!;
            var run = new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
            var inserted = new InsertedRun(run)
            {
                Id = "9001",
                Author = author,
                Date = new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc)
            };
            body.AppendChild(new Paragraph(inserted));
            package.MainDocumentPart.Document.Save();
        }

        return buffer.ToArray();
    }

    private static string FirstParaId(byte[] document) =>
        Client().Inspect(new StreamHandle(new MemoryStream(document))).Paragraphs.First().ParaId;

    private static byte[] Blank() => new WordModule().CreateBlank();

    private static byte[] Apply(byte[] document, PlanOperation operation)
    {
        using var applied = Client().Commit(
            new StreamHandle(new MemoryStream(document)),
            new DocumentPlan { Operations = new[] { operation } });

        Assert.True(applied.Committed,
            string.Join("; ", applied.Report.Errors.Select(e => $"{e.Code}: {e.Message}")));
        return applied.ToBytes();
    }

    private static IReadOnlyList<string> RevisionPaths(byte[] document)
    {
        using var body = Body(document);
        return RevisionPathsIn(body.Root).ToList();
    }

    private static IEnumerable<string> RevisionPathsIn(OpenXmlElement root) =>
        root.Descendants<OpenXmlElement>()
            .Where(e => e is InsertedRun or DeletedRun or MoveFromRun or MoveToRun
                          or CellInsertion or CellDeletion or RunPropertiesChange
                          or ParagraphPropertiesChange or TablePropertiesChange
                          or TableRowPropertiesChange or TableCellPropertiesChange
                     || (e is Inserted or Deleted &&
                         e.Parent is ParagraphMarkRunProperties or TableRowProperties))
            .Select(e => e.LocalName);

    /// <summary>Keeps the package open for the life of the assertions against its body.</summary>
    private sealed class OpenBody : IDisposable
    {
        private readonly MemoryStream _stream;
        private readonly WordprocessingDocument _document;

        public Body Root { get; }

        public OpenBody(byte[] bytes)
        {
            _stream = new MemoryStream(bytes);
            _document = WordprocessingDocument.Open(_stream, isEditable: false);
            Root = _document.MainDocumentPart!.Document.Body!;
        }

        public void Dispose()
        {
            _document.Dispose();
            _stream.Dispose();
        }
    }

    private static OpenBody Body(byte[] document) => new(document);

    private static void AssertValid(byte[] document)
    {
        using var stream = new MemoryStream(document);
        using var opened = WordprocessingDocument.Open(stream, isEditable: false);
        var problems = new OpenXmlValidator(FileFormatVersions.Office2019)
            .Validate(opened)
            .Select(e => $"{e.Path?.XPath}: {e.Description}")
            .ToList();
        Assert.True(problems.Count == 0, string.Join("; ", problems.Take(3)));
    }
}
