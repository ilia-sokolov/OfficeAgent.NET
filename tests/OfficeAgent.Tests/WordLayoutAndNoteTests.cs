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
/// Page geometry, breaks, and notes: the parts of a Word document that are not text and
/// that no other verb reaches. <c>w:sectPr</c> is a strict sequence and the notes parts
/// have entries Word treats as mandatory, so every test here validates the saved package
/// rather than only the element it wrote.
/// </summary>
public class WordLayoutAndNoteTests
{
    private static OfficeAgentClient Client() => new(new WordModule());

    // ── Page setup ───────────────────────────────────────────────────────

    [Fact]
    public void PageSetup_turns_the_section_sideways_and_sets_its_margins()
    {
        var document = Apply(Blank(), new PageSetupOp
        {
            PaperSize = "A4",
            Orientation = PageOrientation.Landscape,
            MarginTopTwips = 720,
            MarginLeftTwips = 1080
        });

        AssertValid(document);
        using var body = Body(document);

        var size = body.Root.GetFirstChild<SectionProperties>()!.GetFirstChild<PageSize>()!;
        // A4 is 11906 × 16838 portrait; landscape swaps them.
        Assert.Equal(16838u, size.Width?.Value);
        Assert.Equal(11906u, size.Height?.Value);
        Assert.Equal(PageOrientationValues.Landscape, size.Orient?.Value);

        var margin = body.Root.GetFirstChild<SectionProperties>()!.GetFirstChild<PageMargin>()!;
        Assert.Equal(720, margin.Top?.Value);
        Assert.Equal(1080u, margin.Left?.Value);
        // The edges the caller did not name keep Word's own defaults rather than resetting.
        Assert.Equal(1440, margin.Bottom?.Value);
        Assert.Equal(1440u, margin.Right?.Value);
    }

    [Fact]
    public void PageSetup_leaves_the_margins_it_was_not_asked_about()
    {
        var once = Apply(Blank(), new PageSetupOp { MarginLeftTwips = 2880 });
        var twice = Apply(once, new PageSetupOp { MarginRightTwips = 360 });

        AssertValid(twice);
        using var body = Body(twice);

        var margin = body.Root.GetFirstChild<SectionProperties>()!.GetFirstChild<PageMargin>()!;
        Assert.Equal(2880u, margin.Left?.Value);
        Assert.Equal(360u, margin.Right?.Value);
    }

    [Fact]
    public void PageSetup_refuses_a_paper_size_and_explicit_dimensions_together()
    {
        var report = Preview(Blank(), new PageSetupOp { PaperSize = "A4", PageWidthTwips = 9000 });

        Assert.False(report.IsValid);
        Assert.Contains(report.Errors, e => e.Code == ValidationErrorCodes.InvalidOperation);
    }

    [Fact]
    public void PageSetup_names_the_paper_sizes_it_knows()
    {
        var report = Preview(Blank(), new PageSetupOp { PaperSize = "Foolscap" });

        Assert.False(report.IsValid);
        Assert.Contains(report.Errors, e => e.Message.Contains("A4") && e.Message.Contains("Letter"));
    }

    [Fact]
    public void PageSetup_with_nothing_set_is_refused_rather_than_silently_doing_nothing()
    {
        var report = Preview(Blank(), new PageSetupOp());

        Assert.False(report.IsValid);
        Assert.Contains(report.Errors, e => e.Code == ValidationErrorCodes.InvalidOperation);
    }

    // ── Breaks ───────────────────────────────────────────────────────────

    [Fact]
    public void A_page_break_arrives_in_a_paragraph_of_its_own_as_a_redline()
    {
        var document = Apply(Blank(), new InsertBreakOp
        {
            Target = new TextSpanAnchor { ParaId = "auto-0000", Expect = string.Empty },
            Kind = BreakKind.Page,
            Position = InsertPosition.After
        });

        AssertValid(document);
        using var body = Body(document);

        var wrapper = body.Root.Descendants<InsertedRun>().Single();
        Assert.Equal(BreakValues.Page, wrapper.Descendants<Break>().Single().Type?.Value);
        Assert.NotNull(wrapper.Ancestors<Paragraph>().Single()
            .ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<Inserted>());
    }

    [Fact]
    public void A_section_break_clones_the_section_it_splits()
    {
        var shaped = Apply(Blank(), new PageSetupOp
        {
            PaperSize = "A4",
            MarginLeftTwips = 2000
        });

        var document = Apply(shaped, new InsertBreakOp
        {
            Target = new TextSpanAnchor { ParaId = FirstParaId(shaped), Expect = string.Empty },
            Kind = BreakKind.SectionNextPage,
            Position = InsertPosition.After,
            Mode = ChangeMode.Direct
        });

        AssertValid(document);
        using var body = Body(document);

        var inline = body.Root.Descendants<Paragraph>()
            .Select(p => p.ParagraphProperties?.SectionProperties)
            .Single(s => s is not null)!;

        // The pages before the break keep the geometry they had.
        Assert.Equal(11906u, inline.GetFirstChild<PageSize>()?.Width?.Value);
        Assert.Equal(2000u, inline.GetFirstChild<PageMargin>()?.Left?.Value);
        Assert.Equal(SectionMarkValues.NextPage, inline.GetFirstChild<SectionType>()?.Val?.Value);
    }

    [Fact]
    public void PageSetup_after_a_break_reaches_only_the_section_it_targets()
    {
        var split = Apply(Blank(), new InsertBreakOp
        {
            Target = new TextSpanAnchor { ParaId = "auto-0000", Expect = string.Empty },
            Kind = BreakKind.SectionNextPage,
            Position = InsertPosition.After,
            Mode = ChangeMode.Direct
        });

        // The first paragraph belongs to the first section, which the break paragraph ends.
        var document = Apply(split, new PageSetupOp
        {
            Target = new TextSpanAnchor { ParaId = FirstParaId(split), Expect = string.Empty },
            Orientation = PageOrientation.Landscape
        });

        AssertValid(document);
        using var body = Body(document);

        var inline = body.Root.Descendants<Paragraph>()
            .Select(p => p.ParagraphProperties?.SectionProperties)
            .Single(s => s is not null)!;
        Assert.Equal(PageOrientationValues.Landscape, inline.GetFirstChild<PageSize>()?.Orient?.Value);

        // The final section, which governs everything after the break, is untouched.
        var final = body.Root.GetFirstChild<SectionProperties>()!;
        Assert.NotEqual(PageOrientationValues.Landscape, final.GetFirstChild<PageSize>()?.Orient?.Value);
    }

    [Fact]
    public void A_section_break_is_refused_outside_the_body()
    {
        var withFooter = Apply(Blank(), new HeaderFooterOp { Footer = "Confidential" });
        var footerParaId = Client()
            .Inspect(new StreamHandle(new MemoryStream(withFooter)))
            .Paragraphs.First(p => p.Location == "footer").ParaId;

        var report = Preview(withFooter, new InsertBreakOp
        {
            Target = new TextSpanAnchor { ParaId = footerParaId, Expect = string.Empty },
            Kind = BreakKind.SectionNextPage
        });

        Assert.False(report.IsValid);
        Assert.Contains(report.Errors, e => e.Code == ValidationErrorCodes.InvalidOperation);
    }

    // ── Footnotes and endnotes ───────────────────────────────────────────

    [Fact]
    public void Adding_a_footnote_builds_the_part_the_reference_and_the_styles()
    {
        var seeded = Seeded("Fees are payable monthly in arrears.");

        var document = Apply(seeded, new NoteOp
        {
            Target = new TextSpanAnchor { ParaId = FirstParaId(seeded), Expect = "monthly in arrears" },
            Kind = NoteKind.Footnote,
            Text = "Subject to clause 8.2.",
            Mode = ChangeMode.Direct
        });

        AssertValid(document);

        using var stream = new MemoryStream(document);
        using var package = WordprocessingDocument.Open(stream, isEditable: false);
        var main = package.MainDocumentPart!;

        // The separators Word treats as mandatory, plus the note itself.
        var notes = main.FootnotesPart!.Footnotes!.Elements<Footnote>().ToList();
        Assert.Contains(notes, n => n.Type?.Value == FootnoteEndnoteValues.Separator);
        Assert.Contains(notes, n => n.Type?.Value == FootnoteEndnoteValues.ContinuationSeparator);

        var note = notes.Single(n => n.Id?.Value == 1);
        Assert.Contains("Subject to clause 8.2.", note.InnerText);

        var reference = main.Document.Body!.Descendants<FootnoteReference>().Single();
        Assert.Equal(1, reference.Id?.Value);

        // A pStyle with no definition renders as body text, which is a note that does not
        // look like one.
        var styles = main.StyleDefinitionsPart!.Styles!.Elements<Style>().Select(s => s.StyleId?.Value).ToList();
        Assert.Contains("FootnoteText", styles);
        Assert.Contains("FootnoteReference", styles);
    }

    [Fact]
    public void A_footnote_reference_lands_right_after_the_text_it_annotates()
    {
        var seeded = Seeded("The Supplier warrants the Deliverables for twelve months.");

        var document = Apply(seeded, new NoteOp
        {
            Target = new TextSpanAnchor { ParaId = FirstParaId(seeded), Expect = "twelve months" },
            Text = "Extendable by agreement.",
            Mode = ChangeMode.Direct
        });

        using var body = Body(document);
        var paragraph = body.Root.Elements<Paragraph>().First();
        var runs = paragraph.Elements<Run>().ToList();

        var annotated = runs.FindIndex(r => r.InnerText == "twelve months");
        Assert.True(annotated >= 0);
        Assert.NotNull(runs[annotated + 1].GetFirstChild<FootnoteReference>());
    }

    [Fact]
    public void Inspect_lists_notes_and_update_rewrites_one()
    {
        var withNote = Noted("Original wording.");

        var node = Client()
            .Inspect(new StreamHandle(new MemoryStream(withNote)))
            .Nodes.Single(n => n.Kind == "note");
        Assert.Equal("footnote#1", node.Path);
        Assert.Contains("Original wording.", node.Summary);

        var document = Apply(withNote, new NoteOp
        {
            Target = new NodeAnchor { Kind = "note", Path = "footnote#1" },
            Action = NoteAction.Update,
            Text = "Replaced wording.",
            Mode = ChangeMode.Direct
        });

        AssertValid(document);

        using var stream = new MemoryStream(document);
        using var package = WordprocessingDocument.Open(stream, isEditable: false);
        var note = package.MainDocumentPart!.FootnotesPart!.Footnotes!
            .Elements<Footnote>().Single(n => n.Id?.Value == 1);

        Assert.Contains("Replaced wording.", note.InnerText);
        Assert.DoesNotContain("Original wording.", note.InnerText);
        // The reference mark is the note's number, and survives a rewrite of its text.
        Assert.Single(note.Descendants<FootnoteReferenceMark>());
    }

    [Fact]
    public void Removing_a_note_takes_its_reference_with_it()
    {
        var document = Apply(Noted("Doomed."), new NoteOp
        {
            Target = new NodeAnchor { Kind = "note", Path = "footnote#1" },
            Action = NoteAction.Remove,
            Mode = ChangeMode.Direct
        });

        AssertValid(document);

        using var stream = new MemoryStream(document);
        using var package = WordprocessingDocument.Open(stream, isEditable: false);
        var main = package.MainDocumentPart!;

        Assert.Empty(main.Document.Body!.Descendants<FootnoteReference>());
        Assert.DoesNotContain(main.FootnotesPart!.Footnotes!.Elements<Footnote>(), n => n.Id?.Value == 1);
        Assert.DoesNotContain(main.Document.Body.Descendants<Run>(), r => !r.HasChildren);
    }

    [Fact]
    public void A_tracked_note_leaves_the_reference_and_the_wording_for_a_reviewer()
    {
        var seeded = Seeded("Interest accrues daily.");

        var document = Apply(seeded, new NoteOp
        {
            Target = new TextSpanAnchor { ParaId = FirstParaId(seeded), Expect = "daily" },
            Text = "At the statutory rate."
        });

        AssertValid(document);

        using var stream = new MemoryStream(document);
        using var package = WordprocessingDocument.Open(stream, isEditable: false);
        var main = package.MainDocumentPart!;

        var reference = main.Document.Body!.Descendants<FootnoteReference>().Single();
        Assert.NotNull(reference.Ancestors<InsertedRun>().FirstOrDefault());
        Assert.Contains(main.FootnotesPart!.Footnotes!.Descendants<InsertedRun>(),
            r => r.InnerText.Contains("At the statutory rate."));
    }

    [Fact]
    public void A_tracked_note_that_is_rejected_leaves_the_document_as_it_was()
    {
        var seeded = Seeded("Interest accrues daily.");

        var withNote = Apply(seeded, new NoteOp
        {
            Target = new TextSpanAnchor { ParaId = FirstParaId(seeded), Expect = "daily" },
            Text = "At the statutory rate."
        });

        var rejected = Apply(withNote, new RevisionOp
        {
            Target = new NodeAnchor { Kind = "revision", Path = "all" },
            Action = RevisionAction.Reject
        });

        AssertValid(rejected);
        using var body = Body(rejected);

        Assert.Empty(body.Root.Descendants<FootnoteReference>());
        Assert.Equal("Interest accrues daily.", body.Root.InnerText);
    }

    [Fact]
    public void A_tracked_note_removal_takes_the_note_when_it_is_accepted()
    {
        var accepted = Apply(Noted("Doomed."), new RevisionOp
        {
            Target = new NodeAnchor { Kind = "revision", Path = "all" },
            Action = RevisionAction.Accept
        });

        AssertValid(accepted);

        // Guard against a half-deleted note: the reference gone from the body while the
        // wording stays at the foot of the page, or the reverse.
        using var removed = new MemoryStream(Apply(
            Apply(accepted, new NoteOp
            {
                Target = new NodeAnchor { Kind = "note", Path = "footnote#1" },
                Action = NoteAction.Remove
            }),
            new RevisionOp
            {
                Target = new NodeAnchor { Kind = "revision", Path = "all" },
                Action = RevisionAction.Accept
            }));

        using var package = WordprocessingDocument.Open(removed, isEditable: false);
        var main = package.MainDocumentPart!;

        Assert.Empty(main.Document.Body!.Descendants<FootnoteReference>());
        Assert.DoesNotContain("Doomed.", main.FootnotesPart!.Footnotes!.InnerText);
        Assert.Empty(main.FootnotesPart.Footnotes.Elements<Footnote>().Where(n => n.Id?.Value == 1));
    }

    [Fact]
    public void An_endnote_is_built_the_same_way_in_its_own_part()
    {
        var seeded = Seeded("See the schedule for detail.");

        var document = Apply(seeded, new NoteOp
        {
            Target = new TextSpanAnchor { ParaId = FirstParaId(seeded), Expect = string.Empty },
            Kind = NoteKind.Endnote,
            Text = "Schedule 3, as amended.",
            Mode = ChangeMode.Direct
        });

        AssertValid(document);

        using var stream = new MemoryStream(document);
        using var package = WordprocessingDocument.Open(stream, isEditable: false);
        var main = package.MainDocumentPart!;

        Assert.Single(main.Document.Body!.Descendants<EndnoteReference>());
        Assert.Contains("Schedule 3, as amended.",
            main.EndnotesPart!.Endnotes!.Elements<Endnote>().Single(n => n.Id?.Value == 1).InnerText);

        var node = Client()
            .Inspect(new StreamHandle(new MemoryStream(document)))
            .Nodes.Single(n => n.Kind == "note");
        Assert.Equal("endnote#1", node.Path);
    }

    [Fact]
    public void A_second_footnote_numbers_itself_after_the_first()
    {
        var first = Noted("First.");
        var paraId = FirstParaId(first);

        var second = Apply(first, new NoteOp
        {
            Target = new TextSpanAnchor { ParaId = paraId, Expect = string.Empty },
            Text = "Second.",
            Mode = ChangeMode.Direct
        });

        AssertValid(second);

        var paths = Client()
            .Inspect(new StreamHandle(new MemoryStream(second)))
            .Nodes.Where(n => n.Kind == "note").Select(n => n.Path).ToList();
        Assert.Equal(new[] { "footnote#1", "footnote#2" }, paths);
    }

    [Fact]
    public void A_note_anchored_outside_the_body_is_refused()
    {
        var withFooter = Apply(Blank(), new HeaderFooterOp { Footer = "Confidential" });
        var footerParaId = Client()
            .Inspect(new StreamHandle(new MemoryStream(withFooter)))
            .Paragraphs.First(p => p.Location == "footer").ParaId;

        var report = Preview(withFooter, new NoteOp
        {
            Target = new TextSpanAnchor { ParaId = footerParaId, Expect = string.Empty },
            Text = "Nowhere to go."
        });

        Assert.False(report.IsValid);
        Assert.Contains(report.Errors, e => e.Code == ValidationErrorCodes.InvalidOperation);
    }

    [Fact]
    public void A_deck_reports_the_Word_only_verbs_by_name()
    {
        var deck = new PowerPointModule().CreateBlank();
        var client = new OfficeAgentClient(new PowerPointModule());

        foreach (PlanOperation operation in new PlanOperation[]
        {
            new PageSetupOp { Orientation = PageOrientation.Landscape },
            new InsertBreakOp
            {
                Target = new TextSpanAnchor { ParaId = "slide256/shape2/p0", Expect = string.Empty },
                Mode = ChangeMode.Direct
            },
            new NoteOp
            {
                Target = new TextSpanAnchor { ParaId = "slide256/shape2/p0", Expect = string.Empty },
                Text = "No such thing on a slide.",
                Mode = ChangeMode.Direct
            }
        })
        {
            var report = client.Preview(
                new StreamHandle(new MemoryStream(deck), "deck.pptx"),
                new DocumentPlan { Operations = new[] { operation } });

            Assert.False(report.IsValid);
            Assert.Contains(report.Errors, e =>
                e.Code == ValidationErrorCodes.UnsupportedOperation &&
                e.Message.Contains(operation.GetType().Name));
        }
    }

    // ── Fixtures ─────────────────────────────────────────────────────────

    private static byte[] Blank() => new WordModule().CreateBlank();

    private static byte[] Seeded(string text) => Apply(Blank(), new ChangeTextOp
    {
        Target = new TextSpanAnchor { ParaId = "auto-0000", Expect = string.Empty },
        With = text,
        Mode = ChangeMode.Direct
    });

    private static byte[] Noted(string text)
    {
        var seeded = Seeded("A clause worth a note.");
        return Apply(seeded, new NoteOp
        {
            Target = new TextSpanAnchor { ParaId = FirstParaId(seeded), Expect = "a note" },
            Text = text,
            Mode = ChangeMode.Direct
        });
    }

    private static string FirstParaId(byte[] document) =>
        Client().Inspect(new StreamHandle(new MemoryStream(document))).Paragraphs.First().ParaId;

    private static byte[] Apply(byte[] document, PlanOperation operation)
    {
        using var applied = Client().Commit(
            new StreamHandle(new MemoryStream(document)),
            new DocumentPlan { Operations = new[] { operation } });

        Assert.True(applied.Committed,
            string.Join("; ", applied.Report.Errors.Select(e => $"{e.Code}: {e.Message}")));
        return applied.ToBytes();
    }

    private static ChangeReport Preview(byte[] document, PlanOperation operation) =>
        Client().Preview(
            new StreamHandle(new MemoryStream(document)),
            new DocumentPlan { Operations = new[] { operation } });

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
