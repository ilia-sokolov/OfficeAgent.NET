using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.PowerPoint;
using OfficeAgent.Word;
using DocFormat = OfficeAgent.Abstractions.DocumentFormat;

namespace OfficeAgent.Tests;

/// <summary>
/// The PowerPoint module: a deck inspects into slide-scoped anchors, find returns
/// content-verified hits across slides and notes, and text edits survive round-tripping
/// without disturbing the formatting of runs they do not cover.
/// </summary>
public class PowerPointModuleTests
{
    private static OfficeAgentClient Client() => new(new PowerPointModule());

    [Fact]
    public void Fixture_deck_is_schema_valid()
    {
        // The fixture is hand-built, so an invalid one would make every other assertion
        // in this file meaningless.
        using var stream = new MemoryStream(PptxFactory.Deck());
        using var document = PresentationDocument.Open(stream, isEditable: false);

        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(document).ToList();
        Assert.True(errors.Count == 0,
            string.Join("; ", errors.Select(e => $"{e.Path?.XPath}: {e.Description}")));
    }

    [Fact]
    public void Inspect_reports_slides_notes_and_addressable_paragraphs()
    {
        var inspection = Client().Inspect(PptxFactory.Deck());

        Assert.Equal(DocFormat.PowerPoint, inspection.Format);
        Assert.False(string.IsNullOrEmpty(inspection.Snapshot.ETag));

        // One outline entry per slide, titled by the title placeholder.
        Assert.Equal(2, inspection.Outline.Count);
        Assert.Equal(PptxFactory.TitleText, inspection.Outline[0].Text);
        Assert.Equal(PptxFactory.SecondSlideTitle, inspection.Outline[1].Text);

        // Paragraph ids name the slide and the shape they live in.
        var title = Assert.Single(inspection.Paragraphs, p => p.Text == PptxFactory.TitleText);
        Assert.StartsWith("slide256/shape2/p", title.ParaId);
        Assert.Equal("slide", title.Location);

        // Notes are addressable and marked as such, not silently folded into the slide.
        var notes = Assert.Single(inspection.Paragraphs, p => p.Text == PptxFactory.NotesText);
        Assert.Contains("notes/", notes.ParaId);
        Assert.Equal("notes", notes.Location);

        // Slides surface as nodes so structure is visible without paragraph text.
        Assert.Equal(2, inspection.Nodes.Count(n => n.Kind == "slide"));
    }

    [Fact]
    public void Inspect_at_outline_fidelity_skips_paragraph_text()
    {
        var inspection = Client().Inspect(
            PptxFactory.Deck(), new InspectOptions { Fidelity = Fidelity.Outline });

        Assert.Equal(2, inspection.Outline.Count);
        Assert.Empty(inspection.Paragraphs);
    }

    [Fact]
    public void Find_returns_content_verified_anchors_across_slides_and_notes()
    {
        var hits = Client().Find(
            new StreamHandle(new MemoryStream(PptxFactory.Deck())),
            new FindQuery { Pattern = "Acme Corp" });

        // Twice in the slide-1 bullet, once in the notes.
        Assert.Equal(3, hits.Count);
        Assert.All(hits, h => Assert.Equal("Acme Corp", h.Text));

        // The two hits inside one paragraph are distinguished by occurrence, which is
        // what makes them separately addressable.
        var bullet = hits.Where(h => ((TextSpanAnchor)h.Anchor!).ParaId.Contains("shape3")).ToList();
        Assert.Equal(2, bullet.Count);
        Assert.Equal(new[] { 0, 1 }, bullet.Select(h => ((TextSpanAnchor)h.Anchor!).Occurrence));
    }

    [Fact]
    public void Change_text_replaces_a_run_spanning_span_and_keeps_the_rest()
    {
        var client = Client();
        var deck = PptxFactory.Deck();
        var anchor = FirstAnchorFor(client, deck, "Acme Corp");

        var plan = new DocumentPlan
        {
            Format = DocFormat.PowerPoint,
            Operations = new PlanOperation[]
            {
                new ChangeTextOp { Target = anchor, With = "Globex Inc.", Mode = ChangeMode.Direct }
            }
        };

        using var applied = client.Commit(new StreamHandle(new MemoryStream(deck)), plan);
        Assert.True(applied.Committed,
            string.Join("; ", applied.Report.Errors.Select(e => $"{e.Code}: {e.Message}")));

        // "Acme Corp" spans a run boundary in the fixture, so this only passes if the
        // span was isolated across runs rather than matched inside a single one.
        var after = client.Inspect(applied.ToBytes());
        var bullet = Assert.Single(after.Paragraphs, p => p.ParaId == anchor.ParaId);
        Assert.Equal("Globex Inc. revenue grew in every region except Acme Corp EMEA.", bullet.Text);
    }

    [Fact]
    public void Replacing_text_inside_a_styled_run_keeps_that_run_s_formatting()
    {
        var client = Client();
        var deck = PptxFactory.DeckWithBoldRun();
        // "BOLD" sits inside a bold run whose remaining text must stay bold, between two
        // plain runs that must stay plain. Replacing only part of a run is where
        // character formatting is easiest to lose.
        var anchor = FirstAnchorFor(client, deck, "BOLD");

        using var applied = client.Commit(
            new StreamHandle(new MemoryStream(deck)),
            new DocumentPlan
            {
                Format = DocFormat.PowerPoint,
                Operations = new PlanOperation[]
                {
                    new ChangeTextOp { Target = anchor, With = "CHANGED", Mode = ChangeMode.Direct }
                }
            });

        Assert.True(applied.Committed,
            string.Join("; ", applied.Report.Errors.Select(e => $"{e.Code}: {e.Message}")));

        using var stream = new MemoryStream(applied.ToBytes());
        using var document = DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(stream, isEditable: false);
        var runs = document.PresentationPart!.SlideParts.First()
            .Slide.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>()
            .SelectMany(p => p.Elements<DocumentFormat.OpenXml.Drawing.Run>())
            .Select(r => (
                Text: string.Concat(r.Elements<DocumentFormat.OpenXml.Drawing.Text>().Select(t => t.Text)),
                Bold: r.RunProperties?.Bold?.Value == true))
            .ToList();

        // The replacement inherits the bold run it landed in…
        Assert.Contains(runs, r => r.Text == "CHANGED" && r.Bold);
        // …the rest of that run stays bold…
        Assert.Contains(runs, r => r.Text == " text" && r.Bold);
        // …and the untouched neighbours stay unbolded.
        Assert.Contains(runs, r => r.Text == "plain " && !r.Bold);
        Assert.Contains(runs, r => r.Text == " tail" && !r.Bold);
    }

    [Fact]
    public void Writing_into_an_empty_placeholder_is_how_a_new_deck_gets_its_title()
    {
        var client = Client();
        var blank = new PowerPointModule().CreateBlank();

        // A new deck's title placeholder is empty, and a slide has no paragraph-inserting
        // verb, so an empty 'expect' is the only route to first text. It stays
        // content-verified: the empty expect asserts the paragraph really is blank.
        using var applied = client.Commit(
            new StreamHandle(new MemoryStream(blank)),
            new DocumentPlan
            {
                Operations = new PlanOperation[]
                {
                    new ChangeTextOp
                    {
                        Target = new TextSpanAnchor { ParaId = "slide256/shape2/p0", Expect = "" },
                        With = "Quarterly Review",
                        Mode = ChangeMode.Direct
                    }
                }
            });

        Assert.True(applied.Committed,
            string.Join("; ", applied.Report.Errors.Select(e => $"{e.Code}: {e.Message}")));

        var inspection = client.Inspect(applied.ToBytes());
        Assert.Equal("Quarterly Review", Assert.Single(inspection.Paragraphs).Text);
    }

    [Fact]
    public void An_empty_expect_against_a_paragraph_that_has_text_is_treated_as_drift()
    {
        var client = Client();

        // The same operation on a paragraph that is NOT empty must fail rather than
        // silently overwriting whatever was there.
        var report = client.Preview(
            new StreamHandle(new MemoryStream(PptxFactory.Deck())),
            new DocumentPlan
            {
                Operations = new PlanOperation[]
                {
                    new ChangeTextOp
                    {
                        Target = new TextSpanAnchor { ParaId = "slide256/shape2/p0", Expect = "" },
                        With = "Replaced",
                        Mode = ChangeMode.Direct
                    }
                }
            });

        Assert.Equal(ValidationErrorCodes.ExpectMismatch, Assert.Single(report.Errors).Code);
    }

    [Fact]
    public void A_plan_that_names_no_format_applies_to_either_format()
    {
        // The tools document a plan as a bare operations list, so a plan with no format
        // must not be silently bound to Word - that made every deck edit through the
        // agent surface fail with contract-mismatch.
        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new ChangeTextOp
                {
                    Target = new TextSpanAnchor { ParaId = "slide256/shape2/p0", Expect = "" },
                    With = "Titled",
                    Mode = ChangeMode.Direct
                }
            }
        };

        Assert.Equal(DocFormat.Unspecified, plan.Format);
        var report = Client().Preview(
            new StreamHandle(new MemoryStream(new PowerPointModule().CreateBlank())), plan);
        Assert.True(report.IsValid, string.Join("; ", report.Errors.Select(e => e.Message)));
    }

    [Fact]
    public void A_plan_that_names_the_wrong_format_is_still_rejected()
    {
        // The assertion is only meaningful if a deliberate mismatch still fails.
        var report = Client().Preview(
            new StreamHandle(new MemoryStream(PptxFactory.Deck())),
            new DocumentPlan
            {
                Format = DocFormat.Word,
                Operations = new PlanOperation[]
                {
                    new ChangeTextOp
                    {
                        Target = new TextSpanAnchor { ParaId = "slide256/shape2/p0", Expect = "x" },
                        With = "y"
                    }
                }
            });

        Assert.Contains(report.Errors, e => e.Code == ValidationErrorCodes.ContractMismatch);
    }

    [Fact]
    public void Change_text_targets_the_requested_occurrence()
    {
        var client = Client();
        var deck = PptxFactory.Deck();
        var first = FirstAnchorFor(client, deck, "Acme Corp");
        var second = new TextSpanAnchor { ParaId = first.ParaId, Expect = "Acme Corp", Occurrence = 1 };

        var plan = new DocumentPlan
        {
            Format = DocFormat.PowerPoint,
            Operations = new PlanOperation[]
            {
                new ChangeTextOp { Target = second, With = "Globex", Mode = ChangeMode.Direct }
            }
        };

        using var applied = client.Commit(new StreamHandle(new MemoryStream(deck)), plan);
        Assert.True(applied.Committed);

        var after = client.Inspect(applied.ToBytes());
        var bullet = Assert.Single(after.Paragraphs, p => p.ParaId == first.ParaId);
        Assert.Equal("Acme Corp revenue grew in every region except Globex EMEA.", bullet.Text);
    }

    [Fact]
    public void Edited_deck_stays_schema_valid()
    {
        var client = Client();
        var deck = PptxFactory.Deck();
        var plan = new DocumentPlan
        {
            Format = DocFormat.PowerPoint,
            Operations = new PlanOperation[]
            {
                new ChangeTextOp
                {
                    Target = FirstAnchorFor(client, deck, "Acme Corp"),
                    With = "Globex Inc.",
                    Mode = ChangeMode.Direct
                }
            }
        };

        using var applied = client.Commit(new StreamHandle(new MemoryStream(deck)), plan);

        using var stream = new MemoryStream(applied.ToBytes());
        using var document = PresentationDocument.Open(stream, isEditable: false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(document).ToList();
        Assert.True(errors.Count == 0,
            string.Join("; ", errors.Select(e => $"{e.Path?.XPath}: {e.Description}")));
    }

    [Fact]
    public void Tracked_mode_is_refused_rather_than_silently_written_as_a_direct_edit()
    {
        var client = Client();
        var deck = PptxFactory.Deck();
        var plan = new DocumentPlan
        {
            Format = DocFormat.PowerPoint,
            Operations = new PlanOperation[]
            {
                new ChangeTextOp
                {
                    Target = FirstAnchorFor(client, deck, "Acme Corp"),
                    With = "Globex Inc.",
                    Mode = ChangeMode.Tracked
                }
            }
        };

        var report = client.Preview(new StreamHandle(new MemoryStream(deck)), plan);

        // PresentationML has no redline vocabulary; honouring Tracked by writing a direct
        // edit would misrepresent what the deck contains.
        Assert.False(report.IsValid);
        Assert.Contains("Tracked", Assert.Single(report.Errors).Message);
    }

    [Fact]
    public void A_drifted_anchor_fails_without_touching_the_deck()
    {
        var client = Client();
        var deck = PptxFactory.Deck();
        var plan = new DocumentPlan
        {
            Format = DocFormat.PowerPoint,
            Operations = new PlanOperation[]
            {
                new ChangeTextOp
                {
                    Target = new TextSpanAnchor
                    {
                        ParaId = "slide256/shape3/p0",
                        Expect = "Text that is not there"
                    },
                    With = "x",
                    Mode = ChangeMode.Direct
                }
            }
        };

        using var applied = client.Commit(new StreamHandle(new MemoryStream(deck)), plan);

        Assert.False(applied.Committed);
        Assert.Equal(ValidationErrorCodes.ExpectMismatch, Assert.Single(applied.Report.Errors).Code);
    }

    [Fact]
    public void An_unsupported_verb_is_reported_rather_than_half_applied()
    {
        var client = Client();
        var deck = PptxFactory.Deck();
        var plan = new DocumentPlan
        {
            Format = DocFormat.PowerPoint,
            Operations = new PlanOperation[]
            {
                new RevisionOp
                {
                    Target = new NodeAnchor { Kind = "revision", Path = "all" },
                    Action = RevisionAction.Accept
                }
            }
        };

        using var applied = client.Commit(new StreamHandle(new MemoryStream(deck)), plan);

        Assert.False(applied.Committed);
        Assert.Equal(ValidationErrorCodes.UnsupportedOperation, Assert.Single(applied.Report.Errors).Code);
    }

    [Fact]
    public void One_client_serves_both_formats_and_routes_each_document_to_its_module()
    {
        // The engine picks the module by package format, so a host that registers both
        // needs no per-call switch.
        var client = new OfficeAgentClient(new WordModule(), new PowerPointModule());

        Assert.Equal(DocFormat.Word, client.Inspect(DocxFactory.Contract()).Format);
        Assert.Equal(DocFormat.PowerPoint, client.Inspect(PptxFactory.Deck()).Format);
    }

    /// <summary>The first anchor the engine issues for some text, as an agent would obtain it.</summary>
    private static TextSpanAnchor FirstAnchorFor(OfficeAgentClient client, byte[] deck, string pattern)
    {
        var hits = client.Find(new StreamHandle(new MemoryStream(deck)), new FindQuery { Pattern = pattern });
        return (TextSpanAnchor)hits.First().Anchor!;
    }
}
