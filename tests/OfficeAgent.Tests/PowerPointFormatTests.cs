using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;
using DocFormat = OfficeAgent.Abstractions.DocumentFormat;

namespace OfficeAgent.Tests;

/// <summary>
/// Formatting on a slide: the shared <c>format</c> verb writes the DrawingML run and
/// paragraph properties PowerPoint renders, over exactly the anchored span, and refuses
/// the Word-only measures rather than accepting them and changing nothing.
/// </summary>
public class PowerPointFormatTests
{
    private static OfficeAgentClient Client() => new(new PowerPointModule());

    [Fact]
    public void Formatting_a_span_bolds_only_that_span()
    {
        var client = Client();
        var deck = PptxFactory.Deck();
        var anchor = AnchorFor(client, deck, "Acme Corp");

        var applied = Apply(client, deck, new FormatOp { Target = anchor, Bold = true });

        var runs = RunsOf(applied);

        // The span crosses a run boundary in the fixture, so isolation leaves two bold
        // runs. What matters is that the bold text is exactly the span - no more, no less.
        Assert.Equal("Acme Corp", string.Concat(runs.Where(r => r.Bold).Select(r => r.Text)));
        Assert.Contains(runs, r => r.Text.Contains("revenue grew") && !r.Bold);
        // The second, unanchored "Acme Corp" later in the paragraph stays untouched.
        Assert.Contains(runs, r => r.Text.Contains("EMEA") && !r.Bold);
        AssertValid(applied);
    }

    [Fact]
    public void An_empty_expect_formats_the_whole_paragraph()
    {
        var client = Client();
        var deck = PptxFactory.Deck();

        // How an agent styles a heading without restating its text.
        var applied = Apply(client, deck, new FormatOp
        {
            Target = new TextSpanAnchor { ParaId = "slide256/shape2/p0", Expect = string.Empty },
            Italic = true,
            SizeHalfPoints = 44
        });

        var title = RunsOf(applied).Single(r => r.Text == PptxFactory.TitleText);
        Assert.True(title.Italic);
        // Half-points on the wire, hundredths of a point in DrawingML.
        Assert.Equal(2200, title.SizeHundredths);
        AssertValid(applied);
    }

    [Fact]
    public void Colour_and_highlight_are_written_as_drawingml_fills()
    {
        var client = Client();
        var deck = PptxFactory.Deck();
        var anchor = AnchorFor(client, deck, "Acme Corp");

        var applied = Apply(client, deck, new FormatOp
        {
            Target = anchor,
            Color = "FF0000",
            Highlight = "yellow",
            FontFamily = "Georgia"
        });

        using var stream = new MemoryStream(applied);
        using var document = PresentationDocument.Open(stream, isEditable: false);
        // The span spreads over two runs after isolation; both must carry the formatting,
        // so assert on every run the span covers rather than on one of them.
        var covered = document.PresentationPart!.SlideParts
            .SelectMany(p => p.Slide.Descendants<A.Run>())
            .Where(r => r.RunProperties?.GetFirstChild<A.Highlight>() is not null)
            .ToList();

        Assert.Equal("Acme Corp", string.Concat(covered
            .Select(r => string.Concat(r.Elements<A.Text>().Select(t => t.Text)))));

        var properties = covered[0].RunProperties!;

        Assert.Equal("FF0000",
            properties.GetFirstChild<A.SolidFill>()!.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
        Assert.Equal("FFFF00",
            properties.GetFirstChild<A.Highlight>()!.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
        Assert.Equal("Georgia", properties.GetFirstChild<A.LatinFont>()!.Typeface!.Value);
        AssertValid(applied);
    }

    [Fact]
    public void Alignment_lands_on_the_paragraph()
    {
        var client = Client();

        var applied = Apply(client, PptxFactory.Deck(), new FormatOp
        {
            Target = new TextSpanAnchor { ParaId = "slide256/shape2/p0", Expect = string.Empty },
            Alignment = "center"
        });

        using var stream = new MemoryStream(applied);
        using var document = PresentationDocument.Open(stream, isEditable: false);
        var paragraph = document.PresentationPart!.SlideParts.First()
            .Slide.Descendants<A.Paragraph>()
            .First(p => string.Concat(p.Descendants<A.Text>().Select(t => t.Text)) == PptxFactory.TitleText);

        Assert.Equal(A.TextAlignmentTypeValues.Center, paragraph.ParagraphProperties!.Alignment!.Value);
        AssertValid(applied);
    }

    [Fact]
    public void Table_cell_text_can_be_formatted_like_any_other_paragraph()
    {
        var client = Client();
        var deck = PptxFactory.DeckWithTable();
        var anchor = AnchorFor(client, deck, "Region");

        var applied = Apply(client, deck, new FormatOp { Target = anchor, Bold = true });

        Assert.Contains(RunsOf(applied), r => r.Text == "Region" && r.Bold);
        AssertValid(applied);
    }

    [Fact]
    public void An_image_is_resized_through_the_same_verb()
    {
        var client = Client();
        var withImage = Apply(client, PptxFactory.Deck(), new InsertImageOp
        {
            Target = new NodeAnchor { Kind = "slide", Path = "slide#256" },
            Base64Bytes = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==",
            WidthPx = 100,
            HeightPx = 100
        });

        var resized = Apply(client, withImage, new FormatOp
        {
            Target = new NodeAnchor { Kind = "image", Path = "image#256/4" },
            WidthPx = 400,
            HeightPx = 250
        });

        using var stream = new MemoryStream(resized);
        using var document = PresentationDocument.Open(stream, isEditable: false);
        var extents = document.PresentationPart!.SlideParts
            .SelectMany(p => p.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Picture>())
            .Select(p => p.ShapeProperties!.Transform2D!.Extents!)
            .Single();

        Assert.Equal(400 * 9525L, extents.Cx!.Value);
        Assert.Equal(250 * 9525L, extents.Cy!.Value);
        AssertValid(resized);
    }

    [Fact]
    public void Word_only_measures_are_refused_rather_than_quietly_ignored()
    {
        var client = Client();
        var deck = PptxFactory.Deck();
        var anchor = AnchorFor(client, deck, "Acme Corp");

        // An agent that sees these succeed would believe the deck changed.
        foreach (var op in new[]
                 {
                     new FormatOp { Target = anchor, IndentLeftTwips = 720 },
                     new FormatOp { Target = anchor, SpacingAfterTwips = 120 },
                     new FormatOp { Target = anchor, BorderStyle = "single" },
                     new FormatOp { Target = anchor, StyleId = "Heading1" }
                 })
        {
            var report = Preview(client, deck, op);
            var error = Assert.Single(report.Errors);
            Assert.Equal(ValidationErrorCodes.InvalidOperation, error.Code);
            // The message must say what IS available, not merely what is not.
            Assert.Contains("bold", error.Message);
        }
    }

    [Fact]
    public void A_format_with_nothing_to_change_and_a_drifted_anchor_are_both_reported()
    {
        var client = Client();
        var deck = PptxFactory.Deck();

        var empty = Preview(client, deck, new FormatOp
        {
            Target = AnchorFor(client, deck, "Acme Corp")
        });
        var drifted = Preview(client, deck, new FormatOp
        {
            Target = new TextSpanAnchor { ParaId = "slide256/shape3/p0", Expect = "Not in the deck" },
            Bold = true
        });

        Assert.Equal(ValidationErrorCodes.InvalidOperation, Assert.Single(empty.Errors).Code);
        Assert.Equal(ValidationErrorCodes.ExpectMismatch, Assert.Single(drifted.Errors).Code);
    }

    // ── helpers ───────────────────────────────────────────────────────────────

    private static TextSpanAnchor AnchorFor(OfficeAgentClient client, byte[] deck, string pattern) =>
        (TextSpanAnchor)client.Find(
            new StreamHandle(new MemoryStream(deck)),
            new FindQuery { Pattern = pattern })[0].Anchor!;

    private static byte[] Apply(OfficeAgentClient client, byte[] deck, PlanOperation operation)
    {
        using var applied = client.Commit(new StreamHandle(new MemoryStream(deck)), Plan(operation));
        Assert.True(applied.Committed,
            string.Join("; ", applied.Report.Errors.Select(e => $"{e.Code}: {e.Message}")));
        return applied.ToBytes();
    }

    private static ChangeReport Preview(OfficeAgentClient client, byte[] deck, PlanOperation operation) =>
        client.Preview(new StreamHandle(new MemoryStream(deck)), Plan(operation));

    private static DocumentPlan Plan(PlanOperation operation) => new()
    {
        Format = DocFormat.PowerPoint,
        Operations = new[] { operation }
    };

    private static List<(string Text, bool Bold, bool Italic, int SizeHundredths)> RunsOf(byte[] deck)
    {
        using var stream = new MemoryStream(deck);
        using var document = PresentationDocument.Open(stream, isEditable: false);

        return document.PresentationPart!.SlideParts
            .SelectMany(p => p.Slide.Descendants<A.Run>())
            .Select(r => (
                Text: string.Concat(r.Elements<A.Text>().Select(t => t.Text)),
                Bold: r.RunProperties?.Bold?.Value == true,
                Italic: r.RunProperties?.Italic?.Value == true,
                SizeHundredths: r.RunProperties?.FontSize?.Value ?? 0))
            .ToList();
    }

    private static void AssertValid(byte[] deck)
    {
        using var stream = new MemoryStream(deck);
        using var document = PresentationDocument.Open(stream, isEditable: false);
        var problems = new OpenXmlValidator(FileFormatVersions.Office2019)
            .Validate(document)
            .Select(e => $"{e.Path?.XPath}: {e.Description}")
            .ToList();
        Assert.True(problems.Count == 0, string.Join("; ", problems.Take(3)));
    }
}
