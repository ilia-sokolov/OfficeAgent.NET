using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.PowerPoint;
using OfficeAgent.Word;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;

namespace OfficeAgent.Tests;

/// <summary>
/// An image behind the content, in both formats, at a chosen strength.
/// </summary>
/// <remarks>
/// The two formats keep a background in completely different places - <c>p:bg</c> on the
/// slide, an anchored picture in the header for a page - so these assert the placement as
/// well as the presence. The opacity is the same mechanism in both:
/// <c>a:alphaModFix</c> on the blip.
/// </remarks>
public class BackgroundImageTests
{
    // A 1x1 PNG. The bytes only have to be a real image; nothing here renders one.
    private const string Png =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

    // ── PowerPoint ────────────────────────────────────────────────────────────

    [Fact]
    public void A_slide_takes_an_image_behind_it()
    {
        var client = new OfficeAgentClient(new PowerPointModule());

        var deck = ApplyDeck(client, Deck(), new BackgroundImageOp
        {
            Target = new NodeAnchor { Kind = "slide", Path = "slide#256" },
            Base64Bytes = Png,
            ImageType = "png"
        });

        var slide = SlideAt(deck, 0);
        var background = slide.CommonSlideData!.Background!;
        var blip = background.Descendants<A.Blip>().Single();

        Assert.NotNull(blip.Embed);

        // p:bg opens p:cSld. After the shape tree PowerPoint refuses the file.
        Assert.Equal(0, slide.CommonSlideData.ChildElements.ToList().IndexOf(background));

        // Full strength writes no alpha at all.
        Assert.Empty(blip.Descendants<A.AlphaModulationFixed>());
        AssertDeckValid(deck);
    }

    [Fact]
    public void Slide_opacity_is_written_as_an_alpha_on_the_blip()
    {
        var client = new OfficeAgentClient(new PowerPointModule());

        var deck = ApplyDeck(client, Deck(), new BackgroundImageOp
        {
            Target = new NodeAnchor { Kind = "slide", Path = "slide#256" },
            Base64Bytes = Png,
            Opacity = 0.2
        });

        var alpha = SlideAt(deck, 0).Descendants<A.AlphaModulationFixed>().Single();

        // DrawingML counts alpha in thousandths of a percent.
        Assert.Equal(20000, alpha.Amount!.Value);
        AssertDeckValid(deck);
    }

    [Fact]
    public void A_background_with_no_slide_named_covers_every_slide()
    {
        var client = new OfficeAgentClient(new PowerPointModule());

        var grown = ApplyDeck(client, Deck(),
            new InsertSlideOp { Slide = new SlideData { Layout = "titleOnly", Title = "Two" } },
            new InsertSlideOp { Slide = new SlideData { Layout = "titleOnly", Title = "Three" } });

        var painted = ApplyDeck(client, grown, new BackgroundImageOp { Base64Bytes = Png });

        using var stream = new MemoryStream(painted);
        using var document = PresentationDocument.Open(stream, isEditable: false);

        var slides = document.PresentationPart!.SlideParts.ToList();
        Assert.Equal(3, slides.Count);
        Assert.All(slides, part => Assert.NotNull(part.Slide.CommonSlideData!.Background));

        // Each slide owns its own relationship rather than sharing one it cannot resolve.
        Assert.All(slides, part => Assert.Single(part.ImageParts));
        AssertDeckValid(painted);
    }

    [Fact]
    public void A_slide_background_can_be_taken_away_again()
    {
        var client = new OfficeAgentClient(new PowerPointModule());

        var painted = ApplyDeck(client, Deck(), new BackgroundImageOp
        {
            Target = new NodeAnchor { Kind = "slide", Path = "slide#256" },
            Base64Bytes = Png
        });
        var cleared = ApplyDeck(client, painted, new BackgroundImageOp
        {
            Target = new NodeAnchor { Kind = "slide", Path = "slide#256" }
        });

        Assert.Null(SlideAt(cleared, 0).CommonSlideData!.Background);
        AssertDeckValid(cleared);
    }

    [Fact]
    public void An_opacity_outside_zero_to_one_is_refused()
    {
        var client = new OfficeAgentClient(new PowerPointModule());

        var report = client.Preview(
            new StreamHandle(new MemoryStream(Deck())),
            new DocumentPlan
            {
                Format = OfficeAgent.Abstractions.DocumentFormat.PowerPoint,
                Operations = new PlanOperation[]
                {
                    new BackgroundImageOp { Base64Bytes = Png, Opacity = 1.5 }
                }
            });

        var error = Assert.Single(report.Errors);
        Assert.Equal(ValidationErrorCodes.InvalidOperation, error.Code);
        Assert.Contains("between 0 and 1", error.Message);
    }

    // ── Word ──────────────────────────────────────────────────────────────────

    [Fact]
    public void A_page_background_is_anchored_behind_the_header()
    {
        var client = new OfficeAgentClient(new WordModule());

        var document = ApplyDoc(client, Blank(), new BackgroundImageOp
        {
            Base64Bytes = Png,
            Opacity = 0.15
        });

        using var stream = new MemoryStream(document);
        using var opened = WordprocessingDocument.Open(stream, isEditable: false);
        var header = Assert.Single(opened.MainDocumentPart!.HeaderParts);

        var anchor = Assert.Single(header.Header!.Descendants<DW.Anchor>());

        // Behind the text and out of the flow: the body runs over it, not around it.
        Assert.True(anchor.BehindDoc!.Value);
        Assert.Single(anchor.Descendants<DW.WrapNone>());

        // Pinned to the page rather than to the paragraph, so it does not move with text.
        Assert.Equal(DW.HorizontalRelativePositionValues.Page,
            anchor.Descendants<DW.HorizontalPosition>().Single().RelativeFrom!.Value);
        Assert.Equal(DW.VerticalRelativePositionValues.Page,
            anchor.Descendants<DW.VerticalPosition>().Single().RelativeFrom!.Value);

        // Letter, in EMU: 8.5in x 11in.
        var extent = anchor.Descendants<DW.Extent>().Single();
        Assert.Equal(12240L * 635L, extent.Cx!.Value);
        Assert.Equal(15840L * 635L, extent.Cy!.Value);

        Assert.Equal(15000, anchor.Descendants<A.AlphaModulationFixed>().Single().Amount!.Value);
        Assert.Single(header.ImageParts);
        AssertDocValid(document);
    }

    [Fact]
    public void Setting_a_page_background_twice_replaces_it()
    {
        var client = new OfficeAgentClient(new WordModule());

        var once = ApplyDoc(client, Blank(), new BackgroundImageOp { Base64Bytes = Png });
        var twice = ApplyDoc(client, once, new BackgroundImageOp { Base64Bytes = Png, Opacity = 0.5 });

        using var stream = new MemoryStream(twice);
        using var opened = WordprocessingDocument.Open(stream, isEditable: false);
        var header = Assert.Single(opened.MainDocumentPart!.HeaderParts);

        // One image, not two stacked ones nobody could tell apart.
        Assert.Single(header.Header!.Descendants<DW.Anchor>());
        Assert.Single(header.ImageParts);
        Assert.Equal(50000, header.Header.Descendants<A.AlphaModulationFixed>().Single().Amount!.Value);
        AssertDocValid(twice);
    }

    [Fact]
    public void A_page_background_can_be_taken_away_again()
    {
        var client = new OfficeAgentClient(new WordModule());

        var painted = ApplyDoc(client, Blank(), new BackgroundImageOp { Base64Bytes = Png });
        var cleared = ApplyDoc(client, painted, new BackgroundImageOp());

        using var stream = new MemoryStream(cleared);
        using var opened = WordprocessingDocument.Open(stream, isEditable: false);

        foreach (var header in opened.MainDocumentPart!.HeaderParts)
        {
            Assert.Empty(header.Header!.Descendants<DW.Anchor>());
            // The image bytes go with it rather than staying as an orphan in the package.
            Assert.Empty(header.ImageParts);
        }
        AssertDocValid(cleared);
    }

    [Fact]
    public void A_distinct_first_page_gets_the_background_too()
    {
        var client = new OfficeAgentClient(new WordModule());

        // The cover case: without this the background starts on page two, which reads as a
        // bug on the one page most likely to be a cover.
        var withCover = ApplyDoc(client, Blank(),
            new HeaderFooterOp { DifferentFirstPage = true, Header = "Running head" });
        var painted = ApplyDoc(client, withCover, new BackgroundImageOp { Base64Bytes = Png });

        using var stream = new MemoryStream(painted);
        using var opened = WordprocessingDocument.Open(stream, isEditable: false);

        var withAnchor = opened.MainDocumentPart!.HeaderParts
            .Count(h => h.Header!.Descendants<DW.Anchor>().Any());
        Assert.Equal(2, withAnchor);
        AssertDocValid(painted);
    }

    [Fact]
    public void A_cover_can_carry_a_different_background_from_the_pages_behind_it()
    {
        var client = new OfficeAgentClient(new WordModule());

        // The reason scope exists: a full-strength cover, and a pale wash the body copy is
        // still readable over. One background for both cannot be either.
        var withCover = ApplyDoc(client, Blank(),
            new HeaderFooterOp { DifferentFirstPage = true, Header = "Running head" });

        var painted = ApplyDoc(client, withCover,
            new BackgroundImageOp { Scope = "firstPage", Base64Bytes = Png },
            new BackgroundImageOp { Scope = "default", Base64Bytes = Png, Opacity = 0.15 });

        using var stream = new MemoryStream(painted);
        using var opened = WordprocessingDocument.Open(stream, isEditable: false);
        var main = opened.MainDocumentPart!;
        var section = main.Document.Body!.GetFirstChild<SectionProperties>()!;

        var first = HeaderOf(main, section, HeaderFooterValues.First);
        var running = HeaderOf(main, section, HeaderFooterValues.Default);

        // The cover is full strength: no alpha written at all.
        Assert.Single(first.Descendants<DW.Anchor>());
        Assert.Empty(first.Descendants<A.AlphaModulationFixed>());

        // The body pages are faded, and the second write did not strip the first.
        Assert.Single(running.Descendants<DW.Anchor>());
        Assert.Equal(15000, running.Descendants<A.AlphaModulationFixed>().Single().Amount!.Value);
        AssertDocValid(painted);
    }

    [Fact]
    public void A_first_page_background_is_skipped_when_there_is_no_distinct_first_page()
    {
        var client = new OfficeAgentClient(new WordModule());

        // Without w:titlePg there is no first-page header, so writing one would create a
        // part Word never shows - which reads to the caller as nothing having happened.
        var painted = ApplyDoc(client, Blank(),
            new BackgroundImageOp { Scope = "firstPage", Base64Bytes = Png });

        using var stream = new MemoryStream(painted);
        using var opened = WordprocessingDocument.Open(stream, isEditable: false);

        Assert.Empty(opened.MainDocumentPart!.HeaderParts
            .SelectMany(h => h.Header!.Descendants<DW.Anchor>()));
        AssertDocValid(painted);
    }

    [Fact]
    public void A_scope_that_is_not_a_page_kind_is_refused()
    {
        var client = new OfficeAgentClient(new WordModule());

        var report = client.Preview(
            new StreamHandle(new MemoryStream(Blank())),
            new DocumentPlan
            {
                Operations = new PlanOperation[]
                {
                    new BackgroundImageOp { Base64Bytes = Png, Scope = "coverPage" }
                }
            });

        Assert.Contains("all, firstPage, default, or evenPage",
            Assert.Single(report.Errors).Message);
    }

    [Fact]
    public void A_deck_refuses_the_page_scope_and_says_what_to_use()
    {
        var client = new OfficeAgentClient(new PowerPointModule());

        var report = client.Preview(
            new StreamHandle(new MemoryStream(Deck())),
            new DocumentPlan
            {
                Format = OfficeAgent.Abstractions.DocumentFormat.PowerPoint,
                Operations = new PlanOperation[]
                {
                    new BackgroundImageOp { Base64Bytes = Png, Scope = "firstPage" }
                }
            });

        Assert.Contains("names the slides instead", Assert.Single(report.Errors).Message);
    }

    // ── helpers ───────────────────────────────────────────────────────────────

    private static Header HeaderOf(
        MainDocumentPart main, SectionProperties section, HeaderFooterValues kind)
    {
        var reference = section.Elements<HeaderReference>().Single(r => r.Type!.Value == kind);
        return ((HeaderPart)main.GetPartById(reference.Id!.Value!)).Header!;
    }

    private static byte[] Deck() => new PowerPointModule().CreateBlank();
    private static byte[] Blank() => new WordModule().CreateBlank();

    private static Slide SlideAt(byte[] deck, int index)
    {
        using var stream = new MemoryStream(deck);
        using var document = PresentationDocument.Open(stream, isEditable: false);
        var main = document.PresentationPart!;
        var entry = main.Presentation.SlideIdList!.Elements<SlideId>().ElementAt(index);
        return ((SlidePart)main.GetPartById(entry.RelationshipId!)).Slide;
    }

    private static byte[] ApplyDeck(OfficeAgentClient client, byte[] deck, params PlanOperation[] operations)
    {
        using var applied = client.Commit(
            new StreamHandle(new MemoryStream(deck)),
            new DocumentPlan
            {
                Format = OfficeAgent.Abstractions.DocumentFormat.PowerPoint,
                Operations = operations
            });
        Assert.True(applied.Committed,
            string.Join("; ", applied.Report.Errors.Select(e => $"{e.Code}: {e.Message}")));
        return applied.ToBytes();
    }

    private static byte[] ApplyDoc(OfficeAgentClient client, byte[] document, params PlanOperation[] operations)
    {
        using var applied = client.Commit(
            new StreamHandle(new MemoryStream(document)),
            new DocumentPlan { Operations = operations });
        Assert.True(applied.Committed,
            string.Join("; ", applied.Report.Errors.Select(e => $"{e.Code}: {e.Message}")));
        return applied.ToBytes();
    }

    private static void AssertDeckValid(byte[] deck)
    {
        using var stream = new MemoryStream(deck);
        using var opened = PresentationDocument.Open(stream, isEditable: false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(opened).ToList();
        Assert.True(errors.Count == 0,
            string.Join("; ", errors.Select(e => $"{e.Description} @ {e.Path?.XPath}")));
    }

    private static void AssertDocValid(byte[] document)
    {
        using var stream = new MemoryStream(document);
        using var opened = WordprocessingDocument.Open(stream, isEditable: false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(opened).ToList();
        Assert.True(errors.Count == 0,
            string.Join("; ", errors.Select(e => $"{e.Description} @ {e.Path?.XPath}")));
    }
}
