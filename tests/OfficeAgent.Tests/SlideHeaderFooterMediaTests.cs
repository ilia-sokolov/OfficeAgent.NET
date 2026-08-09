using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;
using DocFormat = OfficeAgent.Abstractions.DocumentFormat;

namespace OfficeAgent.Tests;

/// <summary>
/// The running items along a slide's edge, and embedded timeline media. Both are shapes
/// PowerPoint treats specially: a footer inherits its place from the layout, and a clip is
/// a picture wearing three relationships at once. Getting either half-right produces a file
/// that validates and then opens wrong, so these assert the structure PowerPoint reads.
/// </summary>
public class SlideHeaderFooterMediaTests
{
    /// <summary>A one-second silent MP4 is unnecessary: the bytes are opaque to the writer.</summary>
    private const string FakeMedia = "AAAAIGZ0eXBpc29tAAACAGlzb21pc28yYXZjMW1wNDE=";

    private const string OnePixelPng =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";

    private static OfficeAgentClient Client() => new(new PowerPointModule());

    // ── header / footer / slide number ────────────────────────────────────────

    [Fact]
    public void A_footer_and_slide_number_land_on_every_slide_when_untargeted()
    {
        var client = Client();
        var deck = Apply(client, ThreeSlides(client), new HeaderFooterOp
        {
            Footer = "Confidential — internal only",
            ShowSlideNumber = true
        });

        foreach (var slide in SlidesOf(deck))
        {
            Assert.Equal("Confidential — internal only", PlaceholderText(slide, PlaceholderValues.Footer));
            Assert.NotNull(Placeholder(slide, PlaceholderValues.SlideNumber));
        }
        AssertValid(deck);
    }

    [Fact]
    public void The_slide_number_is_a_field_so_powerpoint_renumbers_it()
    {
        var client = Client();
        var deck = Apply(client, ThreeSlides(client), new HeaderFooterOp { ShowSlideNumber = true });

        var shape = Placeholder(SlidesOf(deck)[1], PlaceholderValues.SlideNumber)!;
        var field = shape.TextBody!.Descendants<A.Field>().Single();

        // Written as literal text it would be wrong the moment a slide was inserted ahead.
        Assert.Equal("slidenum", field.Type?.Value);
        Assert.DoesNotContain(shape.TextBody.Descendants<A.Run>(), _ => true);
        AssertValid(deck);
    }

    [Fact]
    public void A_date_updates_automatically_unless_a_fixed_one_is_given()
    {
        var client = Client();

        var automatic = Apply(client, ThreeSlides(client), new HeaderFooterOp { ShowDateTime = true });
        var fixedDate = Apply(client, ThreeSlides(client), new HeaderFooterOp { DateTime = "4 January 2027" });

        var auto = Placeholder(SlidesOf(automatic)[0], PlaceholderValues.DateAndTime)!;
        Assert.Equal("datetime1", auto.TextBody!.Descendants<A.Field>().Single().Type?.Value);

        var pinned = Placeholder(SlidesOf(fixedDate)[0], PlaceholderValues.DateAndTime)!;
        Assert.Empty(pinned.TextBody!.Descendants<A.Field>());
        Assert.Equal("4 January 2027", string.Concat(pinned.Descendants<A.Text>().Select(t => t.Text)));
        AssertValid(fixedDate);
    }

    [Fact]
    public void The_layout_declares_the_placeholders_a_slide_inherits_from()
    {
        var client = Client();
        var deck = Apply(client, ThreeSlides(client), new HeaderFooterOp { Footer = "x" });

        using var stream = new MemoryStream(deck);
        using var document = PresentationDocument.Open(stream, isEditable: false);
        var master = document.PresentationPart!.SlideMasterParts.Single();

        // Without these the slide's footer has nothing to inherit geometry from and
        // PowerPoint stacks it in the top-left corner.
        foreach (var type in new[]
                 { PlaceholderValues.Footer, PlaceholderValues.SlideNumber, PlaceholderValues.DateAndTime })
        {
            Assert.Contains(master.SlideMaster!.CommonSlideData!.ShapeTree!.Elements<Shape>(),
                s => Type(s) == type);
            Assert.All(master.SlideLayoutParts, layout =>
                Assert.Contains(layout.SlideLayout!.CommonSlideData!.ShapeTree!.Elements<Shape>(),
                    s => Type(s) == type));
        }

        // The slide's own shape carries no geometry, so a template change restyles it.
        var slideFooter = Placeholder(SlidesOf(deck)[0], PlaceholderValues.Footer)!;
        Assert.Null(slideFooter.ShapeProperties?.Transform2D);
    }

    [Fact]
    public void Hiding_removes_the_placeholder_rather_than_blanking_it()
    {
        var client = Client();
        var shown = Apply(client, ThreeSlides(client), new HeaderFooterOp { Footer = "Draft" });
        Assert.NotNull(Placeholder(SlidesOf(shown)[0], PlaceholderValues.Footer));

        var hidden = Apply(client, shown, new HeaderFooterOp { ShowFooter = false });

        // An empty placeholder still shows PowerPoint's editing prompt, so a "hidden"
        // footer left in place would stay visible to whoever opened the deck.
        Assert.Null(Placeholder(SlidesOf(hidden)[0], PlaceholderValues.Footer));
        AssertValid(hidden);
    }

    [Fact]
    public void One_slide_can_differ_from_the_rest()
    {
        var client = Client();
        var deck = Apply(client, ThreeSlides(client), new HeaderFooterOp { Footer = "Everywhere" });

        var titleSlide = SlideIdAt(deck, 0);
        var amended = Apply(client, deck, new HeaderFooterOp
        {
            Target = new NodeAnchor { Kind = "slide", Path = $"slide#{titleSlide}" },
            ShowFooter = false
        });

        var slides = SlidesOf(amended);
        Assert.Null(Placeholder(slides[0], PlaceholderValues.Footer));
        Assert.Equal("Everywhere", PlaceholderText(slides[1], PlaceholderValues.Footer));
        AssertValid(amended);
    }

    [Fact]
    public void An_operation_that_changes_nothing_says_what_it_wanted()
    {
        var client = Client();

        var report = Preview(client, ThreeSlides(client), new HeaderFooterOp());

        var error = Assert.Single(report.Errors);
        Assert.Equal(ValidationErrorCodes.InvalidOperation, error.Code);
        Assert.Contains("showSlideNumber", error.Message);
    }

    // ── embedded media ────────────────────────────────────────────────────────

    [Fact]
    public void An_embedded_video_carries_the_three_relationships_powerpoint_reads()
    {
        var client = Client();
        var deck = Apply(client, new PowerPointModule().CreateBlank(), new InsertMediaOp
        {
            Target = new NodeAnchor { Kind = "slide", Path = "slide#256" },
            Kind = MediaKind.Video,
            Base64Bytes = FakeMedia,
            MediaType = "mp4",
            PosterBase64 = OnePixelPng,
            XPx = 100, YPx = 120, WidthPx = 480, HeightPx = 270,
            AltText = "Product walkthrough"
        });

        using var stream = new MemoryStream(deck);
        using var document = PresentationDocument.Open(stream, isEditable: false);
        var part = document.PresentationPart!.SlideParts.Single();
        var picture = part.Slide.Descendants<Picture>().Single();
        var nonVisual = picture.NonVisualPictureProperties!.ApplicationNonVisualDrawingProperties!;

        // a:videoFile is what PowerPoint plays; the p14:media extension is what makes it
        // embedded rather than a link to a file that is not there; the blip is the frame.
        Assert.NotNull(nonVisual.GetFirstChild<A.VideoFromFile>());
        var media = nonVisual.Descendants<P14.Media>().Single();
        Assert.False(string.IsNullOrEmpty(media.Embed?.Value));
        Assert.False(string.IsNullOrEmpty(picture.BlipFill?.Blip?.Embed?.Value));

        // Both point at one media part, not two copies of the bytes.
        Assert.Equal(
            part.DataPartReferenceRelationships.OfType<VideoReferenceRelationship>().Single().DataPart,
            part.DataPartReferenceRelationships.OfType<MediaReferenceRelationship>().Single().DataPart);
        Assert.Equal("Product walkthrough",
            picture.NonVisualPictureProperties.NonVisualDrawingProperties!.Description?.Value);
        Assert.Equal("ppaction://media",
            picture.Descendants<A.HyperlinkOnClick>().Single().Action?.Value);
        AssertValid(deck);
    }

    [Fact]
    public void Audio_uses_the_audio_element_not_the_video_one()
    {
        var client = Client();
        var deck = Apply(client, new PowerPointModule().CreateBlank(), new InsertMediaOp
        {
            Target = new NodeAnchor { Kind = "slide", Path = "slide#256" },
            Kind = MediaKind.Audio,
            Base64Bytes = FakeMedia,
            MediaType = "m4a"
        });

        using var stream = new MemoryStream(deck);
        using var document = PresentationDocument.Open(stream, isEditable: false);
        var nonVisual = document.PresentationPart!.SlideParts.Single()
            .Slide.Descendants<Picture>().Single()
            .NonVisualPictureProperties!.ApplicationNonVisualDrawingProperties!;

        Assert.NotNull(nonVisual.GetFirstChild<A.AudioFromFile>());
        Assert.Null(nonVisual.GetFirstChild<A.VideoFromFile>());
        AssertValid(deck);
    }

    [Fact]
    public void Media_appears_as_a_node_and_can_be_removed()
    {
        var client = Client();
        var deck = Apply(client, new PowerPointModule().CreateBlank(), new InsertMediaOp
        {
            Target = new NodeAnchor { Kind = "slide", Path = "slide#256" },
            Base64Bytes = FakeMedia,
            MediaType = "mp4"
        });

        var node = Assert.Single(client.Inspect(deck).Nodes.Where(n => n.Kind == "media"));
        Assert.Contains("embedded video", node.Summary);

        // Media shares the shape addressing space, so removal is the ordinary shape verb.
        var shapePath = "shape#" + node.Path.Substring("media#".Length);
        var pruned = Apply(client, deck, new RemoveShapeOp
        {
            Target = new NodeAnchor { Kind = "shape", Path = shapePath }
        });

        Assert.DoesNotContain(client.Inspect(pruned).Nodes, n => n.Kind == "media");
        AssertValid(pruned);
    }

    [Fact]
    public void A_kind_that_disagrees_with_the_file_type_is_refused()
    {
        var client = Client();

        // The element PowerPoint reads is chosen by the declared kind, not by the bytes,
        // so this would otherwise produce a deck that silently plays nothing.
        var report = Preview(client, new PowerPointModule().CreateBlank(), new InsertMediaOp
        {
            Target = new NodeAnchor { Kind = "slide", Path = "slide#256" },
            Kind = MediaKind.Audio,
            Base64Bytes = FakeMedia,
            MediaType = "mp4"
        });

        var error = Assert.Single(report.Errors);
        Assert.Equal(ValidationErrorCodes.InvalidOperation, error.Code);
        Assert.Contains("is video", error.Message);
    }

    [Fact]
    public void An_unknown_media_type_lists_the_ones_that_work()
    {
        var client = Client();

        var report = Preview(client, new PowerPointModule().CreateBlank(), new InsertMediaOp
        {
            Target = new NodeAnchor { Kind = "slide", Path = "slide#256" },
            Base64Bytes = FakeMedia,
            MediaType = "ogg"
        });

        Assert.Contains("mp4", Assert.Single(report.Errors).Message);
    }

    [Fact]
    public void Exactly_one_source_of_bytes_is_required()
    {
        var client = Client();
        var anchor = new NodeAnchor { Kind = "slide", Path = "slide#256" };

        var neither = Preview(client, new PowerPointModule().CreateBlank(),
            new InsertMediaOp { Target = anchor, MediaType = "mp4" });
        var both = Preview(client, new PowerPointModule().CreateBlank(),
            new InsertMediaOp
            {
                Target = anchor, MediaType = "mp4",
                Base64Bytes = FakeMedia, MediaDocumentId = "abc"
            });

        Assert.Contains("exactly one", Assert.Single(neither.Errors).Message);
        Assert.Contains("exactly one", Assert.Single(both.Errors).Message);
    }

    // ── helpers ───────────────────────────────────────────────────────────────

    private static PlaceholderValues? Type(Shape shape) =>
        shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?
            .PlaceholderShape?.Type?.Value;

    private static Shape? Placeholder(Slide slide, PlaceholderValues type) =>
        slide.CommonSlideData?.ShapeTree?.Elements<Shape>().FirstOrDefault(s => Type(s) == type);

    private static string PlaceholderText(Slide slide, PlaceholderValues type) =>
        string.Concat(Placeholder(slide, type)?.Descendants<A.Text>().Select(t => t.Text)
                      ?? Enumerable.Empty<string>());

    private static byte[] ThreeSlides(OfficeAgentClient client) =>
        Apply(client, new PowerPointModule().CreateBlank(),
            new InsertSlideOp { Slide = new SlideData { Title = "Second" } },
            new InsertSlideOp { Slide = new SlideData { Title = "Third" } });

    private static List<Slide> SlidesOf(byte[] deck)
    {
        using var stream = new MemoryStream(deck);
        using var document = PresentationDocument.Open(stream, isEditable: false);
        var main = document.PresentationPart!;
        return main.Presentation.SlideIdList!.Elements<SlideId>()
            .Select(e => ((SlidePart)main.GetPartById(e.RelationshipId!)).Slide)
            .ToList();
    }

    private static uint SlideIdAt(byte[] deck, int index)
    {
        using var stream = new MemoryStream(deck);
        using var document = PresentationDocument.Open(stream, isEditable: false);
        return document.PresentationPart!.Presentation.SlideIdList!
            .Elements<SlideId>().ElementAt(index).Id!.Value;
    }

    private static byte[] Apply(OfficeAgentClient client, byte[] deck, params PlanOperation[] operations)
    {
        using var applied = client.Commit(new StreamHandle(new MemoryStream(deck)), Plan(operations));
        Assert.True(applied.Committed,
            string.Join("; ", applied.Report.Errors.Select(e => $"{e.Code}: {e.Message}")));
        return applied.ToBytes();
    }

    private static ChangeReport Preview(OfficeAgentClient client, byte[] deck, params PlanOperation[] operations) =>
        client.Preview(new StreamHandle(new MemoryStream(deck)), Plan(operations));

    private static DocumentPlan Plan(PlanOperation[] operations) => new()
    {
        Format = DocFormat.PowerPoint,
        Operations = operations
    };

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
