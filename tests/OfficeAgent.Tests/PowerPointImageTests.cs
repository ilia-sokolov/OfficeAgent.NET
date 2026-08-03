using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.PowerPoint;
using DocFormat = OfficeAgent.Abstractions.DocumentFormat;

namespace OfficeAgent.Tests;

/// <summary>
/// Picture verbs on a slide: an image lands as a real <c>p:pic</c> backed by its own
/// image part, surfaces as an addressable node, and can be removed again.
/// </summary>
public class PowerPointImageTests
{
    /// <summary>The smallest valid PNG: a single transparent pixel.</summary>
    private const string OnePixelPng =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";

    private static OfficeAgentClient Client() => new(new PowerPointModule());

    [Fact]
    public void Insert_image_adds_a_picture_and_its_part()
    {
        var client = Client();

        var applied = Apply(client, PptxFactory.Deck(), new InsertImageOp
        {
            Target = new NodeAnchor { Kind = "slide", Path = "slide#256" },
            Base64Bytes = OnePixelPng,
            ImageType = "png",
            WidthPx = 320,
            HeightPx = 240,
            AltText = "Revenue chart"
        });

        var inspection = client.Inspect(applied);
        var image = Assert.Single(inspection.Nodes, n => n.Kind == "image");
        Assert.Equal("image#256/4", image.Path);
        // Alt text is the accessibility contract, and it has to land in a:descr - the node
        // summary would read the same if it had been written to the shape name instead,
        // where no screen reader looks.
        Assert.Contains("Revenue chart", image.Summary);
        Assert.Equal("Revenue chart", AltTextOf(applied));

        AssertValid(applied);
        AssertPictureHasResolvableImagePart(applied);
    }

    /// <summary>The alt text actually stored on the picture's drawing properties.</summary>
    private static string? AltTextOf(byte[] deck)
    {
        using var stream = new MemoryStream(deck);
        using var document = PresentationDocument.Open(stream, isEditable: false);

        return document.PresentationPart!.SlideParts
            .SelectMany(p => p.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Picture>())
            .Select(p => p.NonVisualPictureProperties?.NonVisualDrawingProperties?.Description?.Value)
            .Single();
    }

    [Fact]
    public void Inserted_image_keeps_the_requested_display_size()
    {
        var client = Client();

        var applied = Apply(client, PptxFactory.Deck(), new InsertImageOp
        {
            Target = new NodeAnchor { Kind = "slide", Path = "slide#256" },
            Base64Bytes = OnePixelPng,
            WidthPx = 100,
            HeightPx = 50
        });

        using var stream = new MemoryStream(applied);
        using var document = PresentationDocument.Open(stream, isEditable: false);
        var extents = document.PresentationPart!.SlideParts
            .SelectMany(p => p.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Picture>())
            .Select(p => p.ShapeProperties!.Transform2D!.Extents!)
            .Single();

        // 96 DPI: one pixel is 9525 EMU.
        Assert.Equal(100 * 9525L, extents.Cx!.Value);
        Assert.Equal(50 * 9525L, extents.Cy!.Value);
    }

    [Fact]
    public void Remove_image_takes_the_picture_off_the_slide()
    {
        var client = Client();
        var withImage = Apply(client, PptxFactory.Deck(), new InsertImageOp
        {
            Target = new NodeAnchor { Kind = "slide", Path = "slide#256" },
            Base64Bytes = OnePixelPng
        });

        var removed = Apply(client, withImage, new RemoveImageOp
        {
            Target = new NodeAnchor { Kind = "image", Path = "image#256/4" }
        });

        var inspection = client.Inspect(removed);
        Assert.DoesNotContain(inspection.Nodes, n => n.Kind == "image");
        // The slide's own content is untouched.
        Assert.Contains(inspection.Paragraphs, p => p.Text == PptxFactory.TitleText);
        AssertValid(removed);
    }

    [Fact]
    public void Malformed_and_unsupported_images_are_refused_before_anything_is_written()
    {
        var client = Client();
        var slide = new NodeAnchor { Kind = "slide", Path = "slide#256" };

        var noBytes = Preview(client, new InsertImageOp { Target = slide });
        var badBase64 = Preview(client, new InsertImageOp { Target = slide, Base64Bytes = "not base64!!" });
        var badType = Preview(client, new InsertImageOp
        {
            Target = slide,
            Base64Bytes = OnePixelPng,
            ImageType = "svg"
        });
        var badSize = Preview(client, new InsertImageOp
        {
            Target = slide,
            Base64Bytes = OnePixelPng,
            WidthPx = 0
        });

        Assert.All(
            new[] { noBytes, badBase64, badType, badSize },
            report => Assert.Equal(
                ValidationErrorCodes.InvalidOperation, Assert.Single(report.Errors).Code));
    }

    [Fact]
    public void A_missing_slide_or_image_is_reported_rather_than_applied()
    {
        var client = Client();

        var noSlide = Preview(client, new InsertImageOp
        {
            Target = new NodeAnchor { Kind = "slide", Path = "slide#999" },
            Base64Bytes = OnePixelPng
        });
        var noImage = Preview(client, new RemoveImageOp
        {
            Target = new NodeAnchor { Kind = "image", Path = "image#256/999" }
        });

        Assert.Equal(ValidationErrorCodes.AnchorNotFound, Assert.Single(noSlide.Errors).Code);
        Assert.Equal(ValidationErrorCodes.AnchorNotFound, Assert.Single(noImage.Errors).Code);
    }

    // ── helpers ───────────────────────────────────────────────────────────────

    private static byte[] Apply(OfficeAgentClient client, byte[] deck, PlanOperation operation)
    {
        using var applied = client.Commit(new StreamHandle(new MemoryStream(deck)), Plan(operation));
        Assert.True(applied.Committed,
            string.Join("; ", applied.Report.Errors.Select(e => $"{e.Code}: {e.Message}")));
        return applied.ToBytes();
    }

    private static ChangeReport Preview(OfficeAgentClient client, PlanOperation operation) =>
        client.Preview(new StreamHandle(new MemoryStream(PptxFactory.Deck())), Plan(operation));

    private static DocumentPlan Plan(PlanOperation operation) => new()
    {
        Format = DocFormat.PowerPoint,
        Operations = new[] { operation }
    };

    private static void AssertValid(byte[] deck)
    {
        using var stream = new MemoryStream(deck);
        using var document = PresentationDocument.Open(stream, isEditable: false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(document).ToList();
        Assert.True(errors.Count == 0,
            string.Join("; ", errors.Select(e => $"{e.Path?.XPath}: {e.Description}")));
    }

    /// <summary>
    /// A blip that names a relationship the slide part does not actually have renders as
    /// a broken-image placeholder, which schema validation does not catch.
    /// </summary>
    private static void AssertPictureHasResolvableImagePart(byte[] deck)
    {
        using var stream = new MemoryStream(deck);
        using var document = PresentationDocument.Open(stream, isEditable: false);

        foreach (var part in document.PresentationPart!.SlideParts)
            foreach (var picture in part.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Picture>())
            {
                var embed = picture.BlipFill?.Blip?.Embed?.Value;
                Assert.False(string.IsNullOrEmpty(embed));
                Assert.IsType<ImagePart>(part.GetPartById(embed!));
            }
    }
}
