using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeAgent.Abstractions;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.PowerPoint;
using OfficeAgent.Word;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using P = DocumentFormat.OpenXml.Presentation;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using SVG = DocumentFormat.OpenXml.Office2019.Drawing.SVG;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.Tests;

/// <summary>
/// replaceImage: a template's picture takes a new image and keeps its look. The frame, its
/// position and effects stay; the crop is recomputed for the new image; the old bytes go only
/// when nothing else uses them; and a tracked replacement accepts and rejects in Word.
/// </summary>
public sealed class ReplaceImageTests
{
    // A 2:1 frame, as a banner slot in a template would be.
    private const long FrameWidth = 1905000;   // 200 px
    private const long FrameHeight = 952500;   // 100 px

    private static readonly byte[] OldImage = Convert.FromBase64String(NativeCorpus.Png(200, 100, 0xC0, 0x30, 0x30));
    private static readonly string Square = NativeCorpus.Png(120, 120, 0x2E, 0x75, 0xB6);
    private static readonly string Tall = NativeCorpus.Png(50, 200, 0x30, 0xA0, 0x40);

    // ── Crop arithmetic and image headers ───────────────────────────────

    [Theory]
    // A square image in a 2:1 frame: cover trims a quarter off the top and bottom; contain pads
    // half the image's width on each side.
    [InlineData(200, 100, 120, 120, ImageFit.Cover, 0, 25000, 0, 25000)]
    [InlineData(200, 100, 120, 120, ImageFit.Contain, -50000, 0, -50000, 0)]
    // A 1:4 image in a 2:1 frame: cover keeps the middle eighth; contain pads 3.5 widths a side.
    [InlineData(200, 100, 50, 200, ImageFit.Cover, 0, 43750, 0, 43750)]
    [InlineData(200, 100, 50, 200, ImageFit.Contain, -350000, 0, -350000, 0)]
    // A 4:1 image in a 1:1 frame: cover trims the sides, contain pads the top and bottom.
    [InlineData(100, 100, 400, 100, ImageFit.Cover, 37500, 0, 37500, 0)]
    [InlineData(100, 100, 400, 100, ImageFit.Contain, 0, -150000, 0, -150000)]
    public void The_crop_rectangle_meets_the_frame(long frameW, long frameH, int imageW, int imageH, ImageFit fit,
        int left, int top, int right, int bottom)
    {
        var crop = PictureFill.Crop(frameW, frameH, imageW, imageH, fit)!;

        Assert.Equal(left, crop.Left?.Value ?? 0);
        Assert.Equal(top, crop.Top?.Value ?? 0);
        Assert.Equal(right, crop.Right?.Value ?? 0);
        Assert.Equal(bottom, crop.Bottom?.Value ?? 0);

        // What remains of the image after the crop has the frame's proportions.
        var visibleW = imageW * (1 - (left + right) / 100000.0);
        var visibleH = imageH * (1 - (top + bottom) / 100000.0);
        Assert.Equal((double)frameW / frameH, visibleW / visibleH, 3);
    }

    [Fact]
    public void An_image_with_the_frames_proportions_needs_no_crop()
    {
        Assert.Null(PictureFill.Crop(FrameWidth, FrameHeight, 400, 200, ImageFit.Cover));
        Assert.Null(PictureFill.Crop(FrameWidth, FrameHeight, 400, 200, ImageFit.Contain));
    }

    [Fact]
    public void Image_headers_give_the_size_and_the_real_type()
    {
        Assert.True(PictureFill.TryMeasure(Convert.FromBase64String(Square), out var png, out var w, out var h));
        Assert.Equal(("image/png", 120, 120), (png, w, h));

        // JPEG: SOI, an APP0 segment, then SOF0 with height 0x0102 and width 0x0304.
        var jpeg = new byte[] { 0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x04, 0x00, 0x00, 0xFF, 0xC0, 0x00, 0x11, 0x08, 0x01, 0x02, 0x03, 0x04, 0x03, 0, 0, 0 };
        Assert.True(PictureFill.TryMeasure(jpeg, out var jpegType, out w, out h));
        Assert.Equal(("image/jpeg", 0x0304, 0x0102), (jpegType, w, h));

        var gif = new byte[] { 0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0x10, 0x00, 0x20, 0x00 };
        Assert.True(PictureFill.TryMeasure(gif, out var gifType, out w, out h));
        Assert.Equal(("image/gif", 16, 32), (gifType, w, h));

        var bmp = new byte[26];
        bmp[0] = 0x42; bmp[1] = 0x4D; bmp[14] = 40; bmp[18] = 7; bmp[22] = 0xF7; bmp[23] = 0xFF; bmp[24] = 0xFF; bmp[25] = 0xFF;
        Assert.True(PictureFill.TryMeasure(bmp, out var bmpType, out w, out h));
        Assert.Equal(("image/bmp", 7, 9), (bmpType, w, h));

        // Little-endian TIFF: one directory with width 30 (SHORT) and height 40 (LONG).
        var tiff = new byte[] { 0x49, 0x49, 0x2A, 0, 8, 0, 0, 0, 2, 0,
            0x00, 0x01, 3, 0, 1, 0, 0, 0, 30, 0, 0, 0,
            0x01, 0x01, 4, 0, 1, 0, 0, 0, 40, 0, 0, 0 };
        Assert.True(PictureFill.TryMeasure(tiff, out var tiffType, out w, out h));
        Assert.Equal(("image/tiff", 30, 40), (tiffType, w, h));

        Assert.False(PictureFill.TryMeasure(new byte[] { 1, 2, 3, 4, 5 }, out _, out _, out _));
    }

    // ── Word ─────────────────────────────────────────────────────────────

    [Theory]
    [InlineData(ImageFit.Cover, 0, 25000)]
    [InlineData(ImageFit.Contain, -50000, 0)]
    public void Word_an_inline_picture_keeps_its_frame_and_takes_the_new_image(ImageFit fit, int sides, int topAndBottom)
    {
        var input = WordDocument(Inline);
        var output = CommitWord(input, Replace("image#0", Square, fit, mode: ChangeMode.Direct));

        using var document = WordprocessingDocument.Open(new MemoryStream(output), false);
        var main = document.MainDocumentPart!;
        var drawing = Assert.Single(main.Document!.Body!.Descendants<W.Drawing>());
        Assert.Equal(FrameWidth, drawing.Inline!.Extent!.Cx!.Value);
        Assert.Equal(FrameHeight, drawing.Inline.Extent.Cy!.Value);
        var picture = drawing.Descendants<PIC.Picture>().Single();
        Assert.Equal(FrameWidth, picture.ShapeProperties!.Transform2D!.Extents!.Cx!.Value);

        var crop = picture.BlipFill!.SourceRectangle!;
        Assert.Equal(sides, crop.Left?.Value ?? 0);
        Assert.Equal(sides, crop.Right?.Value ?? 0);
        Assert.Equal(topAndBottom, crop.Top?.Value ?? 0);
        Assert.Equal(topAndBottom, crop.Bottom?.Value ?? 0);
        Assert.NotNull(picture.BlipFill.GetFirstChild<A.Stretch>()?.FillRectangle);

        // The new bytes are shown; the old ones are gone from the package; effects and alt text stay.
        Assert.Equal(Convert.FromBase64String(Square), Bytes(main, picture.BlipFill.Blip!.Embed!));
        Assert.Single(main.ImageParts);
        Assert.NotNull(picture.BlipFill.Blip.GetFirstChild<A.AlphaModulationFixed>());
        Assert.NotNull(picture.ShapeProperties.GetFirstChild<A.Outline>());
        Assert.Equal("Company banner", drawing.Inline.DocProperties!.Description!.Value);
        Assert.Equal(7U, drawing.Inline.DocProperties.Id!.Value);
        Assert.Empty(SchemaErrors(document));
    }

    [Fact]
    public void Word_a_floating_picture_keeps_its_position_wrap_z_order_and_hyperlink()
    {
        var input = WordDocument(Floating);
        string AnchorWithoutGraphic(byte[] bytes)
        {
            using var document = WordprocessingDocument.Open(new MemoryStream(bytes), false);
            var anchor = (DW.Anchor)document.MainDocumentPart!.Document!.Body!.Descendants<DW.Anchor>().Single().CloneNode(true);
            anchor.RemoveAllChildren<A.Graphic>();
            anchor.GetFirstChild<DW.DocProperties>()!.Description = null;
            return anchor.OuterXml;
        }

        var output = CommitWord(input, Replace("image#0", Tall, ImageFit.Cover, ChangeMode.Direct, description: "New product shot"));

        Assert.Equal(AnchorWithoutGraphic(input), AnchorWithoutGraphic(output));
        using var document = WordprocessingDocument.Open(new MemoryStream(output), false);
        var main = document.MainDocumentPart!;
        var anchor = main.Document!.Body!.Descendants<DW.Anchor>().Single();
        Assert.Equal("New product shot", anchor.GetFirstChild<DW.DocProperties>()!.Description!.Value);
        Assert.Equal("New product shot", anchor.Descendants<PIC.NonVisualDrawingProperties>().Single().Description!.Value);
        var link = anchor.GetFirstChild<DW.DocProperties>()!.GetFirstChild<A.HyperlinkOnClick>()!;
        Assert.Equal("https://example.com/", main.HyperlinkRelationships.Single(r => r.Id == link.Id!.Value).Uri.ToString());
        Assert.Equal(43750, anchor.Descendants<A.SourceRectangle>().Single().Top!.Value);
        Assert.Empty(SchemaErrors(document));
    }

    [Fact]
    public void Word_a_shared_image_part_stays_for_the_picture_still_using_it()
    {
        var input = WordDocument((main, body) =>
        {
            var id = AddImage(main, OldImage);
            body.Append(new W.Paragraph(new W.Run(InlineDrawing(id, 7))));
            body.Append(new W.Paragraph(new W.Run(InlineDrawing(id, 8))));
        });

        var output = CommitWord(input, Replace("image#1", Square, ImageFit.Cover, ChangeMode.Direct));

        using var document = WordprocessingDocument.Open(new MemoryStream(output), false);
        var main = document.MainDocumentPart!;
        var blips = main.Document!.Body!.Descendants<A.Blip>().ToList();
        Assert.Equal(OldImage, Bytes(main, blips[0].Embed!));
        Assert.Equal(Convert.FromBase64String(Square), Bytes(main, blips[1].Embed!));
        Assert.Equal(2, main.ImageParts.Count());
    }

    [Fact]
    public void Word_an_svg_picture_drops_its_vector_image_and_shows_the_new_raster()
    {
        var input = WordDocument((main, body) =>
        {
            var id = AddImage(main, OldImage);
            var svg = main.AddImagePart(ImagePartType.Svg);
            using (var stream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(
                "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"200\" height=\"100\"><rect width=\"200\" height=\"100\" fill=\"#c03030\"/></svg>")))
                svg.FeedData(stream);
            var drawing = InlineDrawing(id, 7);
            drawing.Descendants<A.Blip>().Single().Append(new A.BlipExtensionList(
                new A.BlipExtension(new SVG.SVGBlip { Embed = main.GetIdOfPart(svg) }) { Uri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}" }));
            body.Append(new W.Paragraph(new W.Run(drawing)));
        });

        var output = CommitWord(input, Replace("image#0", Square, ImageFit.Cover, ChangeMode.Direct));

        using var document = WordprocessingDocument.Open(new MemoryStream(output), false);
        var main = document.MainDocumentPart!;
        var blip = main.Document!.Body!.Descendants<A.Blip>().Single();
        Assert.Empty(blip.Descendants<SVG.SVGBlip>());
        Assert.Null(blip.GetFirstChild<A.BlipExtensionList>());
        Assert.Equal(Convert.FromBase64String(Square), Bytes(main, blip.Embed!));
        Assert.Single(main.ImageParts);
        Assert.DoesNotContain(main.ImageParts, part => part.ContentType == "image/svg+xml");
    }

    [Fact]
    public void Word_a_tracked_replacement_accepts_to_the_new_image_and_rejects_to_the_old()
    {
        var input = WordDocument(Inline);
        var client = new OfficeAgentClient(new WordModule());

        var output = CommitWord(input, Replace("image#0", Square, ImageFit.Cover, ChangeMode.Tracked));

        using (var document = WordprocessingDocument.Open(new MemoryStream(output), false))
        {
            var body = document.MainDocumentPart!.Document!.Body!;
            Assert.Single(body.Descendants<W.DeletedRun>().SelectMany(d => d.Descendants<W.Drawing>()));
            Assert.Single(body.Descendants<W.InsertedRun>().SelectMany(i => i.Descendants<W.Drawing>()));
            Assert.Equal(2, document.MainDocumentPart.ImageParts.Count());
            var ids = body.Descendants<DW.DocProperties>().Select(p => p.Id!.Value).ToList();
            Assert.Equal(ids.Count, ids.Distinct().Count());
            Assert.Empty(SchemaErrors(document));
        }

        var accepted = CommitWord(output, new RevisionOp { Target = new NodeAnchor { Kind = "revision", Path = "all" }, Action = RevisionAction.Accept });
        var rejected = CommitWord(output, new RevisionOp { Target = new NodeAnchor { Kind = "revision", Path = "all" }, Action = RevisionAction.Reject });

        Assert.Equal(Convert.FromBase64String(Square), OnlyPicture(accepted));
        Assert.Equal(OldImage, OnlyPicture(rejected));
        Assert.Single(client.Inspect(accepted).Nodes, n => n.Kind == "image");
    }

    [Fact]
    public void Word_tracked_replacement_strikes_only_the_picture_not_text_in_its_run()
    {
        var input = WordDocument((main, body) =>
        {
            var id = AddImage(main, OldImage);
            body.Append(new W.Paragraph(new W.Run(new W.Text("Logo: ") { Space = SpaceProcessingModeValues.Preserve }, InlineDrawing(id, 7), new W.Text(" end"))));
        });

        var output = CommitWord(input, Replace("image#0", Square, ImageFit.Cover, ChangeMode.Tracked));

        using var document = WordprocessingDocument.Open(new MemoryStream(output), false);
        var body = document.MainDocumentPart!.Document!.Body!;
        Assert.Empty(body.Descendants<W.DeletedText>());
        Assert.Equal("Quarterly reportLogo:  end", string.Concat(body.Descendants<W.Text>().Select(t => t.Text)));
        Assert.Single(body.Descendants<W.DeletedRun>());
        Assert.All(body.Descendants<W.Run>().Where(run => run.Descendants<W.Drawing>().Any()),
            run => Assert.Empty(run.Elements<W.Text>()));
        Assert.Empty(SchemaErrors(document));
    }

    [Fact]
    public void Word_refuses_what_it_cannot_replace()
    {
        var client = new OfficeAgentClient(new WordModule());
        var pictures = WordDocument(Inline);
        var shape = WordDocument((main, body) => body.Append(new W.Paragraph(new W.Run(ShapeDrawing()))));
        var deleted = CommitWord(pictures, new RemoveImageOp { Target = new NodeAnchor { Kind = "image", Path = "image#0" }, Mode = ChangeMode.Tracked });

        string Refusal(byte[] document, ReplaceImageOp op)
        {
            var report = client.Preview(new StreamHandle(new MemoryStream(document)), new DocumentPlan { Operations = new PlanOperation[] { op } });
            Assert.False(report.IsValid);
            return report.Errors.Single().Code;
        }

        Assert.Equal(ValidationErrorCodes.AnchorNotFound, Refusal(pictures, Replace("image#5", Square)));
        Assert.Equal(ValidationErrorCodes.InvalidOperation, Refusal(shape, Replace("image#0", Square)));
        Assert.Equal(ValidationErrorCodes.RevisionOverlap, Refusal(deleted, Replace("image#0", Square)));
        Assert.Equal(ValidationErrorCodes.InvalidOperation, Refusal(pictures, new ReplaceImageOp
        {
            Target = new NodeAnchor { Kind = "image", Path = "image#0" }, Base64Bytes = Square, ImageType = "jpeg"
        }));
        Assert.Equal(ValidationErrorCodes.InvalidOperation, Refusal(pictures, new ReplaceImageOp
        {
            Target = new NodeAnchor { Kind = "image", Path = "image#0" }, Base64Bytes = "not base64!"
        }));
        Assert.Equal(ValidationErrorCodes.InvalidOperation, Refusal(pictures, new ReplaceImageOp
        {
            Target = new NodeAnchor { Kind = "image", Path = "image#0" }, Base64Bytes = Square, ImageDocumentId = "x", ImageConnectionId = "c"
        }));
        Assert.Equal(ValidationErrorCodes.InvalidOperation, Refusal(pictures, new ReplaceImageOp
        {
            Target = new NodeAnchor { Kind = "image", Path = "image#0" }, Base64Bytes = Convert.ToBase64String(new byte[] { 1, 2, 3, 4, 5, 6 })
        }));
    }

    // ── PowerPoint ───────────────────────────────────────────────────────

    [Theory]
    [InlineData("image", ImageFit.Cover)]
    [InlineData("shape", ImageFit.Contain)]
    public void PowerPoint_a_picture_keeps_its_frame_and_takes_the_new_image(string kind, ImageFit fit)
    {
        var (input, path) = Deck(placeholder: false);
        var target = kind == "image" ? path : path.Replace("image#", "shape#", StringComparison.Ordinal);

        var output = CommitDeck(input, new ReplaceImageOp
        {
            Target = new NodeAnchor { Kind = kind, Path = target }, Base64Bytes = Square, ImageType = "png", Fit = fit
        });

        using var deck = PresentationDocument.Open(new MemoryStream(output), false);
        var slide = deck.PresentationPart!.SlideParts.Single(s => s.Slide!.Descendants<P.Picture>().Any());
        var picture = slide.Slide!.Descendants<P.Picture>().Single();
        var transform = picture.ShapeProperties!.Transform2D!;
        Assert.Equal((914400L, 1371600L, FrameWidth, FrameHeight),
            (transform.Offset!.X!.Value, transform.Offset.Y!.Value, transform.Extents!.Cx!.Value, transform.Extents.Cy!.Value));
        var crop = picture.BlipFill!.SourceRectangle!;
        if (fit == ImageFit.Cover) Assert.Equal((25000, 25000), (crop.Top!.Value, crop.Bottom!.Value));
        else Assert.Equal((-50000, -50000), (crop.Left!.Value, crop.Right!.Value));
        Assert.Equal(Convert.FromBase64String(Square), Bytes(slide, picture.BlipFill.Blip!.Embed!));
        Assert.Single(slide.ImageParts);
        Assert.Equal("Hero photo", picture.NonVisualPictureProperties!.NonVisualDrawingProperties!.Description!.Value);
        Assert.NotNull(picture.BlipFill.Blip.GetFirstChild<A.AlphaModulationFixed>());
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2019).Validate(deck));
    }

    [Fact]
    public void PowerPoint_a_picture_placeholder_is_fitted_to_the_frame_its_layout_gives_it()
    {
        var (input, path) = Deck(placeholder: true);

        var output = CommitDeck(input, new ReplaceImageOp
        {
            Target = new NodeAnchor { Kind = "image", Path = path }, Base64Bytes = Square, ImageType = "png", Description = "Team photo"
        });

        using var deck = PresentationDocument.Open(new MemoryStream(output), false);
        var picture = deck.PresentationPart!.SlideParts.SelectMany(s => s.Slide!.Descendants<P.Picture>()).Single();
        // The layout's 4:1 frame: a square image keeps the middle quarter of its height.
        Assert.Equal(37500, picture.BlipFill!.SourceRectangle!.Top!.Value);
        Assert.Null(picture.ShapeProperties!.Transform2D);
        Assert.Equal("Team photo", picture.NonVisualPictureProperties!.NonVisualDrawingProperties!.Description!.Value);
    }

    [Fact]
    public void PowerPoint_an_image_part_another_slide_uses_is_kept()
    {
        var (input, path) = Deck(placeholder: false, shareWithSecondSlide: true);

        var output = CommitDeck(input, Replace(path, Square));

        using var deck = PresentationDocument.Open(new MemoryStream(output), false);
        var slides = deck.PresentationPart!.SlideParts.Where(s => s.Slide!.Descendants<P.Picture>().Any()).ToList();
        Assert.Equal(2, slides.Count);
        var shown = slides.Select(s => Bytes(s, s.Slide!.Descendants<A.Blip>().Single().Embed!)).ToList();
        Assert.Contains(shown, bytes => bytes.SequenceEqual(OldImage));
        Assert.Contains(shown, bytes => bytes.SequenceEqual(Convert.FromBase64String(Square)));
    }

    [Fact]
    public void PowerPoint_an_svg_picture_drops_its_vector_image()
    {
        var (input, path) = Deck(placeholder: false, svg: true);

        var output = CommitDeck(input, Replace(path, Square));

        using var deck = PresentationDocument.Open(new MemoryStream(output), false);
        var slide = deck.PresentationPart!.SlideParts.Single(s => s.Slide!.Descendants<P.Picture>().Any());
        Assert.Empty(slide.Slide!.Descendants<SVG.SVGBlip>());
        Assert.Single(slide.ImageParts);
        Assert.Equal("image/png", slide.ImageParts.Single().ContentType);
    }

    [Fact]
    public void PowerPoint_refuses_tracked_mode_and_targets_that_hold_no_picture()
    {
        var (input, path) = Deck(placeholder: false, emptyPicturePlaceholder: true);
        var client = new OfficeAgentClient(new PowerPointModule());
        var shapes = client.Inspect(input).Nodes.Where(n => n.Kind == "shape").ToList();
        var empty = shapes.Single(n => n.Summary.Contains("placeholder 'Picture Placeholder'", StringComparison.Ordinal)).Path;
        var title = shapes.First(n => n.Summary.Contains("placeholder", StringComparison.Ordinal) && n.Path != empty).Path;

        ValidationError Refusal(ReplaceImageOp op)
        {
            var report = client.Preview(new StreamHandle(new MemoryStream(input), "d.pptx"), new DocumentPlan { Operations = new PlanOperation[] { op } });
            Assert.False(report.IsValid);
            return report.Errors.First();
        }

        var tracked = Refusal(new ReplaceImageOp { Target = new NodeAnchor { Kind = "image", Path = path }, Base64Bytes = Square, Mode = ChangeMode.Tracked });
        Assert.Equal(ValidationErrorCodes.InvalidOperation, tracked.Code);
        Assert.Contains("Tracked", tracked.Message, StringComparison.Ordinal);
        Assert.Contains("empty picture placeholder", Refusal(Replace(empty, Square, kind: "shape")).Message, StringComparison.Ordinal);
        Assert.Equal(ValidationErrorCodes.InvalidOperation, Refusal(Replace(title, Square, kind: "shape")).Code);
        Assert.Equal(ValidationErrorCodes.AnchorNotFound, Refusal(Replace("image#256/999", Square)).Code);
    }

    // ── Wiring ───────────────────────────────────────────────────────────

    [Fact]
    public void The_wire_shape_reads_with_fit_in_any_case()
    {
        var plan = JsonSerializer.Deserialize<DocumentPlan>("""
            { "operations": [ { "op": "replaceImage", "target": { "kind": "image", "path": "image#0" },
                                "base64Bytes": "iVBORw0KGgo=", "imageType": "png", "fit": "contain", "description": "Logo" } ] }
            """, OfficeAgentTools.PlanJson)!;

        var op = Assert.IsType<ReplaceImageOp>(Assert.Single(plan.Operations));
        Assert.Equal(ImageFit.Contain, op.Fit);
        Assert.Equal("Logo", op.Description);
        Assert.Equal(ImageFit.Cover, new ReplaceImageOp().Fit);
    }

    [Fact]
    public void Both_formats_advertise_replaceImage()
    {
        Assert.Contains("replaceImage", CapabilityDiscoveryTests.Describe(new WordModule()).Operations);
        Assert.Contains("replaceImage", CapabilityDiscoveryTests.Describe(new PowerPointModule()).Operations);
        Assert.DoesNotContain("replaceImage", CapabilityDiscoveryTests.Describe(new OfficeAgent.Excel.ExcelModule()).Operations);
        Assert.Contains("\"op\": \"replaceImage\"", OfficeAgentTools.PlanOperations, StringComparison.Ordinal);
        Assert.Contains("removeImage, replaceImage, backgroundImage", OfficeAgentTools.SystemPromptGuidance, StringComparison.Ordinal);
    }

    [Fact]
    public async Task An_image_held_by_a_connection_replaces_through_the_tools()
    {
        var root = Path.Combine(Path.GetTempPath(), $"officeagent-replace-{Guid.NewGuid():N}");
        Directory.CreateDirectory(root);
        try
        {
            var provider = new FileSystemDocumentProvider(new FileSystemDocumentProviderOptions
            {
                ConnectionId = "workspace", RootPath = root, DefaultChangeMode = ChangeMode.Direct,
                AllowedExtensions = new[] { ".docx", ".png" }
            });
            var client = new OfficeAgentClient(new DocumentProviderRegistry(new[] { provider }), new WordModule());
            File.WriteAllBytes(Path.Combine(root, "report.docx"), WordDocument(Inline));
            File.WriteAllBytes(Path.Combine(root, "logo.png"), Convert.FromBase64String(Square));
            var document = await client.RegisterAsync("workspace", Path.Combine(root, "report.docx"));
            var image = await client.RegisterAsync("workspace", Path.Combine(root, "logo.png"));
            var tools = new OfficeAgentTools(client);

            using var applied = JsonDocument.Parse(await tools.ApplyPlan("workspace", document.ItemId, $$"""
                [ { "op": "replaceImage", "target": { "kind": "image", "path": "image#0" },
                    "imageConnectionId": "workspace", "imageDocumentId": "{{image.ItemId}}", "imageType": "png" } ]
                """));

            Assert.True(applied.RootElement.GetProperty("committed").GetBoolean(), applied.RootElement.ToString());
            Assert.Equal(Convert.FromBase64String(Square), OnlyPicture(File.ReadAllBytes(Path.Combine(root, "report.docx"))));
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    // ── Fixtures ─────────────────────────────────────────────────────────

    internal static ReplaceImageOp Replace(string path, string base64, ImageFit fit = ImageFit.Cover,
        ChangeMode? mode = null, string? description = null, string kind = "image") => new()
    {
        Target = new NodeAnchor { Kind = kind, Path = path },
        Base64Bytes = base64,
        ImageType = "png",
        Fit = fit,
        Mode = mode,
        Description = description
    };

    private static byte[] CommitWord(byte[] input, PlanOperation op)
    {
        var client = new OfficeAgentClient(new WordModule());
        using var result = client.Commit(new StreamHandle(new MemoryStream(input), "input.docx"), new DocumentPlan { Operations = new[] { op } });
        Assert.True(result.Committed, string.Join("; ", result.Report.Errors.Select(e => $"{e.Code}: {e.Message}")));
        return result.ToBytes();
    }

    private static byte[] CommitDeck(byte[] input, PlanOperation op)
    {
        var client = new OfficeAgentClient(new PowerPointModule());
        using var result = client.Commit(new StreamHandle(new MemoryStream(input), "input.pptx"), new DocumentPlan { Operations = new[] { op } });
        Assert.True(result.Committed, string.Join("; ", result.Report.Errors.Select(e => $"{e.Code}: {e.Message}")));
        return result.ToBytes();
    }

    /// <summary>A Word document built on the engine's own blank, so it opens in Word as created.</summary>
    internal static byte[] WordDocument(Action<MainDocumentPart, W.Body> build)
    {
        using var stream = new MemoryStream();
        stream.Write(new WordModule().CreateBlank());
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var main = document.MainDocumentPart!;
            var body = main.Document!.Body!;
            body.InsertAt(new W.Paragraph(new W.Run(new W.Text("Quarterly report"))), 0);
            build(main, body);
            // The section properties close the body.
            if (body.GetFirstChild<W.SectionProperties>() is { } section) { section.Remove(); body.Append(section); }
        }
        return stream.ToArray();
    }

    internal static void Inline(MainDocumentPart main, W.Body body) =>
        body.Append(new W.Paragraph(new W.Run(InlineDrawing(AddImage(main, OldImage), 7))));

    internal static void Floating(MainDocumentPart main, W.Body body)
    {
        var id = AddImage(main, OldImage);
        var link = main.AddHyperlinkRelationship(new Uri("https://example.com/"), true).Id;
        var properties = new DW.DocProperties(new A.HyperlinkOnClick { Id = link }) { Id = 9U, Name = "Picture 9", Description = "Product shot" };
        body.Append(new W.Paragraph(new W.Run(new W.Text("Text that wraps around the product shot."))),
            new W.Paragraph(new W.Run(new W.Drawing(new DW.Anchor(
                new DW.SimplePosition { X = 0L, Y = 0L },
                new DW.HorizontalPosition(new DW.PositionOffset("1828800")) { RelativeFrom = DW.HorizontalRelativePositionValues.Column },
                new DW.VerticalPosition(new DW.PositionOffset("457200")) { RelativeFrom = DW.VerticalRelativePositionValues.Paragraph },
                new DW.Extent { Cx = FrameWidth, Cy = FrameHeight },
                new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                new DW.WrapSquare { WrapText = DW.WrapTextValues.BothSides },
                properties,
                new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
                Graphic(id, "Product shot"))
            {
                DistanceFromTop = 0U, DistanceFromBottom = 0U, DistanceFromLeft = 114300U, DistanceFromRight = 114300U,
                SimplePos = false, RelativeHeight = 251659264U, BehindDoc = false, Locked = false, LayoutInCell = true, AllowOverlap = true
            }))));
    }

    internal static string AddImage(OpenXmlPartContainer owner, byte[] bytes)
    {
        var part = owner switch
        {
            MainDocumentPart main => main.AddImagePart(ImagePartType.Png),
            SlidePart slide => slide.AddImagePart(ImagePartType.Png),
            _ => throw new ArgumentOutOfRangeException(nameof(owner))
        };
        using (var stream = new MemoryStream(bytes)) part.FeedData(stream);
        return owner.GetIdOfPart(part);
    }

    internal static W.Drawing InlineDrawing(string relationshipId, uint id) => new(new DW.Inline(
        new DW.Extent { Cx = FrameWidth, Cy = FrameHeight },
        new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
        new DW.DocProperties { Id = id, Name = $"Picture {id}", Description = "Company banner" },
        new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
        Graphic(relationshipId, "Company banner"))
    { DistanceFromTop = 0U, DistanceFromBottom = 0U, DistanceFromLeft = 0U, DistanceFromRight = 0U });

    private static A.Graphic Graphic(string relationshipId, string description) => new(new A.GraphicData(
        new PIC.Picture(
            new PIC.NonVisualPictureProperties(
                new PIC.NonVisualDrawingProperties { Id = 0U, Name = "Picture", Description = description },
                new PIC.NonVisualPictureDrawingProperties()),
            new PIC.BlipFill(
                new A.Blip(new A.AlphaModulationFixed { Amount = 90000 }) { Embed = relationshipId },
                new A.Stretch(new A.FillRectangle())),
            new PIC.ShapeProperties(
                new A.Transform2D(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = FrameWidth, Cy = FrameHeight }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle },
                new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = "1F3864" })) { Width = 19050 })))
    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" });

    /// <summary>A drawing holding a text box shape rather than a picture.</summary>
    private static W.Drawing ShapeDrawing() => new(new DW.Inline(
        new DW.Extent { Cx = FrameWidth, Cy = FrameHeight },
        new DW.DocProperties { Id = 3U, Name = "Shape 3" },
        new A.Graphic(new A.GraphicData(
            new DocumentFormat.OpenXml.Office2010.Word.DrawingShape.WordprocessingShape(
                new DocumentFormat.OpenXml.Office2010.Word.DrawingShape.NonVisualDrawingShapeProperties(),
                new DocumentFormat.OpenXml.Office2010.Word.DrawingShape.ShapeProperties(
                    new A.Transform2D(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = FrameWidth, Cy = FrameHeight }),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }),
                new DocumentFormat.OpenXml.Office2010.Word.DrawingShape.TextBodyProperties()))
        { Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape" })));

    /// <summary>
    /// A deck on the engine's blank with one picture on its first slide: a plain one at a fixed
    /// frame, or a picture placeholder that takes its frame from the layout.
    /// </summary>
    internal static (byte[] Deck, string Path) Deck(bool placeholder, bool shareWithSecondSlide = false, bool svg = false,
        bool emptyPicturePlaceholder = false)
    {
        var client = new OfficeAgentClient(new PowerPointModule());
        var start = new PowerPointModule().CreateBlank();
        if (shareWithSecondSlide)
        {
            using var added = client.Commit(new StreamHandle(new MemoryStream(start), "d.pptx"), new DocumentPlan
            {
                Operations = new PlanOperation[] { new InsertSlideOp { Slide = new SlideData { Layout = "blank" } } }
            });
            start = added.ToBytes();
        }

        using var stream = new MemoryStream();
        stream.Write(start);
        using (var deck = PresentationDocument.Open(stream, true))
        {
            var presentation = deck.PresentationPart!;
            var slides = presentation.Presentation!.SlideIdList!.Elements<P.SlideId>()
                .Select(id => (SlidePart)presentation.GetPartById(id.RelationshipId!)).ToList();
            var slide = slides[0];
            var tree = slide.Slide!.CommonSlideData!.ShapeTree!;
            var imageId = AddImage(slide, OldImage);

            P.Picture picture;
            if (placeholder)
            {
                // The layout offers a 4:1 picture placeholder; the slide's picture inherits its frame.
                var layoutTree = slide.SlideLayoutPart!.SlideLayout!.CommonSlideData!.ShapeTree!;
                layoutTree.Append(new P.Shape(
                    new P.NonVisualShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 40U, Name = "Picture Placeholder" },
                        new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                        new P.ApplicationNonVisualDrawingProperties(new P.PlaceholderShape { Type = P.PlaceholderValues.Picture, Index = 21U })),
                    new P.ShapeProperties(new A.Transform2D(new A.Offset { X = 914400L, Y = 2743200L }, new A.Extents { Cx = 7315200L, Cy = 1828800L }))));
                picture = Picture(50U, imageId, "Hero photo", ownFrame: false);
                picture.NonVisualPictureProperties!.ApplicationNonVisualDrawingProperties!.Append(
                    new P.PlaceholderShape { Type = P.PlaceholderValues.Picture, Index = 21U });
            }
            else
            {
                picture = Picture(50U, imageId, "Hero photo", ownFrame: true);
            }

            if (svg)
            {
                var vector = slide.AddImagePart(ImagePartType.Svg);
                using (var data = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(
                    "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"200\" height=\"100\"><rect width=\"200\" height=\"100\"/></svg>")))
                    vector.FeedData(data);
                picture.BlipFill!.Blip!.Append(new A.BlipExtensionList(
                    new A.BlipExtension(new SVG.SVGBlip { Embed = slide.GetIdOfPart(vector) }) { Uri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}" }));
            }
            tree.Append(picture);

            if (shareWithSecondSlide)
            {
                var second = slides[1];
                var shared = second.AddPart(slide.ImageParts.Single());
                second.Slide!.CommonSlideData!.ShapeTree!.Append(Picture(60U, second.GetIdOfPart(shared), "Shared photo", ownFrame: true));
            }

            if (emptyPicturePlaceholder)
                tree.Append(new P.Shape(
                    new P.NonVisualShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 70U, Name = "Picture Placeholder" },
                        new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                        new P.ApplicationNonVisualDrawingProperties(new P.PlaceholderShape { Type = P.PlaceholderValues.Picture, Index = 22U })),
                    new P.ShapeProperties(new A.Transform2D(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = 914400L, Cy = 914400L }))));
        }

        var bytes = stream.ToArray();
        var path = client.Inspect(bytes).Nodes.First(n => n.Kind == "image").Path;
        return (bytes, path);
    }

    private static P.Picture Picture(uint id, string relationshipId, string description, bool ownFrame) => new(
        new P.NonVisualPictureProperties(
            new P.NonVisualDrawingProperties { Id = id, Name = $"Picture {id}", Description = description },
            new P.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true }),
            new P.ApplicationNonVisualDrawingProperties()),
        new P.BlipFill(
            new A.Blip(new A.AlphaModulationFixed { Amount = 90000 }) { Embed = relationshipId },
            new A.Stretch(new A.FillRectangle())),
        ownFrame
            ? new P.ShapeProperties(
                new A.Transform2D(new A.Offset { X = 914400L, Y = 1371600L }, new A.Extents { Cx = FrameWidth, Cy = FrameHeight }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })
            : new P.ShapeProperties());

    private static byte[] Bytes(OpenXmlPartContainer owner, string relationshipId)
    {
        using var stream = ((ImagePart)owner.GetPartById(relationshipId)).GetStream();
        using var copy = new MemoryStream();
        stream.CopyTo(copy);
        return copy.ToArray();
    }

    /// <summary>The image bytes of a document's only picture.</summary>
    internal static byte[] OnlyPicture(byte[] docx)
    {
        using var document = WordprocessingDocument.Open(new MemoryStream(docx), false);
        var main = document.MainDocumentPart!;
        return Bytes(main, main.Document!.Body!.Descendants<A.Blip>().Single().Embed!);
    }

    private static string[] SchemaErrors(OpenXmlPackage package) =>
        new OpenXmlValidator(FileFormatVersions.Office2019).Validate(package)
            .Select(e => $"{e.Path?.XPath}: {e.Description}").ToArray();
}
