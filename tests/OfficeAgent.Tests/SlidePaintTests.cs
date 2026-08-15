using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;
using DocFormat = OfficeAgent.Abstractions.DocumentFormat;

namespace OfficeAgent.Tests;

/// <summary>
/// Solid colour on a shape and behind a slide - the difference between text on white and
/// something that looks designed. Both are a DrawingML solid fill, but they live in
/// different places (<c>p:spPr</c> versus <c>p:bg</c>) and both sit inside sequences, so
/// these assert placement as well as presence.
/// </summary>
public class SlidePaintTests
{
    private static OfficeAgentClient Client() => new(new PowerPointModule());

    [Fact]
    public void A_slide_background_is_painted_before_the_shape_tree()
    {
        var client = Client();

        var deck = Apply(client, Deck(), new FormatOp
        {
            Target = new NodeAnchor { Kind = "slide", Path = "slide#256" },
            FillColor = "#1F3A5F"
        });

        var slide = SlideAt(deck, 0);
        var background = slide.CommonSlideData!.Background!;
        Assert.Equal("1F3A5F",
            background.Descendants<A.RgbColorModelHex>().Single().Val!.Value);

        // p:bg is the first child of p:cSld; appending would put it after the shape tree
        // and PowerPoint would refuse the file.
        Assert.Equal(0, slide.CommonSlideData.ChildElements.ToList().IndexOf(background));
        AssertValid(deck);
    }

    [Fact]
    public void A_background_can_be_cleared_again()
    {
        var client = Client();
        var painted = Apply(client, Deck(), new FormatOp
        {
            Target = new NodeAnchor { Kind = "slide", Path = "slide#256" },
            FillColor = "1F3A5F"
        });
        Assert.NotNull(SlideAt(painted, 0).CommonSlideData!.Background);

        var cleared = Apply(client, painted, new FormatOp
        {
            Target = new NodeAnchor { Kind = "slide", Path = "slide#256" },
            FillColor = "none"
        });

        Assert.Null(SlideAt(cleared, 0).CommonSlideData!.Background);
        AssertValid(cleared);
    }

    [Fact]
    public void A_shape_takes_a_fill_and_an_outline()
    {
        var client = Client();
        var withBox = Apply(client, Deck(), new InsertShapeOp
        {
            Target = new NodeAnchor { Kind = "slide", Path = "slide#256" },
            Text = new[] { "Accent" },
            XPx = 40, YPx = 400, WidthPx = 300, HeightPx = 60
        });

        var path = client.Inspect(withBox).Nodes
            .Single(n => n.Kind == "shape" && n.Summary.Contains("text box")).Path;

        var painted = Apply(client, withBox, new FormatOp
        {
            Target = new NodeAnchor { Kind = "shape", Path = path },
            FillColor = "E8B33A",
            LineColor = "1F3A5F",
            LineWidthPx = 2
        });

        var shape = SlideAt(painted, 0).Descendants<Shape>()
            .Single(s => s.Descendants<A.Text>().Any(t => t.Text == "Accent"));
        var properties = shape.ShapeProperties!;

        Assert.Equal("E8B33A",
            properties.GetFirstChild<A.SolidFill>()!.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);

        var outline = properties.GetFirstChild<A.Outline>()!;
        Assert.Equal("1F3A5F",
            outline.GetFirstChild<A.SolidFill>()!.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
        Assert.Equal(2 * 9525, outline.Width!.Value);

        // The fill follows the geometry and precedes the outline in the p:spPr sequence.
        var names = properties.ChildElements.Select(c => c.LocalName).ToList();
        Assert.True(names.IndexOf("prstGeom") < names.IndexOf("solidFill"), string.Join(", ", names));
        Assert.True(names.IndexOf("solidFill") < names.IndexOf("ln"), string.Join(", ", names));
        AssertValid(painted);
    }

    [Fact]
    public void Painting_a_shape_does_not_move_it_and_moving_does_not_repaint_it()
    {
        var client = Client();
        var withBox = Apply(client, Deck(), new InsertShapeOp
        {
            Target = new NodeAnchor { Kind = "slide", Path = "slide#256" },
            Text = new[] { "Card" },
            XPx = 40, YPx = 400, WidthPx = 300, HeightPx = 60
        });
        var path = client.Inspect(withBox).Nodes
            .Single(n => n.Kind == "shape" && n.Summary.Contains("text box")).Path;

        var painted = Apply(client, withBox, new FormatOp
        {
            Target = new NodeAnchor { Kind = "shape", Path = path }, FillColor = "222222"
        });
        var moved = Apply(client, painted, new FormatOp
        {
            Target = new NodeAnchor { Kind = "shape", Path = path }, XPx = 500
        });

        var shape = SlideAt(moved, 0).Descendants<Shape>()
            .Single(s => s.Descendants<A.Text>().Any(t => t.Text == "Card"));

        // Each operation touches only what it named: the paint survives the move, and the
        // paint-only operation left the original geometry alone.
        Assert.Equal("222222",
            shape.ShapeProperties!.GetFirstChild<A.SolidFill>()!.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
        Assert.Equal(500 * 9525L, shape.ShapeProperties.Transform2D!.Offset!.X!.Value);
        Assert.Equal(400 * 9525L, shape.ShapeProperties.Transform2D.Offset.Y!.Value);
        AssertValid(moved);
    }

    [Fact]
    public void Text_can_be_anchored_in_the_middle_of_a_shape()
    {
        var client = Client();
        var withBox = Apply(client, Deck(), new InsertShapeOp
        {
            Target = new NodeAnchor { Kind = "slide", Path = "slide#256" },
            Text = new[] { "133.4" },
            XPx = 40, YPx = 200, WidthPx = 400, HeightPx = 240
        });
        var path = client.Inspect(withBox).Nodes
            .Single(n => n.Kind == "shape" && n.Summary.Contains("text box")).Path;

        var anchored = Apply(client, withBox, new FormatOp
        {
            Target = new NodeAnchor { Kind = "shape", Path = path },
            VerticalAlignment = "middle"
        });

        var shape = SlideAt(anchored, 0).Descendants<Shape>()
            .Single(s => s.Descendants<A.Text>().Any(t => t.Text == "133.4"));
        var body = shape.TextBody!.GetFirstChild<A.BodyProperties>()!;

        Assert.Equal(A.TextAnchoringTypeValues.Center, body.Anchor!.Value);

        // a:bodyPr opens the text body; a paragraph before it invalidates the shape.
        Assert.Equal(0, shape.TextBody.ChildElements.ToList().IndexOf(body));
        AssertValid(anchored);
    }

    [Fact]
    public void An_anchor_that_is_not_a_position_is_refused()
    {
        var client = Client();
        var withBox = Apply(client, Deck(), new InsertShapeOp
        {
            Target = new NodeAnchor { Kind = "slide", Path = "slide#256" },
            Text = new[] { "Card" },
            XPx = 40, YPx = 400, WidthPx = 300, HeightPx = 60
        });
        var path = client.Inspect(withBox).Nodes
            .Single(n => n.Kind == "shape" && n.Summary.Contains("text box")).Path;

        var report = Preview(client, withBox, new FormatOp
        {
            Target = new NodeAnchor { Kind = "shape", Path = path },
            VerticalAlignment = "centre"
        });

        var error = Assert.Single(report.Errors);
        Assert.Equal(ValidationErrorCodes.InvalidOperation, error.Code);
        Assert.Contains("top, middle, or bottom", error.Message);
    }

    [Fact]
    public void Filling_a_placeholder_gives_it_a_shape_to_fill()
    {
        var client = Client();
        var deck = Apply(client, Deck(), new ChangeTextOp
        {
            Target = new TextSpanAnchor { ParaId = "slide256/shape2/p0", Expect = string.Empty },
            With = "Q2 Board Review",
            Mode = ChangeMode.Direct
        });

        var painted = Apply(client, deck, new FormatOp
        {
            Target = new NodeAnchor { Kind = "shape", Path = "shape#256/2" },
            FillColor = "F4F1ED"
        });

        var shape = SlideAt(painted, 0).Descendants<Shape>()
            .Single(s => s.Descendants<A.Text>().Any(t => t.Text == "Q2 Board Review"));
        var properties = shape.ShapeProperties!;

        // A placeholder takes its position from the layout and carries no geometry of its
        // own. Without one, PowerPoint has nothing to fill: the colour is written, survives
        // a round trip, and is never drawn - which reads as the operation having done
        // nothing at all.
        Assert.NotNull(properties.GetFirstChild<A.PresetGeometry>());
        Assert.Equal(A.ShapeTypeValues.Rectangle,
            properties.GetFirstChild<A.PresetGeometry>()!.Preset!.Value);

        var names = properties.ChildElements.Select(c => c.LocalName).ToList();
        Assert.True(names.IndexOf("prstGeom") < names.IndexOf("solidFill"), string.Join(", ", names));
        AssertValid(painted);
    }

    [Fact]
    public void Clearing_a_fill_does_not_invent_geometry()
    {
        var client = Client();
        var deck = Apply(client, Deck(), new ChangeTextOp
        {
            Target = new TextSpanAnchor { ParaId = "slide256/shape2/p0", Expect = string.Empty },
            With = "Untouched",
            Mode = ChangeMode.Direct
        });

        var cleared = Apply(client, deck, new FormatOp
        {
            Target = new NodeAnchor { Kind = "shape", Path = "shape#256/2" },
            FillColor = "none"
        });

        var shape = SlideAt(cleared, 0).Descendants<Shape>()
            .Single(s => s.Descendants<A.Text>().Any(t => t.Text == "Untouched"));

        // Nothing is being drawn, so the placeholder keeps inheriting from its layout
        // rather than being pinned to a rectangle it never asked for.
        Assert.Null(shape.ShapeProperties!.GetFirstChild<A.PresetGeometry>());
        AssertValid(cleared);
    }

    [Fact]
    public void A_slide_format_that_is_not_a_background_says_what_it_wanted()
    {
        var client = Client();

        var report = Preview(client, Deck(), new FormatOp
        {
            Target = new NodeAnchor { Kind = "slide", Path = "slide#256" },
            Bold = true
        });

        var error = Assert.Single(report.Errors);
        Assert.Equal(ValidationErrorCodes.InvalidOperation, error.Code);
        Assert.Contains("fillColor is required", error.Message);
    }

    [Fact]
    public void A_value_that_is_not_a_colour_is_refused()
    {
        var client = Client();

        foreach (var bad in new[] { "cornflower", "12345", "#GGGGGG" })
        {
            var report = Preview(client, Deck(), new FormatOp
            {
                Target = new NodeAnchor { Kind = "slide", Path = "slide#256" },
                FillColor = bad
            });
            Assert.Contains("six hex digits", Assert.Single(report.Errors).Message);
        }
    }

    // ── helpers ───────────────────────────────────────────────────────────────

    private static byte[] Deck() => new PowerPointModule().CreateBlank();

    private static Slide SlideAt(byte[] deck, int index)
    {
        using var stream = new MemoryStream(deck);
        using var document = PresentationDocument.Open(stream, isEditable: false);
        var main = document.PresentationPart!;
        var entry = main.Presentation.SlideIdList!.Elements<SlideId>().ElementAt(index);
        return ((SlidePart)main.GetPartById(entry.RelationshipId!)).Slide;
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
