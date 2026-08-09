using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.PowerPoint;
using OfficeAgent.Word;
using A = DocumentFormat.OpenXml.Drawing;
using DocFormat = OfficeAgent.Abstractions.DocumentFormat;

namespace OfficeAgent.Tests;

/// <summary>
/// Adding a line to existing text, and the shape verbs around it: a text box appears, any
/// shape moves, resizes or goes away. The through-line is that a slide paragraph id is
/// positional, so inserting one renumbers the rest - and the module refuses a plan that
/// would then address that body by index rather than quietly editing the wrong line.
/// </summary>
public class SlideShapeTests
{
    private const long EmuPerPixel = 9525L;

    private static OfficeAgentClient Client() => new(new PowerPointModule());

    [Fact]
    public void A_bullet_is_added_after_an_existing_one()
    {
        var client = Client();
        var deck = BulletDeck(client);

        var applied = Apply(client, deck, new InsertOp
        {
            Target = new TextSpanAnchor { ParaId = BodyParagraph(1), Expect = "Rebuild the pipeline" },
            Position = InsertPosition.After,
            Text = "Hold headcount flat"
        });

        Assert.Equal(
            new[] { "Finish the migration", "Rebuild the pipeline", "Hold headcount flat" },
            BodyLines(client, applied));
        AssertValid(applied);
    }

    [Fact]
    public void An_inserted_bullet_keeps_the_styling_of_the_line_it_joins()
    {
        var client = Client();
        var deck = BulletDeck(client);

        // The neighbour is bolded first, so the insert has something to inherit that the
        // body default would not give it.
        var styled = Apply(client, deck, new FormatOp
        {
            Target = new TextSpanAnchor { ParaId = BodyParagraph(1), Expect = "Rebuild the pipeline" },
            Bold = true
        });

        var applied = Apply(client, styled, new InsertOp
        {
            Target = new TextSpanAnchor { ParaId = BodyParagraph(1), Expect = "Rebuild the pipeline" },
            Text = "Hold headcount flat",
            Level = 1
        });

        var inserted = ParagraphsOf(applied)
            .Single(p => string.Concat(p.Descendants<A.Text>().Select(t => t.Text)) == "Hold headcount flat");

        Assert.True(inserted.Elements<A.Run>().Single().RunProperties?.Bold?.Value);
        // level makes it a sub-bullet, which is the whole reason to state one.
        Assert.Equal(1, inserted.ParagraphProperties?.Level?.Value);
        AssertValid(applied);
    }

    [Fact]
    public void A_plan_that_inserts_and_then_addresses_the_same_body_is_refused()
    {
        var client = Client();
        var deck = BulletDeck(client);

        // p1 gains a line, so the old p1 becomes p2 - and this second operation still says
        // p1. Content verification would catch this one, but not an empty expect, so the
        // module refuses the shape of the plan rather than relying on luck.
        var report = Preview(client, deck,
            new InsertOp
            {
                Target = new TextSpanAnchor { ParaId = BodyParagraph(1), Expect = "Rebuild the pipeline" },
                Text = "Hold headcount flat"
            },
            new FormatOp
            {
                Target = new TextSpanAnchor { ParaId = BodyParagraph(1), Expect = string.Empty },
                Bold = true
            });

        var error = Assert.Single(report.Errors);
        Assert.Equal(ValidationErrorCodes.OperationConflict, error.Code);
        Assert.Contains("re-inspect", error.Message);
    }

    [Fact]
    public void An_earlier_paragraph_in_the_same_body_is_still_addressable()
    {
        var client = Client();
        var deck = BulletDeck(client);

        // Inserting at p1 cannot renumber p0, so this plan is safe and must be allowed -
        // refusing every co-occurrence would make the verb almost unusable.
        var applied = Apply(client, deck,
            new InsertOp
            {
                Target = new TextSpanAnchor { ParaId = BodyParagraph(1), Expect = "Rebuild the pipeline" },
                Text = "Hold headcount flat"
            },
            new FormatOp
            {
                Target = new TextSpanAnchor { ParaId = BodyParagraph(0), Expect = "Finish the migration" },
                Bold = true
            });

        Assert.Equal(
            new[] { "Finish the migration", "Rebuild the pipeline", "Hold headcount flat" },
            BodyLines(client, applied));

        // ...and the format landed on the line it named, not on the one that shifted.
        var bolded = ParagraphsOf(applied)
            .Single(p => p.Elements<A.Run>().Any(r => r.RunProperties?.Bold?.Value == true));
        Assert.Equal("Finish the migration", string.Concat(bolded.Descendants<A.Text>().Select(t => t.Text)));
        AssertValid(applied);
    }

    [Fact]
    public void A_text_box_is_added_and_can_be_removed_again()
    {
        var client = Client();
        var deck = Apply(client, new PowerPointModule().CreateBlank(), new InsertShapeOp
        {
            Target = new NodeAnchor { Kind = "slide", Path = "slide#256" },
            Text = new[] { "Draft - not for circulation" },
            XPx = 40,
            YPx = 500,
            WidthPx = 300,
            HeightPx = 60
        });

        var box = client.Inspect(deck).Nodes
            .Single(n => n.Kind == "shape" && n.Summary.Contains("text box"));
        Assert.Contains("Draft", box.Summary);

        var pruned = Apply(client, deck, new RemoveShapeOp
        {
            Target = new NodeAnchor { Kind = "shape", Path = box.Path }
        });

        Assert.DoesNotContain(client.Inspect(pruned).Nodes, n => n.Kind == "shape" && n.Summary.Contains("text box"));
        AssertValid(pruned);
    }

    [Fact]
    public void Any_shape_can_be_moved_and_resized()
    {
        var client = Client();
        var deck = Apply(client, new PowerPointModule().CreateBlank(), new InsertShapeOp
        {
            Target = new NodeAnchor { Kind = "slide", Path = "slide#256" },
            Text = new[] { "Footnote" },
            XPx = 10,
            YPx = 10,
            WidthPx = 100,
            HeightPx = 50
        });

        var path = client.Inspect(deck).Nodes.Single(n => n.Kind == "shape" && n.Summary.Contains("text box")).Path;

        var moved = Apply(client, deck, new FormatOp
        {
            Target = new NodeAnchor { Kind = "shape", Path = path },
            XPx = 200,
            YPx = 400,
            WidthPx = 640,
            HeightPx = 120
        });

        var located = BoxOf(moved, "Footnote");
        Assert.Equal(200 * EmuPerPixel, located.X);
        Assert.Equal(400 * EmuPerPixel, located.Y);
        Assert.Equal(640 * EmuPerPixel, located.Cx);
        Assert.Equal(120 * EmuPerPixel, located.Cy);
        AssertValid(moved);
    }

    [Fact]
    public void A_table_frame_moves_through_the_same_verb_as_a_text_box()
    {
        var client = Client();
        var withTable = Apply(client, new PowerPointModule().CreateBlank(), new InsertTableOp
        {
            Target = new NodeAnchor { Kind = "slide", Path = "slide#256" },
            Table = new TableData { Headers = new[] { "Region", "Revenue" } }
        });

        // A graphic frame keeps its transform in p:xfrm rather than p:spPr, so this is the
        // case that would break had the two been assumed identical.
        var path = client.Inspect(withTable).Nodes.Single(n => n.Kind == "shape" && n.Summary.Contains("table")).Path;

        var moved = Apply(client, withTable, new FormatOp
        {
            Target = new NodeAnchor { Kind = "shape", Path = path },
            XPx = 100,
            YPx = 150
        });

        using var stream = new MemoryStream(moved);
        using var document = PresentationDocument.Open(stream, isEditable: false);
        var frame = document.PresentationPart!.SlideParts.First()
            .Slide.Descendants<GraphicFrame>().Single();

        Assert.Equal(100 * EmuPerPixel, frame.Transform!.Offset!.X!.Value);
        Assert.Equal(150 * EmuPerPixel, frame.Transform.Offset.Y!.Value);
        AssertValid(moved);
    }

    [Fact]
    public void A_placeholder_is_not_removable()
    {
        var client = Client();
        var deck = new PowerPointModule().CreateBlank();

        var path = client.Inspect(deck).Nodes
            .Single(n => n.Kind == "shape" && n.Summary.Contains("placeholder")).Path;

        // Deleting it would leave the layout re-offering an empty prompt: the slide looks
        // unchanged while the content is gone.
        var report = Preview(client, deck, new RemoveShapeOp
        {
            Target = new NodeAnchor { Kind = "shape", Path = path }
        });

        var error = Assert.Single(report.Errors);
        Assert.Equal(ValidationErrorCodes.InvalidOperation, error.Code);
        Assert.Contains("changeText", error.Message);
    }

    [Fact]
    public void A_format_on_a_shape_with_no_geometry_says_what_it_wanted()
    {
        var client = Client();
        var deck = new PowerPointModule().CreateBlank();
        var path = client.Inspect(deck).Nodes.First(n => n.Kind == "shape").Path;

        var report = Preview(client, deck, new FormatOp
        {
            Target = new NodeAnchor { Kind = "shape", Path = path },
            Bold = true
        });

        var error = Assert.Single(report.Errors);
        Assert.Equal(ValidationErrorCodes.InvalidOperation, error.Code);
        Assert.Contains("target a paragraph instead", error.Message);
    }

    [Fact]
    public void A_null_paragraph_id_is_reported_against_the_anchor_that_is_wrong()
    {
        var client = Client();

        // An agent that failed to find a paragraph and passed the miss straight through
        // sends "paraId": null. Looking that up in the alias map raises "Value cannot be
        // null. (Parameter 'key')" - an error naming nothing the agent can act on.
        foreach (var op in new PlanOperation[]
                 {
                     new FormatOp { Target = new TextSpanAnchor { ParaId = null!, Expect = "" }, Bold = true },
                     new ChangeTextOp { Target = new TextSpanAnchor { ParaId = null!, Expect = "x" }, With = "y", Mode = ChangeMode.Direct },
                     new ClearStylesOp { Target = new TextSpanAnchor { ParaId = null!, Expect = "" } }
                 })
        {
            var report = Preview(client, BulletDeck(client), op);
            var error = Assert.Single(report.Errors);
            Assert.Equal(ValidationErrorCodes.AnchorNotFound, error.Code);
            Assert.DoesNotContain("Parameter 'key'", error.Message);
        }
    }

    [Fact]
    public void Word_refuses_a_bullet_level_rather_than_dropping_it()
    {
        var client = new OfficeAgentClient(new WordModule());

        var report = client.Preview(
            new StreamHandle(new MemoryStream(DocxFactory.Contract())),
            new DocumentPlan
            {
                Operations = new PlanOperation[]
                {
                    new InsertOp
                    {
                        Target = new TextSpanAnchor { ParaId = "w14:00000002", Expect = "Acme Corp" },
                        Text = "A new clause",
                        Level = 1
                    }
                }
            });

        var error = Assert.Single(report.Errors);
        Assert.Equal(ValidationErrorCodes.InvalidOperation, error.Code);
        Assert.Contains("styleId", error.Message);
    }

    [Fact]
    public void A_deck_refuses_a_style_id_on_insert()
    {
        var client = Client();
        var deck = BulletDeck(client);

        var report = Preview(client, deck, new InsertOp
        {
            Target = new TextSpanAnchor { ParaId = BodyParagraph(0), Expect = "Finish the migration" },
            Text = "Another",
            StyleId = "ListParagraph"
        });

        Assert.Contains("no paragraph style table", Assert.Single(report.Errors).Message);
    }

    // ── helpers ───────────────────────────────────────────────────────────────

    /// <summary>The body placeholder of the slide the bullet tests author.</summary>
    private static string BodyParagraph(int index) => $"slide257/shape3/p{index}";

    private static byte[] BulletDeck(OfficeAgentClient client) =>
        Apply(client, new PowerPointModule().CreateBlank(), new InsertSlideOp
        {
            Slide = new SlideData
            {
                Layout = "titleAndContent",
                Title = "FY27 Priorities",
                Body = new[] { "Finish the migration", "Rebuild the pipeline" }
            }
        });

    private static List<string> BodyLines(OfficeAgentClient client, byte[] deck) =>
        client.Inspect(deck).Paragraphs
            .Where(p => p.ParaId.StartsWith("slide257/shape3/p", StringComparison.Ordinal))
            .Select(p => p.Text)
            .ToList();

    private static IEnumerable<A.Paragraph> ParagraphsOf(byte[] deck)
    {
        using var stream = new MemoryStream(deck);
        using var document = PresentationDocument.Open(stream, isEditable: false);
        return document.PresentationPart!.SlideParts
            .SelectMany(p => p.Slide.Descendants<A.Paragraph>())
            .ToList();
    }

    private static (long X, long Y, long Cx, long Cy) BoxOf(byte[] deck, string text)
    {
        using var stream = new MemoryStream(deck);
        using var document = PresentationDocument.Open(stream, isEditable: false);
        var shape = document.PresentationPart!.SlideParts
            .SelectMany(p => p.Slide.Descendants<Shape>())
            .Single(s => s.Descendants<A.Text>().Any(t => t.Text == text));

        var transform = shape.ShapeProperties!.Transform2D!;
        return (transform.Offset!.X!.Value, transform.Offset.Y!.Value,
                transform.Extents!.Cx!.Value, transform.Extents.Cy!.Value);
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
