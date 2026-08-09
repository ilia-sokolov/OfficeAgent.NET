using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.PowerPoint;
using OfficeAgent.Word;
using A = DocumentFormat.OpenXml.Drawing;
using DocFormat = OfficeAgent.Abstractions.DocumentFormat;

namespace OfficeAgent.Tests;

/// <summary>
/// The slide-lifecycle verbs: a deck gains, loses, reorders and copies slides, and a whole
/// deck can be authored in one plan. The invariant running through all of it is that a
/// slide's id is durable - reordering must not move anyone's anchor, and a new slide must
/// not collide with an existing one.
/// </summary>
public class SlideLifecycleTests
{
    private static OfficeAgentClient Client() => new(new PowerPointModule());

    [Fact]
    public void A_deck_is_authored_from_nothing_in_a_single_plan()
    {
        var client = Client();

        // The whole point of the feature: one call turns a blank deck into a real one.
        var deck = Apply(client, new PowerPointModule().CreateBlank(),
            new InsertSlideOp
            {
                Slide = new SlideData
                {
                    Layout = "title",
                    Title = "FY27 Operating Plan",
                    Body = new[] { "Finance review - 12 March 2027" }
                }
            },
            new InsertSlideOp
            {
                Slide = new SlideData
                {
                    Layout = "titleAndContent",
                    Title = "Priorities",
                    Body = new[] { "Migrate billing", "Close the APAC gap", "Hold headcount flat" },
                    Notes = "Do not read the bullets out; talk to the second one."
                }
            },
            new InsertSlideOp
            {
                Slide = new SlideData { Layout = "sectionHeader", Title = "Financials" }
            });

        var slides = SlidesOf(deck);
        // Four: the blank deck's own starting slide, then the three authored ones.
        Assert.Equal(4, slides.Count);
        Assert.Equal(
            new[] { "", "FY27 Operating Plan", "Priorities", "Financials" },
            slides.Select(TitleOf));

        Assert.Contains("Hold headcount flat", TextOf(slides[2]));
        Assert.Equal(3, BodyParagraphCount(slides[2]));
        AssertValid(deck);
    }

    [Fact]
    public void An_authored_slide_inherits_its_geometry_from_the_layout()
    {
        var client = Client();
        var deck = Apply(client, new PowerPointModule().CreateBlank(),
            new InsertSlideOp
            {
                Slide = new SlideData { Title = "Inherited", Body = new[] { "One" } }
            });

        var slide = SlidesOf(deck)[1];

        // An empty p:spPr is what makes the layout own position and size. Writing a
        // transform here would pin the shape and stop a template change from restyling it.
        Assert.All(
            slide.CommonSlideData!.ShapeTree!.Elements<Shape>(),
            shape => Assert.Null(shape.ShapeProperties?.Transform2D));

        // ...and each shape must still name the placeholder it inherits from.
        Assert.All(
            slide.CommonSlideData.ShapeTree.Elements<Shape>(),
            shape => Assert.NotNull(shape.NonVisualShapeProperties?
                .ApplicationNonVisualDrawingProperties?.PlaceholderShape));
    }

    [Fact]
    public void Reordering_moves_the_slide_but_not_anybody_anchor()
    {
        var client = Client();
        var deck = ThreeSlideDeck(client);

        var before = client.Inspect(deck).Paragraphs
            .ToDictionary(p => p.ParaId, p => p.Text, StringComparer.Ordinal);
        var third = SlideIdAt(deck, 2);

        var moved = Apply(client, deck, new MoveSlideOp
        {
            Target = SlideAnchor(third),
            Position = SlidePosition.Start
        });

        Assert.Equal(third, SlideIdAt(moved, 0));

        // The anchors are keyed on slide id, so every one of them still resolves to the
        // same text. A positional scheme would have silently redirected all three.
        var after = client.Inspect(moved).Paragraphs
            .ToDictionary(p => p.ParaId, p => p.Text, StringComparer.Ordinal);
        Assert.Equal(before, after);
        AssertValid(moved);
    }

    [Fact]
    public void A_duplicate_is_independent_of_the_slide_it_came_from()
    {
        var client = Client();
        var deck = ThreeSlideDeck(client);
        var source = SlideIdAt(deck, 1);

        var withCopy = Apply(client, deck, new DuplicateSlideOp { Target = SlideAnchor(source) });

        var slides = SlidesOf(withCopy);
        Assert.Equal(4, slides.Count);
        // The copy lands right after the original, which is what PowerPoint does.
        Assert.Equal("Second", TitleOf(slides[1]));
        Assert.Equal("Second", TitleOf(slides[2]));

        var copyId = SlideIdAt(withCopy, 2);
        Assert.NotEqual(source, copyId);

        // Editing the copy must not touch the original.
        var edited = Apply(client, withCopy, new ChangeTextOp
        {
            Target = new TextSpanAnchor
            {
                ParaId = $"slide{copyId}/shape2/p0",
                Expect = "Second"
            },
            With = "Second (revised)",
            Mode = ChangeMode.Direct
        });

        var final = SlidesOf(edited);
        Assert.Equal("Second", TitleOf(final[1]));
        Assert.Equal("Second (revised)", TitleOf(final[2]));
        AssertValid(edited);
    }

    [Fact]
    public void Duplicating_a_slide_that_has_notes_gives_the_copy_its_own()
    {
        var client = Client();
        var deck = Apply(client, new PowerPointModule().CreateBlank(),
            new InsertSlideOp
            {
                Slide = new SlideData
                {
                    Title = "Delivery risks",
                    Body = new[] { "Migration is four weeks behind" },
                    Notes = "Do not commit to a recovery date."
                }
            });

        var source = SlideIdAt(deck, 1);
        var withCopy = Apply(client, deck, new DuplicateSlideOp { Target = SlideAnchor(source) });

        // A slide may hold only one notes part, so sharing the source's fails outright -
        // and had it succeeded, editing either slide's notes would have changed both.
        Assert.Equal(2, NotesPartCount(withCopy));

        var copyId = SlideIdAt(withCopy, 2);
        var edited = Apply(client, withCopy, new ChangeTextOp
        {
            Target = new TextSpanAnchor
            {
                ParaId = $"slide{copyId}/notes/shape2/p0",
                Expect = "Do not commit to a recovery date."
            },
            With = "Recovery plan due 17 August.",
            Mode = ChangeMode.Direct
        });

        var notes = client.Inspect(edited).Paragraphs
            .Where(p => p.Location == "notes")
            .Select(p => p.Text)
            .ToList();

        Assert.Contains("Do not commit to a recovery date.", notes);
        Assert.Contains("Recovery plan due 17 August.", notes);
        AssertValid(edited);
    }

    [Fact]
    public void A_removed_slide_takes_its_notes_with_it()
    {
        var client = Client();
        var deck = Apply(client, new PowerPointModule().CreateBlank(),
            new InsertSlideOp
            {
                Slide = new SlideData { Title = "Doomed", Notes = "Say nothing about this one." }
            });

        var doomed = SlideIdAt(deck, 1);
        Assert.Equal(2, NotesPartCount(deck) + 1);

        var pruned = Apply(client, deck, new RemoveSlideOp { Target = SlideAnchor(doomed) });

        Assert.Single(SlidesOf(pruned));
        // The notes part hung off the slide part; leaving it behind would orphan it.
        Assert.Equal(0, NotesPartCount(pruned));
        AssertValid(pruned);
    }

    [Fact]
    public void The_last_slide_cannot_be_removed()
    {
        var client = Client();
        var deck = new PowerPointModule().CreateBlank();
        var only = SlideIdAt(deck, 0);

        // PowerPoint cannot open a deck with no slides, so the refusal is what stops the
        // agent producing a file the user cannot recover.
        var report = Preview(client, deck, new RemoveSlideOp { Target = SlideAnchor(only) });

        var error = Assert.Single(report.Errors);
        Assert.Equal(ValidationErrorCodes.InvalidOperation, error.Code);
        Assert.Contains("no slides", error.Message);
    }

    [Fact]
    public void Positions_that_need_a_reference_slide_say_so()
    {
        var client = Client();
        var deck = ThreeSlideDeck(client);

        var missing = Preview(client, deck, new MoveSlideOp
        {
            Target = SlideAnchor(SlideIdAt(deck, 0)),
            Position = SlidePosition.After
        });
        var unknown = Preview(client, deck, new MoveSlideOp
        {
            Target = SlideAnchor(SlideIdAt(deck, 0)),
            Position = SlidePosition.After,
            RelativeTo = "slide#9999"
        });
        var itself = Preview(client, deck, new MoveSlideOp
        {
            Target = SlideAnchor(SlideIdAt(deck, 0)),
            Position = SlidePosition.After,
            RelativeTo = $"slide#{SlideIdAt(deck, 0)}"
        });

        Assert.Contains("relativeTo", Assert.Single(missing.Errors).Message);
        Assert.Equal(ValidationErrorCodes.AnchorNotFound, Assert.Single(unknown.Errors).Code);
        Assert.Contains("relative to itself", Assert.Single(itself.Errors).Message);
    }

    [Fact]
    public void Editing_a_slide_the_same_plan_removes_writes_nothing()
    {
        var client = Client();
        var deck = ThreeSlideDeck(client);
        var doomed = SlideIdAt(deck, 1);

        // Preview validates against the pre-apply document, so this passes validation and
        // only collides at apply time. What matters is that the deck is left untouched
        // rather than half-edited.
        using var applied = client.Commit(
            new StreamHandle(new MemoryStream(deck)),
            Plan(new PlanOperation[]
            {
                new RemoveSlideOp { Target = SlideAnchor(doomed) },
                new ChangeTextOp
                {
                    Target = new TextSpanAnchor { ParaId = $"slide{doomed}/shape2/p0", Expect = "Second" },
                    With = "Never written",
                    Mode = ChangeMode.Direct
                }
            }));

        Assert.False(applied.Committed);
        Assert.Contains(applied.Report.Errors, e => e.Code == ValidationErrorCodes.AnchorNotFound);
        // The original still has all three slides, the doomed one included.
        Assert.Equal(3, SlidesOf(deck).Count);
    }

    [Fact]
    public void An_unknown_layout_is_refused_with_the_ones_that_exist()
    {
        var client = Client();

        var report = Preview(client, new PowerPointModule().CreateBlank(), new InsertSlideOp
        {
            Slide = new SlideData { Layout = "twoContentWithPicture", Title = "Nope" }
        });

        var error = Assert.Single(report.Errors);
        Assert.Equal(ValidationErrorCodes.InvalidOperation, error.Code);
        Assert.Contains("titleAndContent", error.Message);
    }

    [Fact]
    public void A_word_document_refuses_the_slide_verbs()
    {
        var client = new OfficeAgentClient(new WordModule());

        // The vocabulary is shared, so these have to be named as unsupported rather than
        // silently skipped - a deck verb aimed at a .docx is an agent mistake worth seeing.
        foreach (var op in new PlanOperation[]
                 {
                     new InsertSlideOp { Slide = new SlideData { Title = "x" } },
                     new RemoveSlideOp { Target = SlideAnchor(256) },
                     new MoveSlideOp { Target = SlideAnchor(256) },
                     new DuplicateSlideOp { Target = SlideAnchor(256) }
                 })
        {
            var report = client.Preview(
                new StreamHandle(new MemoryStream(DocxFactory.Contract())),
                new DocumentPlan { Operations = new[] { op } });

            Assert.Equal(
                ValidationErrorCodes.UnsupportedOperation,
                Assert.Single(report.Errors).Code);
        }
    }

    // ── helpers ───────────────────────────────────────────────────────────────

    private static NodeAnchor SlideAnchor(uint slideId) =>
        new() { Kind = "slide", Path = $"slide#{slideId}" };

    /// <summary>A blank deck plus three titled slides, for the ordering tests.</summary>
    private static byte[] ThreeSlideDeck(OfficeAgentClient client)
    {
        var deck = Apply(client, new PowerPointModule().CreateBlank(),
            new InsertSlideOp { Slide = new SlideData { Title = "Second" } },
            new InsertSlideOp { Slide = new SlideData { Title = "Third" } });

        // Give the blank deck's starting slide a title too, so every slide is identifiable.
        return Apply(client, deck, new ChangeTextOp
        {
            Target = new TextSpanAnchor { ParaId = "slide256/shape2/p0", Expect = "" },
            With = "First",
            Mode = ChangeMode.Direct
        });
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

    private static List<Slide> SlidesOf(byte[] deck)
    {
        using var stream = new MemoryStream(deck);
        using var document = PresentationDocument.Open(stream, isEditable: false);
        return OrderedParts(document).Select(p => p.Slide).ToList();
    }

    private static uint SlideIdAt(byte[] deck, int index)
    {
        using var stream = new MemoryStream(deck);
        using var document = PresentationDocument.Open(stream, isEditable: false);
        return document.PresentationPart!.Presentation.SlideIdList!
            .Elements<SlideId>().ElementAt(index).Id!.Value;
    }

    private static int NotesPartCount(byte[] deck)
    {
        using var stream = new MemoryStream(deck);
        using var document = PresentationDocument.Open(stream, isEditable: false);
        return document.PresentationPart!.SlideParts.Count(p => p.NotesSlidePart is not null);
    }

    /// <summary>Slide parts in presentation order - the part list itself is unordered.</summary>
    private static IEnumerable<SlidePart> OrderedParts(PresentationDocument document)
    {
        var main = document.PresentationPart!;
        foreach (var entry in main.Presentation.SlideIdList!.Elements<SlideId>())
            if (main.GetPartById(entry.RelationshipId!) is SlidePart part)
                yield return part;
    }

    private static string TitleOf(Slide slide)
    {
        var title = slide.CommonSlideData?.ShapeTree?.Elements<Shape>().FirstOrDefault(s =>
        {
            var type = s.NonVisualShapeProperties?
                .ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value;
            return type == PlaceholderValues.Title || type == PlaceholderValues.CenteredTitle;
        });

        return title is null ? string.Empty : TextOfShape(title);
    }

    private static string TextOf(Slide slide) =>
        string.Concat(slide.Descendants<A.Text>().Select(t => t.Text));

    private static string TextOfShape(Shape shape) =>
        string.Concat(shape.Descendants<A.Text>().Select(t => t.Text));

    private static int BodyParagraphCount(Slide slide)
    {
        var body = slide.CommonSlideData?.ShapeTree?.Elements<Shape>().FirstOrDefault(s =>
        {
            var type = s.NonVisualShapeProperties?
                .ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value;
            return type != PlaceholderValues.Title && type != PlaceholderValues.CenteredTitle;
        });

        return body?.TextBody?.Elements<A.Paragraph>().Count() ?? 0;
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
