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
/// Template population, style copying and clearing, and sections - the three things a deck
/// previously reported as unsupported. The slot vocabulary is the shape name, which is what
/// a template author actually controls and what PowerPoint's Selection Pane shows.
/// </summary>
public class SlideTemplateTests
{
    private static OfficeAgentClient Client() => new(new PowerPointModule());

    // ── fill ──────────────────────────────────────────────────────────────────

    [Fact]
    public void Named_shapes_are_surfaced_as_fillable_slots()
    {
        var client = Client();

        var slots = client.Inspect(TemplateDeck(client)).StructuralAnchors
            .Select(a => a.Tag)
            .ToList();

        // The blank deck's own title is "Title 1"; the template box is named by the author.
        Assert.Contains("ClientName", slots);
        Assert.All(client.Inspect(TemplateDeck(client)).StructuralAnchors,
            a => Assert.Equal("shapeName", a.Kind));
    }

    [Fact]
    public void A_slot_is_filled_by_name_without_knowing_where_it_sits()
    {
        var client = Client();

        var filled = Apply(client, TemplateDeck(client), new FillOp
        {
            Target = new StructuralAnchor { Tag = "ClientName" },
            Value = "Northwind Traders Limited"
        });

        Assert.Contains("Northwind Traders Limited", AllText(filled));
        Assert.DoesNotContain("[CLIENT]", AllText(filled));
        AssertValid(filled);
    }

    [Fact]
    public void Filling_replaces_the_slot_rather_than_appending_to_it()
    {
        var client = Client();

        var once = Apply(client, TemplateDeck(client), new FillOp
        {
            Target = new StructuralAnchor { Tag = "ClientName" },
            Value = "First"
        });
        var twice = Apply(client, once, new FillOp
        {
            Target = new StructuralAnchor { Tag = "ClientName" },
            Value = "Second"
        });

        // A slot holds the value it was given, not a running history of them.
        Assert.Contains("Second", AllText(twice));
        Assert.DoesNotContain("First", AllText(twice));
    }

    [Fact]
    public void A_name_used_on_two_slides_is_refused_until_it_is_qualified()
    {
        var client = Client();
        var deck = TwoSlideTemplate(client);

        var ambiguous = Preview(client, deck, new FillOp
        {
            Target = new StructuralAnchor { Tag = "Footer" },
            Value = "Confidential"
        });

        var error = Assert.Single(ambiguous.Errors);
        Assert.Equal(ValidationErrorCodes.AmbiguousAnchor, error.Code);
        Assert.Contains("slide256/Footer", error.Message);

        // Qualifying picks one, and leaves the other alone.
        var filled = Apply(client, deck, new FillOp
        {
            Target = new StructuralAnchor { Tag = "slide256/Footer" },
            Value = "Confidential"
        });

        var footers = client.Inspect(filled).Paragraphs
            .Where(p => p.Text is "Confidential" or "[FOOTER]")
            .Select(p => p.Text)
            .ToList();
        Assert.Equal(new[] { "Confidential", "[FOOTER]" }, footers);
    }

    [Fact]
    public void An_unknown_slot_names_where_slots_come_from()
    {
        var client = Client();

        var report = Preview(client, TemplateDeck(client), new FillOp
        {
            Target = new StructuralAnchor { Tag = "NoSuchSlot" },
            Value = "x"
        });

        var error = Assert.Single(report.Errors);
        Assert.Equal(ValidationErrorCodes.AnchorNotFound, error.Code);
        Assert.Contains("structuralAnchors", error.Message);
    }

    // ── copyStyles / clearStyles ──────────────────────────────────────────────

    [Fact]
    public void Formatting_is_copied_from_one_line_to_another()
    {
        var client = Client();
        var deck = BulletDeck(client);

        var styled = Apply(client, deck, new FormatOp
        {
            Target = new TextSpanAnchor { ParaId = Body(0), Expect = string.Empty },
            Bold = true,
            Color = "C00000"
        });

        var copied = Apply(client, styled, new CopyStylesOp
        {
            Target = new TextSpanAnchor { ParaId = Body(1), Expect = string.Empty },
            Source = new TextSpanAnchor { ParaId = Body(0), Expect = string.Empty }
        });

        var second = ParagraphAt(copied, 1);
        var properties = second.Elements<A.Run>().Single().RunProperties!;
        Assert.True(properties.Bold?.Value);
        Assert.Equal("C00000",
            properties.GetFirstChild<A.SolidFill>()!.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
        AssertValid(copied);
    }

    [Fact]
    public void Clearing_returns_a_line_to_the_layouts_own_styling()
    {
        var client = Client();
        var deck = BulletDeck(client);

        var styled = Apply(client, deck, new FormatOp
        {
            Target = new TextSpanAnchor { ParaId = Body(0), Expect = string.Empty },
            Bold = true,
            SizeHalfPoints = 40
        });
        Assert.True(ParagraphAt(styled, 0).Elements<A.Run>().Single().RunProperties!.Bold?.Value);

        var cleared = Apply(client, styled, new ClearStylesOp
        {
            Target = new TextSpanAnchor { ParaId = Body(0), Expect = string.Empty }
        });

        var run = ParagraphAt(cleared, 0).Elements<A.Run>().Single();
        Assert.Null(run.RunProperties?.Bold);
        Assert.Null(run.RunProperties?.FontSize);
        // The language tag is not formatting; dropping it makes the spell checker treat
        // the text as undetermined.
        Assert.Equal("en-US", run.RunProperties?.Language?.Value);
        Assert.Equal("Finish the migration", string.Concat(run.Descendants<A.Text>().Select(t => t.Text)));
        AssertValid(cleared);
    }

    [Fact]
    public void Clearing_a_span_leaves_the_rest_of_the_line_alone()
    {
        var client = Client();
        var deck = BulletDeck(client);

        var styled = Apply(client, deck, new FormatOp
        {
            Target = new TextSpanAnchor { ParaId = Body(0), Expect = string.Empty },
            Bold = true
        });

        var cleared = Apply(client, styled, new ClearStylesOp
        {
            Target = new TextSpanAnchor { ParaId = Body(0), Expect = "migration" },
            Scope = "run"
        });

        var runs = ParagraphAt(cleared, 0).Elements<A.Run>()
            .Select(r => (Text: string.Concat(r.Descendants<A.Text>().Select(t => t.Text)),
                          Bold: r.RunProperties?.Bold?.Value == true))
            .ToList();

        Assert.Contains(runs, r => r.Text == "migration" && !r.Bold);
        Assert.Contains(runs, r => r.Text.Contains("Finish") && r.Bold);
        AssertValid(cleared);
    }

    [Fact]
    public void An_unknown_scope_is_refused()
    {
        var client = Client();

        var report = Preview(client, BulletDeck(client), new ClearStylesOp
        {
            Target = new TextSpanAnchor { ParaId = Body(0), Expect = string.Empty },
            Scope = "everything"
        });

        Assert.Contains("'run', 'paragraph', or 'all'", Assert.Single(report.Errors).Message);
    }

    // ── sections ──────────────────────────────────────────────────────────────

    [Fact]
    public void A_section_groups_the_slides_from_where_it_starts()
    {
        var client = Client();
        var deck = ThreeSlides(client);

        var sectioned = Apply(client, deck, new SectionOp
        {
            Action = SectionAction.Add,
            Name = "Financials",
            Target = new NodeAnchor { Kind = "slide", Path = $"slide#{SlideIdAt(deck, 1)}" }
        });

        var sections = SectionsOf(sectioned);
        // The slides before it need a home too, or PowerPoint reports the file as damaged.
        Assert.Equal(new[] { "Default Section", "Financials" }, sections.Select(s => s.Name));
        Assert.Equal(new[] { SlideIdAt(deck, 0) }, sections[0].Ids);
        Assert.Equal(new[] { SlideIdAt(deck, 1), SlideIdAt(deck, 2) }, sections[1].Ids);
        AssertValid(sectioned);
    }

    [Fact]
    public void Sections_follow_the_deck_when_slides_move_and_arrive()
    {
        var client = Client();
        var deck = Apply(client, ThreeSlides(client), new SectionOp
        {
            Action = SectionAction.Add,
            Name = "Financials",
            Target = new NodeAnchor { Kind = "slide", Path = $"slide#{SlideIdAt(ThreeSlides(client), 1)}" }
        });

        // A slide added inside a section joins it, as it does in PowerPoint. A slide left
        // in no section at all is what makes PowerPoint offer to repair the deck.
        var grown = Apply(client, deck, new InsertSlideOp
        {
            Slide = new SlideData { Title = "Appendix" }
        });

        var sections = SectionsOf(grown);
        var everySectioned = sections.SelectMany(s => s.Ids).ToList();
        var everySlide = SlideIdsOf(grown);

        Assert.Equal(everySlide, everySectioned);
        Assert.Equal(4, everySlide.Count);
        AssertValid(grown);
    }

    [Fact]
    public void Sections_are_stored_in_the_order_powerpoint_shows_them()
    {
        var client = Client();
        var three = ThreeSlides(client);
        var deck = Apply(client, three, new SectionOp
        {
            Action = SectionAction.Add,
            Name = "Financials",
            Target = new NodeAnchor { Kind = "slide", Path = $"slide#{SlideIdAt(three, 1)}" }
        });

        // Moving the section's opening slide to the front puts that section first. PowerPoint
        // orders sections by where their slides are, so a stored order that disagreed would
        // make inspect_document list them differently from the thumbnail pane.
        var moved = Apply(client, deck, new MoveSlideOp
        {
            Target = new NodeAnchor { Kind = "slide", Path = $"slide#{SlideIdAt(three, 1)}" },
            Position = SlidePosition.Start
        });

        Assert.Equal(new[] { "Financials", "Default Section" }, SectionsOf(moved).Select(s => s.Name));
        Assert.Equal(SlideIdsOf(moved), SectionsOf(moved).SelectMany(s => s.Ids).ToList());
        AssertValid(moved);
    }

    [Fact]
    public void A_slide_moved_into_another_section_joins_it_rather_than_splitting_its_own()
    {
        var client = Client();
        var four = Apply(client, ThreeSlides(client),
            new InsertSlideOp { Slide = new SlideData { Title = "Fourth" } });

        // Two sections: [0] and [1,2,3].
        var deck = Apply(client, four, new SectionOp
        {
            Action = SectionAction.Add,
            Name = "Body",
            Target = new NodeAnchor { Kind = "slide", Path = $"slide#{SlideIdAt(four, 1)}" }
        });

        // Move the last slide to the front, into the first section's territory.
        var moved = Apply(client, deck, new MoveSlideOp
        {
            Target = new NodeAnchor { Kind = "slide", Path = $"slide#{SlideIdAt(four, 3)}" },
            Position = SlidePosition.Start
        });

        // A section owns a contiguous run. Letting the moved slide keep its old membership
        // would leave "Body" owning slides either side of "Default Section" - a shape
        // PresentationML does not allow and PowerPoint will not open.
        var order = SlideIdsOf(moved);
        var sections = SectionsOf(moved);
        Assert.Equal(order, sections.SelectMany(s => s.Ids).ToList());

        foreach (var section in sections.Where(s => s.Ids.Count > 0))
        {
            var positions = section.Ids.Select(id => order.IndexOf(id)).ToList();
            Assert.Equal(positions.Count, positions.Max() - positions.Min() + 1);
        }
        AssertValid(moved);
    }

    [Fact]
    public void A_removed_slide_leaves_no_dangling_section_entry()
    {
        var client = Client();
        var three = ThreeSlides(client);
        var deck = Apply(client, three, new SectionOp
        {
            Action = SectionAction.Add,
            Name = "Financials",
            Target = new NodeAnchor { Kind = "slide", Path = $"slide#{SlideIdAt(three, 1)}" }
        });

        var pruned = Apply(client, deck, new RemoveSlideOp
        {
            Target = new NodeAnchor { Kind = "slide", Path = $"slide#{SlideIdAt(three, 1)}" }
        });

        Assert.Equal(SlideIdsOf(pruned), SectionsOf(pruned).SelectMany(s => s.Ids).ToList());
        AssertValid(pruned);
    }

    [Fact]
    public void A_section_is_renamed_and_removed_without_losing_its_slides()
    {
        var client = Client();
        var three = ThreeSlides(client);
        var deck = Apply(client, three, new SectionOp
        {
            Action = SectionAction.Add,
            Name = "Financials",
            Target = new NodeAnchor { Kind = "slide", Path = $"slide#{SlideIdAt(three, 1)}" }
        });

        var path = client.Inspect(deck).Nodes
            .Single(n => n.Kind == "section" && n.Summary.Contains("Financials")).Path;

        var renamed = Apply(client, deck, new SectionOp
        {
            Action = SectionAction.Rename,
            Name = "FY27 Financials",
            Target = new NodeAnchor { Kind = "section", Path = path }
        });
        Assert.Contains("FY27 Financials", SectionsOf(renamed).Select(s => s.Name));

        var removed = Apply(client, renamed, new SectionOp
        {
            Action = SectionAction.Remove,
            Target = new NodeAnchor { Kind = "section", Path = path }
        });

        // Deleting a grouping must not delete the things being grouped.
        Assert.Equal(3, SlideIdsOf(removed).Count);
        Assert.DoesNotContain("FY27 Financials", SectionsOf(removed).Select(s => s.Name));
        AssertValid(removed);
    }

    [Fact]
    public void Two_sections_cannot_start_at_one_slide()
    {
        var client = Client();
        var three = ThreeSlides(client);
        var anchor = new NodeAnchor { Kind = "slide", Path = $"slide#{SlideIdAt(three, 1)}" };

        var deck = Apply(client, three, new SectionOp
        {
            Action = SectionAction.Add, Name = "Financials", Target = anchor
        });

        var report = Preview(client, deck, new SectionOp
        {
            Action = SectionAction.Add, Name = "Also here", Target = anchor
        });

        // The second would own no slides, which PowerPoint shows as a grouping you cannot fill.
        Assert.Contains("already starts at", Assert.Single(report.Errors).Message);
    }

    // ── helpers ───────────────────────────────────────────────────────────────

    private static string Body(int index) => $"slide257/shape3/p{index}";

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

    private static byte[] ThreeSlides(OfficeAgentClient client) =>
        Apply(client, new PowerPointModule().CreateBlank(),
            new InsertSlideOp { Slide = new SlideData { Title = "Second" } },
            new InsertSlideOp { Slide = new SlideData { Title = "Third" } });

    /// <summary>A deck with an author-named slot, the way a real template carries one.</summary>
    private static byte[] TemplateDeck(OfficeAgentClient client) =>
        Named(Apply(client, new PowerPointModule().CreateBlank(), new InsertShapeOp
        {
            Target = new NodeAnchor { Kind = "slide", Path = "slide#256" },
            Text = new[] { "[CLIENT]" },
            XPx = 40, YPx = 400, WidthPx = 400, HeightPx = 60
        }), "[CLIENT]", "ClientName");

    private static byte[] TwoSlideTemplate(OfficeAgentClient client)
    {
        var deck = Apply(client, new PowerPointModule().CreateBlank(),
            new InsertShapeOp
            {
                Target = new NodeAnchor { Kind = "slide", Path = "slide#256" },
                Text = new[] { "[FOOTER]" }, XPx = 40, YPx = 620, WidthPx = 300, HeightPx = 40
            },
            new InsertSlideOp { Slide = new SlideData { Title = "Second" } });

        deck = Apply(client, deck, new InsertShapeOp
        {
            Target = new NodeAnchor { Kind = "slide", Path = "slide#257" },
            Text = new[] { "[FOOTER]" }, XPx = 40, YPx = 620, WidthPx = 300, HeightPx = 40
        });

        return Named(deck, "[FOOTER]", "Footer");
    }

    /// <summary>Names the shape holding some text, as a template author would.</summary>
    private static byte[] Named(byte[] deck, string text, string name)
    {
        var buffer = new MemoryStream();
        buffer.Write(deck, 0, deck.Length);
        using (var document = PresentationDocument.Open(buffer, isEditable: true))
        {
            foreach (var part in document.PresentationPart!.SlideParts)
                foreach (var shape in part.Slide.Descendants<Shape>())
                    if (shape.Descendants<A.Text>().Any(t => t.Text == text))
                        shape.NonVisualShapeProperties!.NonVisualDrawingProperties!.Name = name;
        }
        return buffer.ToArray();
    }

    private static string AllText(byte[] deck)
    {
        using var stream = new MemoryStream(deck);
        using var document = PresentationDocument.Open(stream, isEditable: false);
        return string.Concat(document.PresentationPart!.SlideParts
            .SelectMany(p => p.Slide.Descendants<A.Text>())
            .Select(t => t.Text));
    }

    private static A.Paragraph ParagraphAt(byte[] deck, int index)
    {
        using var stream = new MemoryStream(deck);
        using var document = PresentationDocument.Open(stream, isEditable: false);
        var body = document.PresentationPart!.SlideParts
            .SelectMany(p => p.Slide.Descendants<Shape>())
            .Single(s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value == 3U)
            .TextBody!;
        return body.Elements<A.Paragraph>().ElementAt(index);
    }

    private static List<uint> SlideIdsOf(byte[] deck)
    {
        using var stream = new MemoryStream(deck);
        using var document = PresentationDocument.Open(stream, isEditable: false);
        return document.PresentationPart!.Presentation.SlideIdList!
            .Elements<SlideId>().Select(e => e.Id!.Value).ToList();
    }

    private static uint SlideIdAt(byte[] deck, int index) => SlideIdsOf(deck)[index];

    private static List<(string Name, List<uint> Ids)> SectionsOf(byte[] deck)
    {
        using var stream = new MemoryStream(deck);
        using var document = PresentationDocument.Open(stream, isEditable: false);
        var list = document.PresentationPart!.Presentation
            .GetFirstChild<PresentationExtensionList>()?
            .Elements<PresentationExtension>()
            .Select(e => e.GetFirstChild<P14.SectionList>())
            .FirstOrDefault(s => s is not null);

        return list is null
            ? new List<(string, List<uint>)>()
            : list.Elements<P14.Section>()
                .Select(s => (
                    s.Name?.Value ?? string.Empty,
                    s.SectionSlideIdList?.Elements<P14.SectionSlideIdListEntry>()
                        .Select(e => e.Id!.Value).ToList() ?? new List<uint>()))
                .ToList();
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
