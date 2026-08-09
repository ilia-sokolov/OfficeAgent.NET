using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.PowerPoint;
using DocFormat = OfficeAgent.Abstractions.DocumentFormat;

namespace OfficeAgent.Tests;

/// <summary>
/// Slide transitions and shape animations. A transition is one element; an animation is a
/// four-level time-node tree whose ids must be unique and whose nesting <em>is</em> the
/// trigger semantics, so these assert the structure PowerPoint reads rather than merely
/// that something was written.
/// </summary>
public class SlideAnimationTests
{
    private static OfficeAgentClient Client() => new(new PowerPointModule());

    // ── transitions ───────────────────────────────────────────────────────────

    [Fact]
    public void A_transition_lands_on_every_slide_when_untargeted()
    {
        var client = Client();
        var deck = Apply(client, ThreeSlides(client), new TransitionOp
        {
            Effect = "fade",
            DurationMs = 700
        });

        Assert.All(SlidesOf(deck), slide =>
        {
            var transition = slide.Transition;
            Assert.NotNull(transition);
            Assert.NotNull(transition!.GetFirstChild<FadeTransition>());
            Assert.Equal("700", transition.Duration?.Value);
            // The legacy bucket, for readers predating the 2010 attribute.
            Assert.Equal(TransitionSpeedValues.Medium, transition.Speed?.Value);
        });
        AssertValid(deck);
    }

    [Fact]
    public void A_directional_transition_carries_its_direction()
    {
        var client = Client();
        var deck = Apply(client, ThreeSlides(client), new TransitionOp
        {
            Effect = "push",
            Direction = "up"
        });

        var push = SlidesOf(deck)[0].Transition!.GetFirstChild<PushTransition>()!;
        Assert.Equal(TransitionSlideDirectionValues.Up, push.Direction?.Value);
        AssertValid(deck);
    }

    [Fact]
    public void A_self_running_deck_advances_without_a_click()
    {
        var client = Client();
        var deck = Apply(client, ThreeSlides(client), new TransitionOp
        {
            Effect = "wipe",
            AdvanceOnClick = false,
            AdvanceAfterMs = 5000
        });

        var transition = SlidesOf(deck)[0].Transition!;
        Assert.False(transition.AdvanceOnClick?.Value);
        Assert.Equal("5000", transition.AdvanceAfterTime?.Value);
        AssertValid(deck);
    }

    [Fact]
    public void The_transition_element_sits_where_the_schema_requires()
    {
        var client = Client();

        // p:transition must follow p:clrMapOvr and precede p:timing. Animating first is
        // what makes the ordering matter: appending would put it after p:timing.
        var animated = Apply(client, ThreeSlides(client), new AnimateOp
        {
            Target = ShapeOn(ThreeSlides(client), 1),
            Effect = "fade"
        });
        var both = Apply(client, animated, new TransitionOp { Effect = "fade" });

        var slide = SlidesOf(both)[1];
        var children = slide.ChildElements.Select(c => c.LocalName).ToList();
        Assert.True(children.IndexOf("transition") < children.IndexOf("timing"));
        AssertValid(both);
    }

    [Fact]
    public void Transition_none_removes_it_again()
    {
        var client = Client();
        var deck = Apply(client, ThreeSlides(client), new TransitionOp { Effect = "fade" });
        Assert.NotNull(SlidesOf(deck)[0].Transition);

        var cleared = Apply(client, deck, new TransitionOp { Effect = "none" });

        Assert.All(SlidesOf(cleared), slide => Assert.Null(slide.Transition));
        AssertValid(cleared);
    }

    [Fact]
    public void An_unknown_transition_lists_the_ones_that_exist()
    {
        var client = Client();

        var report = Preview(client, ThreeSlides(client), new TransitionOp { Effect = "morph" });

        var error = Assert.Single(report.Errors);
        Assert.Equal(ValidationErrorCodes.InvalidOperation, error.Code);
        Assert.Contains("dissolve", error.Message);
    }

    // ── animations ────────────────────────────────────────────────────────────

    [Fact]
    public void An_entrance_animation_builds_the_tree_powerpoint_reads()
    {
        var client = Client();
        var deck = ThreeSlides(client);
        var applied = Apply(client, deck, new AnimateOp
        {
            Target = ShapeOn(deck, 1),
            Effect = "fade",
            DurationMs = 750
        });

        var slide = SlidesOf(applied)[1];
        var timing = slide.Timing;
        Assert.NotNull(timing);

        // The main sequence is what PowerPoint walks; without it the effects never run.
        var sequence = timing!.Descendants<SequenceTimeNode>().Single();
        Assert.Equal(TimeNodeValues.MainSequence, sequence.CommonTimeNode!.NodeType?.Value);

        var effect = timing.Descendants<AnimateEffect>().Single();
        Assert.Equal("fade", effect.Filter?.Value);
        Assert.Equal(AnimateEffectTransitionValues.In, effect.Transition?.Value);
        Assert.Equal("750", effect.CommonBehavior!.CommonTimeNode!.Duration?.Value);

        // The shape is made visible before it fades in, or it stays hidden.
        Assert.Equal("visible",
            timing.Descendants<StringVariantValue>().Single().Val?.Value);
        Assert.Equal(ShapeIdOn(deck, 1).ToString(),
            timing.Descendants<ShapeTarget>().First().ShapeId?.Value);
        AssertValid(applied);
    }

    [Fact]
    public void Every_time_node_id_is_unique()
    {
        var client = Client();
        var deck = ThreeSlides(client);

        // Three effects means three insertions, each renumbering the whole tree.
        var applied = Apply(client, deck,
            new AnimateOp { Target = ShapeOn(deck, 1), Effect = "fade" },
            new AnimateOp { Target = ShapeOn(deck, 1), Effect = "wipe", Trigger = AnimationTrigger.WithPrevious },
            new AnimateOp { Target = ShapeOn(deck, 1), Effect = "circle", Trigger = AnimationTrigger.AfterPrevious });

        var ids = SlidesOf(applied)[1].Timing!.Descendants<CommonTimeNode>()
            .Select(n => n.Id?.Value)
            .ToList();

        // Duplicate ids are the classic way a hand-built timing tree opens as "repaired".
        Assert.Equal(ids.Count, ids.Distinct().Count());
        Assert.DoesNotContain(null, ids);
        AssertValid(applied);
    }

    [Fact]
    public void The_trigger_decides_where_in_the_tree_an_effect_lands()
    {
        var client = Client();
        var deck = ThreeSlides(client);

        var applied = Apply(client, deck,
            new AnimateOp { Target = ShapeOn(deck, 1), Effect = "fade" },
            new AnimateOp { Target = ShapeOn(deck, 1), Effect = "wipe", Trigger = AnimationTrigger.WithPrevious },
            new AnimateOp { Target = ShapeOn(deck, 1), Effect = "circle", Trigger = AnimationTrigger.OnClick });

        var sequence = SlidesOf(applied)[1].Timing!.Descendants<SequenceTimeNode>().Single();
        var clicks = sequence.CommonTimeNode!.ChildTimeNodeList!.Elements<ParallelTimeNode>().ToList();

        // Two clicks: the first carries fade and wipe together, the second carries circle.
        Assert.Equal(2, clicks.Count);
        Assert.Equal(2, clicks[0].Descendants<AnimateEffect>().Count());
        Assert.Single(clicks[1].Descendants<AnimateEffect>());

        var nodeTypes = SlidesOf(applied)[1].Timing!.Descendants<CommonTimeNode>()
            .Select(n => n.NodeType?.Value)
            .Where(t => t is not null)
            .ToList();
        Assert.Contains(TimeNodeValues.ClickEffect, nodeTypes);
        Assert.Contains(TimeNodeValues.WithEffect, nodeTypes);
        AssertValid(applied);
    }

    [Fact]
    public void An_exit_animation_hides_rather_than_shows()
    {
        var client = Client();
        var deck = ThreeSlides(client);
        var applied = Apply(client, deck, new AnimateOp
        {
            Target = ShapeOn(deck, 1),
            Effect = "fade",
            Kind = AnimationKind.Exit
        });

        var timing = SlidesOf(applied)[1].Timing!;
        Assert.Equal(AnimateEffectTransitionValues.Out,
            timing.Descendants<AnimateEffect>().Single().Transition?.Value);
        Assert.Equal("hidden", timing.Descendants<StringVariantValue>().Single().Val?.Value);
        Assert.Equal(TimeNodePresetClassValues.Exit,
            timing.Descendants<CommonTimeNode>().Single(n => n.PresetClass is not null).PresetClass?.Value);
        AssertValid(applied);
    }

    [Fact]
    public void Appear_needs_no_filter_at_all()
    {
        var client = Client();
        var deck = ThreeSlides(client);
        var applied = Apply(client, deck, new AnimateOp { Target = ShapeOn(deck, 1), Effect = "appear" });

        var timing = SlidesOf(applied)[1].Timing!;
        // Appear is a visibility flip; there is nothing to render, so no animEffect.
        Assert.Empty(timing.Descendants<AnimateEffect>());
        Assert.Single(timing.Descendants<SetBehavior>());
        AssertValid(applied);
    }

    [Fact]
    public void Removing_an_animation_takes_the_scaffold_with_it()
    {
        var client = Client();
        var deck = ThreeSlides(client);
        var animated = Apply(client, deck, new AnimateOp { Target = ShapeOn(deck, 1), Effect = "fade" });
        Assert.NotNull(SlidesOf(animated)[1].Timing);

        var cleared = Apply(client, animated, new AnimateOp { Target = ShapeOn(deck, 1), Effect = "none" });

        // An empty main sequence left behind is a timing tree that animates nothing.
        Assert.Null(SlidesOf(cleared)[1].Timing);
        AssertValid(cleared);
    }

    [Fact]
    public void Removing_one_shapes_animation_leaves_anothers()
    {
        var client = Client();
        var deck = ThreeSlides(client);
        var title = ShapeOn(deck, 1);
        var body = ShapeOn(deck, 1, second: true);

        var animated = Apply(client, deck,
            new AnimateOp { Target = title, Effect = "fade" },
            new AnimateOp { Target = body, Effect = "wipe" });
        Assert.Equal(2, SlidesOf(animated)[1].Timing!.Descendants<AnimateEffect>().Count());

        var cleared = Apply(client, animated, new AnimateOp { Target = title, Effect = "none" });

        var remaining = SlidesOf(cleared)[1].Timing!.Descendants<AnimateEffect>().Single();
        Assert.Equal("wipe(up)", remaining.Filter?.Value);
        AssertValid(cleared);
    }

    [Fact]
    public void An_effect_that_needs_a_motion_path_is_refused_by_name()
    {
        var client = Client();
        var deck = ThreeSlides(client);

        var report = Preview(client, deck, new AnimateOp { Target = ShapeOn(deck, 1), Effect = "flyIn" });

        var error = Assert.Single(report.Errors);
        Assert.Equal(ValidationErrorCodes.InvalidOperation, error.Code);
        // The message must say what IS available, not merely that this is not.
        Assert.Contains("fade", error.Message);
        Assert.Contains("motion paths", error.Message);
    }

    [Fact]
    public void Removing_an_animation_that_is_not_there_says_so()
    {
        var client = Client();
        var deck = ThreeSlides(client);

        var report = Preview(client, deck, new AnimateOp { Target = ShapeOn(deck, 1), Effect = "none" });

        Assert.Contains("no animation to remove", Assert.Single(report.Errors).Message);
    }

    // ── helpers ───────────────────────────────────────────────────────────────

    private static byte[] ThreeSlides(OfficeAgentClient client) =>
        Apply(client, new PowerPointModule().CreateBlank(),
            new InsertSlideOp
            {
                Slide = new SlideData
                {
                    Layout = "titleAndContent",
                    Title = "Second",
                    Body = new[] { "A point" }
                }
            },
            new InsertSlideOp { Slide = new SlideData { Title = "Third" } });

    /// <summary>The shape node for a slide's title, or its body when <paramref name="second"/>.</summary>
    private static NodeAnchor ShapeOn(byte[] deck, int slideIndex, bool second = false)
    {
        var slideId = SlideIdAt(deck, slideIndex);
        return new NodeAnchor { Kind = "shape", Path = $"shape#{slideId}/{(second ? 3U : 2U)}" };
    }

    private static uint ShapeIdOn(byte[] deck, int slideIndex) => 2U;

    private static uint SlideIdAt(byte[] deck, int index)
    {
        using var stream = new MemoryStream(deck);
        using var document = PresentationDocument.Open(stream, isEditable: false);
        return document.PresentationPart!.Presentation.SlideIdList!
            .Elements<SlideId>().ElementAt(index).Id!.Value;
    }

    private static List<Slide> SlidesOf(byte[] deck)
    {
        using var stream = new MemoryStream(deck);
        using var document = PresentationDocument.Open(stream, isEditable: false);
        var main = document.PresentationPart!;
        return main.Presentation.SlideIdList!.Elements<SlideId>()
            .Select(e => ((SlidePart)main.GetPartById(e.RelationshipId!)).Slide)
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
