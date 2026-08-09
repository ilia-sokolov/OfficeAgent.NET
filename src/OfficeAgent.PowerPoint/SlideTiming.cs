using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using OfficeAgent.Abstractions;

namespace OfficeAgent.PowerPoint;

/// <summary>
/// Builds and extends a slide's animation timing tree.
/// </summary>
/// <remarks>
/// <c>p:timing</c> is a nested time-node tree rather than a list of effects. The shape
/// PowerPoint reads is four levels deep: a root node holds a <c>p:seq</c> "main sequence";
/// the sequence holds one node per <em>click</em>; each click holds one node per effect
/// <em>group</em>; each group holds the effects themselves. Which level an effect joins is
/// exactly what the trigger means - a new click, a new group after the previous one, or
/// another effect inside the current group running alongside it.
/// <para>
/// Every <c>p:cTn/@id</c> must be unique within the slide, so the whole tree is renumbered
/// after each insertion rather than tracking a high-water mark that a later removal would
/// invalidate.
/// </para>
/// </remarks>
internal static class SlideTiming
{
    /// <summary>An effect's preset identity and the filter that renders it.</summary>
    internal sealed class Effect
    {
        public Effect(int presetId, string? filter)
        {
            PresetId = presetId;
            Filter = filter;
        }

        public int PresetId { get; }

        /// <summary>Null for <c>appear</c>, which is a visibility flip with nothing to render.</summary>
        public string? Filter { get; }
    }

    /// <summary>
    /// The effects expressible as a filtered <c>p:animEffect</c>. Fly-in, zoom and the
    /// motion paths need interpolated properties rather than a filter, so they are absent
    /// rather than approximated by something that looks different from what was asked for.
    /// </summary>
    public static readonly IReadOnlyDictionary<string, Effect> Effects =
        new Dictionary<string, Effect>(StringComparer.OrdinalIgnoreCase)
        {
            ["appear"] = new(1, null),
            ["fade"] = new(10, "fade"),
            ["wipe"] = new(22, "wipe(up)"),
            ["blinds"] = new(3, "blinds(horizontal)"),
            ["checkerboard"] = new(5, "checkerboard(across)"),
            ["circle"] = new(6, "circle"),
            ["diamond"] = new(9, "diamond"),
            ["dissolve"] = new(11, "dissolve"),
            ["plus"] = new(13, "plus"),
            ["randomBar"] = new(15, "randomBar(horizontal)"),
            ["split"] = new(17, "barn(inVertical)"),
            ["wedge"] = new(19, "wedge"),
            ["wheel"] = new(21, "wheel(1)"),
            ["box"] = new(4, "box(in)")
        };

    public static string Names => string.Join(", ", Effects.Keys);

    /// <summary>Appends an effect for one shape, creating the timing scaffold if absent.</summary>
    public static void Add(
        SlideRef slide, uint shapeId, Effect effect, AnimationKind kind,
        AnimationTrigger trigger, int durationMs, int delayMs)
    {
        var sequence = MainSequence(slide);
        var clicks = sequence.CommonTimeNode!.ChildTimeNodeList!;

        var node = EffectNode(shapeId, effect, kind, trigger, durationMs, delayMs);

        // Where the effect joins is what the trigger means.
        var lastClick = clicks.Elements<ParallelTimeNode>().LastOrDefault();
        if (trigger == AnimationTrigger.OnClick || lastClick is null)
        {
            clicks.Append(ClickGroup(node));
        }
        else
        {
            var groups = lastClick.CommonTimeNode!.ChildTimeNodeList!;
            if (trigger == AnimationTrigger.AfterPrevious)
            {
                // Its own group inside the same click: it starts when the one before ends.
                groups.Append(EffectGroup(node));
            }
            else
            {
                // Alongside the previous effect, so into that group's own child list.
                var lastGroup = groups.Elements<ParallelTimeNode>().LastOrDefault();
                if (lastGroup is null) groups.Append(EffectGroup(node));
                else lastGroup.CommonTimeNode!.ChildTimeNodeList!.Append(node);
            }
        }

        Renumber(slide);
    }

    /// <summary>Removes every effect targeting one shape, and the scaffold if nothing is left.</summary>
    public static bool Remove(SlideRef slide, uint shapeId)
    {
        var timing = slide.Part.Slide.Timing;
        if (timing is null) return false;

        var removed = false;
        foreach (var node in timing.Descendants<ParallelTimeNode>().ToList())
        {
            var targets = node.Descendants<ShapeTarget>().Select(t => t.ShapeId?.Value);
            if (!targets.Contains(shapeId.ToString())) continue;
            // Only the innermost node owning this shape, so a sibling effect survives.
            if (node.Descendants<ParallelTimeNode>().Any()) continue;

            node.Remove();
            removed = true;
        }

        // Prune groups and clicks left holding nothing, then the tree itself.
        foreach (var node in timing.Descendants<ParallelTimeNode>().Reverse().ToList())
            if (node.CommonTimeNode?.ChildTimeNodeList is { } children && !children.HasChildren)
                node.Remove();

        var sequence = timing.Descendants<SequenceTimeNode>().FirstOrDefault();
        if (sequence?.CommonTimeNode?.ChildTimeNodeList is { HasChildren: false })
            timing.Remove();
        else
            Renumber(slide);

        return removed;
    }

    /// <summary>Whether any effect on the slide targets the shape.</summary>
    public static bool Animates(SlideRef slide, uint shapeId) =>
        slide.Part.Slide.Timing?.Descendants<ShapeTarget>()
            .Any(t => t.ShapeId?.Value == shapeId.ToString()) == true;

    /// <summary>The slide's main sequence, built along with the tree above it when missing.</summary>
    private static SequenceTimeNode MainSequence(SlideRef slide)
    {
        var slideElement = slide.Part.Slide;
        var existing = slideElement.Timing?.Descendants<SequenceTimeNode>().FirstOrDefault();
        if (existing is not null) return existing;

        var sequence = new SequenceTimeNode(
            new CommonTimeNode(new ChildTimeNodeList())
            {
                Id = 2U,
                Duration = "indefinite",
                NodeType = TimeNodeValues.MainSequence
            },
            new PreviousConditionList(SlideCondition("onPrev")),
            new NextConditionList(SlideCondition("onNext")))
        {
            Concurrent = true,
            NextAction = NextActionValues.Seek
        };

        var timing = new Timing(
            new TimeNodeList(
                new ParallelTimeNode(
                    new CommonTimeNode(new ChildTimeNodeList(sequence))
                    {
                        Id = 1U,
                        Duration = "indefinite",
                        Restart = TimeNodeRestartValues.Never,
                        NodeType = TimeNodeValues.TmingRoot
                    })));

        // p:timing is the last child of p:sld; appending keeps the sequence valid.
        slideElement.Append(timing);
        return sequence;
    }

    private static Condition SlideCondition(string trigger) =>
        new(new TargetElement(new SlideTarget())) { Event = new EnumValue<TriggerEventValues>(
            trigger == "onPrev" ? TriggerEventValues.OnPrevious : TriggerEventValues.OnNext),
            Delay = "0" };

    /// <summary>A click step: it waits for the click, then runs the groups inside it.</summary>
    private static ParallelTimeNode ClickGroup(ParallelTimeNode effect) =>
        new(new CommonTimeNode(
            new StartConditionList(new Condition { Delay = "indefinite" }),
            new ChildTimeNodeList(EffectGroup(effect)))
        {
            Fill = TimeNodeFillValues.Hold
        });

    /// <summary>A group within a click: everything inside runs together.</summary>
    private static ParallelTimeNode EffectGroup(ParallelTimeNode effect) =>
        new(new CommonTimeNode(
            new StartConditionList(new Condition { Delay = "0" }),
            new ChildTimeNodeList(effect))
        {
            Fill = TimeNodeFillValues.Hold
        });

    /// <summary>The effect itself: make the shape visible, then render the filter.</summary>
    private static ParallelTimeNode EffectNode(
        uint shapeId, Effect effect, AnimationKind kind,
        AnimationTrigger trigger, int durationMs, int delayMs)
    {
        var entrance = kind == AnimationKind.Entrance;

        var children = new ChildTimeNodeList(
            new SetBehavior(
                new CommonBehavior(
                    new CommonTimeNode(new StartConditionList(new Condition { Delay = "0" }))
                    {
                        Duration = "1",
                        Fill = TimeNodeFillValues.Hold
                    },
                    new TargetElement(new ShapeTarget { ShapeId = shapeId.ToString() }),
                    new AttributeNameList(new AttributeName("style.visibility"))),
                new ToVariantValue(new StringVariantValue { Val = entrance ? "visible" : "hidden" })));

        if (effect.Filter is { } filter)
            children.Append(new AnimateEffect(
                new CommonBehavior(
                    new CommonTimeNode { Duration = durationMs.ToString() },
                    new TargetElement(new ShapeTarget { ShapeId = shapeId.ToString() })))
            {
                Transition = entrance
                    ? AnimateEffectTransitionValues.In
                    : AnimateEffectTransitionValues.Out,
                Filter = filter
            });

        return new ParallelTimeNode(
            new CommonTimeNode(
                new StartConditionList(new Condition { Delay = delayMs.ToString() }),
                children)
            {
                PresetId = effect.PresetId,
                PresetClass = entrance ? TimeNodePresetClassValues.Entrance : TimeNodePresetClassValues.Exit,
                PresetSubtype = 0,
                Fill = TimeNodeFillValues.Hold,
                GroupId = 0U,
                NodeType = trigger switch
                {
                    AnimationTrigger.WithPrevious => TimeNodeValues.WithEffect,
                    AnimationTrigger.AfterPrevious => TimeNodeValues.AfterEffect,
                    _ => TimeNodeValues.ClickEffect
                }
            });
    }

    /// <summary>
    /// Renumbers every <c>p:cTn</c> in document order. The ids must be unique within the
    /// slide, and a high-water mark would drift the moment an effect was removed.
    /// </summary>
    private static void Renumber(SlideRef slide)
    {
        uint next = 1;
        foreach (var node in slide.Part.Slide.Timing?.Descendants<CommonTimeNode>()
                 ?? Enumerable.Empty<CommonTimeNode>())
            node.Id = next++;
    }
}
