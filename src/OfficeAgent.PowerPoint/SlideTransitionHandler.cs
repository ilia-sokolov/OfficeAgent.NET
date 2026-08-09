using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;

namespace OfficeAgent.PowerPoint;

/// <summary>
/// Sets the transition played when a slide arrives, and how the deck advances off it.
/// </summary>
/// <remarks>
/// <c>p:transition</c> sits between <c>p:clrMapOvr</c> and <c>p:timing</c> in
/// <c>p:sld</c>, and PresentationML enforces that order - appending it would produce a
/// file PowerPoint refuses. Duration is written as the legacy <c>speed</c> bucket as well
/// as the exact millisecond value, so the slide behaves sensibly in readers that predate
/// the 2010 attribute.
/// </remarks>
internal sealed class SlideTransitionHandler : IOperationHandler
{
    /// <summary>Effects taking a direction, and the ones that ignore one.</summary>
    private static readonly HashSet<string> Directional =
        new(StringComparer.OrdinalIgnoreCase) { "push", "wipe" };

    private static readonly string[] Known =
    {
        "none", "cut", "fade", "push", "wipe", "split", "dissolve", "checker", "blinds",
        "circle", "diamond", "plus", "wedge", "wheel", "randomBar", "zoom", "newsflash",
        "comb", "strips", "cover", "pull", "random"
    };

    /// <inheritdoc />
    public bool CanHandle(PlanOperation operation) => operation is TransitionOp;

    /// <inheritdoc />
    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (TransitionOp)operation;

        if (!Known.Contains(op.Effect, StringComparer.OrdinalIgnoreCase))
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"Unknown transition '{op.Effect}'. Expected one of: {string.Join(", ", Known)}.",
                op.Target));

        if (op.Direction is { Length: > 0 } direction &&
            !new[] { "up", "down", "left", "right" }.Contains(direction, StringComparer.OrdinalIgnoreCase))
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"Unknown direction '{direction}'. Expected up, down, left, or right.", op.Target));

        if (op.DurationMs is <= 0 || op.AdvanceAfterMs is < 0)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "durationMs must be positive and advanceAfterMs cannot be negative.", op.Target));

        var slides = Scope(context, op.Target, "transition", out var error);
        if (error is not null) return OperationPreview.Fail(error);

        var detail = op.Effect.ToLowerInvariant();
        if (op.Direction is { Length: > 0 } d && Directional.Contains(op.Effect)) detail += $" {d}";
        if (op.DurationMs is { } ms) detail += $", {ms}ms";
        if (op.AdvanceAfterMs is { } after) detail += $", auto-advance {after}ms";

        return OperationPreview.Ok(new ProposedChange
        {
            Target = op.Target,
            Verb = "transition",
            Before = string.Empty,
            After = detail,
            Context = slides.Count == 1 ? $"slide {slides[0].Number}" : $"{slides.Count} slides",
            BlastRadius = slides.Count
        });
    }

    /// <inheritdoc />
    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (TransitionOp)operation;

        foreach (var slide in Scope(context, op.Target, "transition", out _))
        {
            slide.Part.Slide.Transition?.Remove();
            if (string.Equals(op.Effect, "none", StringComparison.OrdinalIgnoreCase)) continue;

            var transition = new Transition();
            if (Effect(op) is { } element) transition.Append(element);

            if (op.DurationMs is { } ms)
            {
                transition.Duration = ms.ToString();
                // The legacy bucket, for anything reading the pre-2010 attribute.
                transition.Speed = ms <= 400 ? TransitionSpeedValues.Fast
                    : ms >= 1200 ? TransitionSpeedValues.Slow
                    : TransitionSpeedValues.Medium;
            }

            if (op.AdvanceOnClick is { } click) transition.AdvanceOnClick = click;
            if (op.AdvanceAfterMs is { } after) transition.AdvanceAfterTime = after.ToString();

            // p:transition must follow p:clrMapOvr and precede p:timing.
            var timing = slide.Part.Slide.Timing;
            if (timing is not null) slide.Part.Slide.InsertBefore(transition, timing);
            else slide.Part.Slide.Append(transition);
        }
    }

    /// <summary>The slides an untargeted or slide-targeted deck operation covers.</summary>
    internal static IReadOnlyList<SlideRef> Scope(
        ApplyContext context, Anchor? target, string verb, out ValidationError? error)
    {
        error = null;

        if (target is null)
            return PowerPointModel.Slides(context.Package).ToList();

        if (target is not NodeAnchor { Kind: "slide" } anchor)
        {
            error = new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"{verb} targets a slide - {{ \"kind\": \"slide\", \"path\": \"slide#256\" }} - " +
                "or no target at all to apply to every slide.",
                target);
            return Array.Empty<SlideRef>();
        }

        var slide = SlideList.Target(context, anchor);
        if (slide is null)
        {
            error = SlideList.NoSuchSlide(anchor);
            return Array.Empty<SlideRef>();
        }

        return new[] { slide };
    }

    /// <summary>The element naming the effect, with its direction where one applies.</summary>
    private static OpenXmlElement? Effect(TransitionOp op)
    {
        var direction = op.Direction?.ToLowerInvariant();

        return op.Effect.ToLowerInvariant() switch
        {
            "cut" => new CutTransition(),
            "fade" => new FadeTransition(),
            "dissolve" => new DissolveTransition(),
            "circle" => new CircleTransition(),
            "diamond" => new DiamondTransition(),
            "plus" => new PlusTransition(),
            "wedge" => new WedgeTransition(),
            "newsflash" => new NewsflashTransition(),
            "random" => new RandomTransition(),
            "zoom" => new ZoomTransition(),
            "wheel" => new WheelTransition(),
            "checker" => new CheckerTransition(),
            "blinds" => new BlindsTransition(),
            "randomBar" => new RandomBarTransition(),
            "split" => new SplitTransition(),
            "comb" => new CombTransition(),
            "strips" => new StripsTransition(),
            "push" => new PushTransition { Direction = Side(direction) },
            "wipe" => new WipeTransition { Direction = Side(direction) },
            // Cover and pull take eight directions in the schema, a union this SDK does
            // not surface as a typed enum, so they run in their default direction.
            "cover" => new CoverTransition(),
            "pull" => new PullTransition(),
            _ => null
        };
    }

    private static TransitionSlideDirectionValues Side(string? direction) => direction switch
    {
        "up" => TransitionSlideDirectionValues.Up,
        "down" => TransitionSlideDirectionValues.Down,
        "right" => TransitionSlideDirectionValues.Right,
        _ => TransitionSlideDirectionValues.Left
    };


}
