using OfficeAgent.Abstractions;
using OfficeAgent.Core;

namespace OfficeAgent.PowerPoint;

/// <summary>
/// Animates a shape. Effects are appended to the slide's sequence in the order the
/// operations run, which is the order they play.
/// </summary>
internal sealed class SlideAnimateHandler : IOperationHandler
{
    private const int DefaultDurationMs = 500;

    /// <inheritdoc />
    public bool CanHandle(PlanOperation operation) =>
        operation is AnimateOp { Target: NodeAnchor { Kind: "shape" } };

    /// <inheritdoc />
    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (AnimateOp)operation;
        var anchor = (NodeAnchor)op.Target;

        var located = ShapeNodeProvider.Locate(anchor.Path, context.Package);
        if (located is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No shape with path '{anchor.Path}'. Shape paths come from inspect_document.nodes.",
                anchor));

        var removing = string.Equals(op.Effect, "none", StringComparison.OrdinalIgnoreCase);
        if (!removing && !SlideTiming.Effects.ContainsKey(op.Effect))
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"Unknown animation '{op.Effect}'. Expected none or one of: {SlideTiming.Names}. " +
                "Fly-in, zoom, grow and motion paths need interpolated properties rather than " +
                "a filter, so they are not available here.",
                anchor));

        if (op.DurationMs is <= 0 || op.DelayMs is < 0)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "durationMs must be positive and delayMs cannot be negative.", anchor));

        var shapeId = PowerPointModel.ShapeIdOf(located.Element);
        if (shapeId is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"The shape at '{anchor.Path}' has no id to animate.", anchor));

        if (removing && !SlideTiming.Animates(located.Slide, shapeId.Value))
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "That shape has no animation to remove.", anchor));

        var after = removing
            ? "animation removed"
            : $"{op.Kind.ToString().ToLowerInvariant()} {op.Effect.ToLowerInvariant()}, " +
              $"{op.Trigger}, {op.DurationMs ?? DefaultDurationMs}ms";

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "animate",
            Before = SlideTiming.Animates(located.Slide, shapeId.Value) ? "animated" : string.Empty,
            After = after,
            Context = $"slide {located.Slide.Number}",
            BlastRadius = 1
        });
    }

    /// <inheritdoc />
    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (AnimateOp)operation;
        var anchor = (NodeAnchor)op.Target;

        var located = ShapeNodeProvider.Locate(anchor.Path, context.Package)
            ?? throw new InvalidOperationException($"Shape '{anchor.Path}' vanished before apply.");
        var shapeId = PowerPointModel.ShapeIdOf(located.Element)!.Value;

        if (string.Equals(op.Effect, "none", StringComparison.OrdinalIgnoreCase))
        {
            SlideTiming.Remove(located.Slide, shapeId);
            return;
        }

        SlideTiming.Add(
            located.Slide,
            shapeId,
            SlideTiming.Effects[op.Effect],
            op.Kind,
            op.Trigger,
            op.DurationMs ?? DefaultDurationMs,
            op.DelayMs ?? 0);
    }
}
