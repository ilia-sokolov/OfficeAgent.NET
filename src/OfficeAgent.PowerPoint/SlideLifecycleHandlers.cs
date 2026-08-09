using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeAgent.PowerPoint;

/// <summary>
/// Shared slide-list mechanics for the four slide-lifecycle verbs.
/// </summary>
/// <remarks>
/// Ordering lives entirely in <c>p:sldIdLst</c>: the slide parts themselves carry no
/// order, so reordering is a list operation and never rewrites a slide. That is also why
/// these verbs leave every existing anchor intact - a paragraph is addressed by its
/// slide's durable id, not by the slide's position.
/// </remarks>
internal static class SlideList
{
    /// <summary>PowerPoint requires a slide id in [256, 2147483647]; 256 is the first.</summary>
    private const uint MinimumSlideId = 256U;
    private const uint MaximumSlideId = 2147483647U;

    public static SlideIdList Require(IOpenXmlPackage package)
    {
        var presentation = PowerPointModel.Main(package).Presentation
            ?? throw new InvalidOperationException("Presentation part has no presentation.");
        return presentation.SlideIdList
            ??= new SlideIdList();
    }

    /// <summary>Allocates a slide id no other slide is using, inside the range PowerPoint accepts.</summary>
    public static uint NextSlideId(SlideIdList list)
    {
        uint highest = MinimumSlideId - 1;
        foreach (var entry in list.Elements<SlideId>())
            if (entry.Id?.Value is { } id && id > highest) highest = id;

        if (highest >= MaximumSlideId)
            throw new InvalidOperationException(
                "The deck has reached the highest slide id PresentationML allows.");

        return highest + 1;
    }

    /// <summary>Finds the <c>p:sldId</c> entry for one slide.</summary>
    public static SlideId? EntryFor(SlideIdList list, uint slideId) =>
        list.Elements<SlideId>().FirstOrDefault(e => e.Id?.Value == slideId);

    /// <summary>
    /// Places an entry per the requested position. The reference entry is only consulted
    /// for Before/After; Start and End address the list itself and need no target.
    /// </summary>
    public static void Place(SlideIdList list, SlideId entry, SlidePosition position, SlideId? reference)
    {
        switch (position)
        {
            case SlidePosition.Start:
                var first = list.Elements<SlideId>().FirstOrDefault();
                if (first is null) list.Append(entry); else list.InsertBefore(entry, first);
                break;

            case SlidePosition.Before when reference is not null:
                list.InsertBefore(entry, reference);
                break;

            case SlidePosition.After when reference is not null:
                list.InsertAfter(entry, reference);
                break;

            default:
                list.Append(entry);
                break;
        }
    }

    /// <summary>
    /// Resolves the reference slide a Before/After position needs, and reports what is
    /// wrong when it cannot. Start and End never need one, so they never fail here.
    /// </summary>
    public static ValidationError? ResolveReference(
        ApplyContext context,
        SlidePosition position,
        string? relativeTo,
        Anchor? target,
        string verb,
        out SlideRef? reference)
    {
        reference = null;
        if (position is SlidePosition.Start or SlidePosition.End) return null;

        var path = relativeTo;
        if (string.IsNullOrEmpty(path))
            return new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"{verb} with position '{position}' needs a reference slide. " +
                "Supply \"relativeTo\": \"slide#<id>\", or use position Start or End.",
                target);

        if (!SlideNodeProvider.TryParseSlideId(path!, out var slideId))
            return new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"'{path}' is not a slide path. Slide paths read 'slide#<id>' and come from inspect_document.nodes.",
                target);

        reference = PowerPointModel.Slide(context.Package, slideId);
        return reference is null
            ? new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No slide with path '{path}'. Slide paths come from inspect_document.nodes.",
                target)
            : null;
    }

    /// <summary>Resolves the slide a verb acts on from its own target anchor.</summary>
    public static SlideRef? Target(ApplyContext context, NodeAnchor anchor) =>
        SlideNodeProvider.TryParseSlideId(anchor.Path, out var slideId)
            ? PowerPointModel.Slide(context.Package, slideId)
            : null;

    public static ValidationError NoSuchSlide(NodeAnchor anchor) => new(
        ValidationErrorCodes.AnchorNotFound,
        $"No slide with path '{anchor.Path}'. Slide paths come from inspect_document.nodes.",
        anchor);
}

/// <summary>
/// Adds a slide, built from one of the deck's own layouts.
/// </summary>
/// <remarks>
/// The slide states only its text: position, size, font, and bullet styling are inherited
/// from the layout's placeholders, so a generated deck looks like the template rather than
/// like hand-placed text boxes. Several of these in one plan is how a whole deck is
/// authored in a single call.
/// </remarks>
internal sealed class SlideInsertHandler : IOperationHandler
{
    /// <inheritdoc />
    public bool CanHandle(PlanOperation operation) => operation is InsertSlideOp;

    /// <inheritdoc />
    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (InsertSlideOp)operation;

        var layoutName = op.Slide.Layout ?? SlideLayouts.DefaultFor(op.Slide);
        if (!SlideLayouts.All.Any(l => string.Equals(l.Name, layoutName, StringComparison.OrdinalIgnoreCase)))
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"Unknown layout '{layoutName}'. Expected one of: {SlideLayouts.Names}.",
                op.Target));

        var error = SlideList.ResolveReference(
            context, op.Position, RelativePath(op), op.Target, "insertSlide", out var reference);
        if (error is not null) return OperationPreview.Fail(error);

        var summary = op.Slide.Title is { Length: > 0 } title ? title : "(untitled)";
        var detail = op.Slide.Body.Count > 0 ? $", {op.Slide.Body.Count} bullet(s)" : string.Empty;
        if (op.Slide.Notes is { Length: > 0 }) detail += ", notes";

        return OperationPreview.Ok(new ProposedChange
        {
            Target = op.Target,
            Verb = "insertSlide",
            Before = string.Empty,
            After = $"[{layoutName}] {summary}{detail}",
            Context = Describe(op.Position, reference),
            BlastRadius = 1
        });
    }

    /// <inheritdoc />
    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (InsertSlideOp)operation;
        var layoutName = op.Slide.Layout ?? SlideLayouts.DefaultFor(op.Slide);

        var main = PowerPointModel.Main(context.Package);
        var master = main.SlideMasterParts.FirstOrDefault()
            ?? throw new InvalidOperationException("The presentation has no slide master.");

        // A deck from someone else's template may not define the layout asked for; using
        // whatever it does define beats refusing the edit, and beats inventing geometry.
        var layoutPart = SlideLayouts.Find(master, layoutName)
            ?? master.SlideLayoutParts.FirstOrDefault()
            ?? throw new InvalidOperationException("The slide master defines no layouts.");

        var slidePart = main.AddNewPart<SlidePart>();
        slidePart.Slide = SlideBuilder.Build(op.Slide, layoutPart);
        slidePart.AddPart(layoutPart);

        if (op.Slide.Notes is { Length: > 0 } notes)
            SlideBuilder.AddNotes(slidePart, notes);

        var list = SlideList.Require(context.Package);
        var entry = new SlideId
        {
            Id = SlideList.NextSlideId(list),
            RelationshipId = main.GetIdOfPart(slidePart)
        };

        SlideList.ResolveReference(
            context, op.Position, RelativePath(op), op.Target, "insertSlide", out var reference);
        SlideList.Place(list, entry, op.Position,
            reference is null ? null : SlideList.EntryFor(list, reference.SlideId));

        // A sectioned deck must claim the new slide, or PowerPoint sees a slide belonging
        // to no section and offers to repair the file.
        Sections.Reconcile(context.Package);
    }

    /// <summary>The reference slide comes from the target anchor for this verb.</summary>
    private static string? RelativePath(InsertSlideOp op) =>
        op.Target is NodeAnchor { Kind: "slide" } anchor ? anchor.Path : null;

    internal static string Describe(SlidePosition position, SlideRef? reference) => position switch
    {
        SlidePosition.Start => "at the start",
        SlidePosition.Before when reference is not null => $"before slide {reference.Number}",
        SlidePosition.After when reference is not null => $"after slide {reference.Number}",
        _ => "at the end"
    };
}

/// <summary>Removes a slide, its notes, and its relationships.</summary>
internal sealed class SlideRemoveHandler : IOperationHandler
{
    /// <inheritdoc />
    public bool CanHandle(PlanOperation operation) =>
        operation is RemoveSlideOp { Target: NodeAnchor { Kind: "slide" } };

    /// <inheritdoc />
    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var anchor = (NodeAnchor)operation.Target!;
        var slide = SlideList.Target(context, anchor);
        if (slide is null) return OperationPreview.Fail(SlideList.NoSuchSlide(anchor));

        // Refusing to empty the deck entirely: PowerPoint cannot open a presentation with
        // no slides, so allowing it would produce a file the user cannot recover from.
        var total = PowerPointModel.Slides(context.Package).Count();
        if (total <= 1)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "removeSlide would leave the deck with no slides, which PowerPoint cannot open. " +
                "Add a replacement slide in the same plan, or delete the file instead.",
                anchor));

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "removeSlide",
            Before = $"slide {slide.Number}",
            After = string.Empty,
            Context = $"slide {slide.Number} of {total}",
            BlastRadius = 1
        });
    }

    /// <inheritdoc />
    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var anchor = (NodeAnchor)operation.Target!;
        var slide = SlideList.Target(context, anchor)
            ?? throw new InvalidOperationException($"Slide '{anchor.Path}' vanished before apply.");

        var main = PowerPointModel.Main(context.Package);
        var list = SlideList.Require(context.Package);

        SlideList.EntryFor(list, slide.SlideId)?.Remove();
        // Drop the part last: the list entry is what makes it reachable, and deleting the
        // part first would briefly leave a dangling relationship id in the presentation.
        main.DeletePart(slide.Part);

        // Its id has to leave the section that listed it too, or the grouping references
        // a slide the deck no longer has.
        Sections.Reconcile(context.Package);
    }
}

/// <summary>Reorders the deck by moving one slide's entry in <c>p:sldIdLst</c>.</summary>
internal sealed class SlideMoveHandler : IOperationHandler
{
    /// <inheritdoc />
    public bool CanHandle(PlanOperation operation) =>
        operation is MoveSlideOp { Target: NodeAnchor { Kind: "slide" } };

    /// <inheritdoc />
    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (MoveSlideOp)operation;
        var anchor = (NodeAnchor)op.Target!;

        var slide = SlideList.Target(context, anchor);
        if (slide is null) return OperationPreview.Fail(SlideList.NoSuchSlide(anchor));

        var error = SlideList.ResolveReference(
            context, op.Position, op.RelativeTo, anchor, "moveSlide", out var reference);
        if (error is not null) return OperationPreview.Fail(error);

        if (reference is not null && reference.SlideId == slide.SlideId)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "moveSlide cannot place a slide relative to itself.", anchor));

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "moveSlide",
            Before = $"position {slide.Number}",
            After = SlideInsertHandler.Describe(op.Position, reference),
            Context = $"slide {slide.Number}",
            BlastRadius = 1
        });
    }

    /// <inheritdoc />
    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (MoveSlideOp)operation;
        var anchor = (NodeAnchor)op.Target!;

        var slide = SlideList.Target(context, anchor)
            ?? throw new InvalidOperationException($"Slide '{anchor.Path}' vanished before apply.");

        var list = SlideList.Require(context.Package);
        var entry = SlideList.EntryFor(list, slide.SlideId)
            ?? throw new InvalidOperationException($"Slide '{anchor.Path}' has no list entry.");

        SlideList.ResolveReference(
            context, op.Position, op.RelativeTo, anchor, "moveSlide", out var reference);
        var referenceEntry = reference is null ? null : SlideList.EntryFor(list, reference.SlideId);

        // Detach before placing: the reference must be located while the list still holds
        // it, but re-inserting an attached element is what actually moves it.
        entry.Remove();
        SlideList.Place(list, entry, op.Position, referenceEntry);

        // Moving a slide moves it between sections: the grouping follows deck order.
        Sections.Reconcile(context.Package);
    }
}

/// <summary>Copies a slide, content and all, into a new slide with its own ids.</summary>
internal sealed class SlideDuplicateHandler : IOperationHandler
{
    /// <inheritdoc />
    public bool CanHandle(PlanOperation operation) =>
        operation is DuplicateSlideOp { Target: NodeAnchor { Kind: "slide" } };

    /// <inheritdoc />
    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (DuplicateSlideOp)operation;
        var anchor = (NodeAnchor)op.Target!;

        var slide = SlideList.Target(context, anchor);
        if (slide is null) return OperationPreview.Fail(SlideList.NoSuchSlide(anchor));

        // Defaulting the reference to the slide being copied is what makes the common case
        // - "duplicate this slide" - need nothing but a target.
        var error = SlideList.ResolveReference(
            context, op.Position, op.RelativeTo ?? anchor.Path, anchor, "duplicateSlide", out var reference);
        if (error is not null) return OperationPreview.Fail(error);

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "duplicateSlide",
            Before = $"slide {slide.Number}",
            After = SlideInsertHandler.Describe(op.Position, reference),
            Context = $"slide {slide.Number}",
            BlastRadius = 1
        });
    }

    /// <inheritdoc />
    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (DuplicateSlideOp)operation;
        var anchor = (NodeAnchor)op.Target!;

        var slide = SlideList.Target(context, anchor)
            ?? throw new InvalidOperationException($"Slide '{anchor.Path}' vanished before apply.");

        var main = PowerPointModel.Main(context.Package);
        var copyPart = main.AddNewPart<SlidePart>();
        copyPart.Slide = (Slide)slide.Part.Slide.CloneNode(deep: true);

        // Share the layout and any images: parts are immutable content addressed by
        // relationship id, so two slides pointing at one image is what PowerPoint itself
        // produces. Notes are the exception - they are per-slide content the user edits,
        // so the copy gets its own, and sharing the source's part would both alias the
        // text and fail outright, a slide being allowed only one notes part.
        foreach (var pair in slide.Part.Parts)
        {
            if (pair.OpenXmlPart is NotesSlidePart) continue;
            copyPart.AddPart(pair.OpenXmlPart, pair.RelationshipId);
        }

        if (slide.Part.NotesSlidePart?.NotesSlide is { } notes)
        {
            var notesPart = copyPart.AddNewPart<NotesSlidePart>();
            notesPart.NotesSlide = (NotesSlide)notes.CloneNode(deep: true);
        }

        var list = SlideList.Require(context.Package);
        var entry = new SlideId
        {
            Id = SlideList.NextSlideId(list),
            RelationshipId = main.GetIdOfPart(copyPart)
        };

        SlideList.ResolveReference(
            context, op.Position, op.RelativeTo ?? anchor.Path, anchor, "duplicateSlide", out var reference);
        SlideList.Place(list, entry, op.Position,
            reference is null ? null : SlideList.EntryFor(list, reference.SlideId));

        // The copy joins the section its neighbour is in, as it does in PowerPoint.
        Sections.Reconcile(context.Package);
    }
}
