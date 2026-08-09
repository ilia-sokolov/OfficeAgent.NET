using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeAgent.PowerPoint;

/// <summary>
/// Fills a named slot in a template deck.
/// </summary>
/// <remarks>
/// WordprocessingML has content controls and bookmarks; PresentationML has neither. What a
/// template author does have is the <em>shape name</em> - the label PowerPoint shows in the
/// Selection Pane and that templates set deliberately ("ClientName", "EngagementDate").
/// That is the addressable slot here, so <c>fill</c> means the same thing it does in Word:
/// put this value in the slot called that, without the caller needing to know where on the
/// slide it sits or what it currently says.
/// <para>
/// Names are not unique across a deck. A bare name that matches more than one shape is
/// refused as ambiguous rather than filling an arbitrary one; qualify it as
/// <c>slide256/ClientName</c> to pick.
/// </para>
/// </remarks>
internal sealed class SlideFillHandler : IOperationHandler
{
    /// <inheritdoc />
    public bool CanHandle(PlanOperation operation) =>
        operation is FillOp { Target: StructuralAnchor };

    /// <inheritdoc />
    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (FillOp)operation;
        var anchor = (StructuralAnchor)op.Target;

        var matches = Matching(context.Package, anchor.Tag).ToList();

        if (matches.Count == 0)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No shape named '{anchor.Tag}'. Slot names come from inspect_document.structuralAnchors; " +
                "they are the shape names a template sets, shown in PowerPoint's Selection Pane.",
                anchor));

        if (matches.Count > 1)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AmbiguousAnchor,
                $"'{anchor.Tag}' names {matches.Count} shapes ({string.Join(", ", matches.Select(m => m.Qualified))}). " +
                "Qualify the slot with its slide, for example \"slide256/" + anchor.Tag + "\".",
                anchor));

        var slot = matches[0];
        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "fill",
            Before = slot.Text,
            After = op.Value,
            Context = $"slide {slot.Slide.Number}, shape '{slot.Name}'",
            BlastRadius = 1
        });
    }

    /// <inheritdoc />
    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (FillOp)operation;
        var anchor = (StructuralAnchor)op.Target;

        var slot = Matching(context.Package, anchor.Tag).Single();
        var body = slot.Body;

        // Keep the first paragraph so its properties - bullet, alignment, level - survive,
        // and drop the rest: a filled slot holds the value it was given, not the value
        // plus whatever the template left behind.
        var first = body.Elements<A.Paragraph>().FirstOrDefault();
        foreach (var extra in body.Elements<A.Paragraph>().Skip(1).ToList())
            extra.Remove();

        if (first is null)
        {
            first = new A.Paragraph();
            body.Append(first);
        }

        var template = first.Elements<A.Run>().FirstOrDefault();
        var runProperties = template?.RunProperties is { } rp
            ? (A.RunProperties)rp.CloneNode(deep: true)
            : new A.RunProperties { Language = "en-US" };

        foreach (var child in first.Elements<A.Run>().ToList()) child.Remove();
        foreach (var child in first.Elements<A.Field>().ToList()) child.Remove();
        foreach (var child in first.Elements<A.Break>().ToList()) child.Remove();

        first.Append(new A.Run(runProperties, new A.Text(op.Value)));
    }

    /// <summary>
    /// Every fillable slot in the deck: a named shape carrying a text body. Placeholders
    /// count - a template's "Title 1" is as fillable as a custom-named box.
    /// </summary>
    public static IEnumerable<SlotRef> Slots(IOpenXmlPackage package)
    {
        foreach (var slide in PowerPointModel.Slides(package))
        {
            var tree = slide.Part.Slide.CommonSlideData?.ShapeTree;
            if (tree is null) continue;

            foreach (var shape in tree.Elements<Shape>())
            {
                var name = PowerPointModel.ShapeNameOf(shape);
                if (name.Length == 0 || shape.TextBody is null) continue;

                yield return new SlotRef(slide, name, shape.TextBody);
            }
        }
    }

    /// <summary>Slots matching a tag, which may be bare or qualified as <c>slide256/Name</c>.</summary>
    private static IEnumerable<SlotRef> Matching(IOpenXmlPackage package, string tag)
    {
        if (string.IsNullOrEmpty(tag)) return Enumerable.Empty<SlotRef>();

        var slash = tag.IndexOf('/');
        if (slash > 0 && SlideNodeProvider.TryParseSlideId(tag.Substring(0, slash), out var slideId))
        {
            var name = tag.Substring(slash + 1);
            return Slots(package).Where(s =>
                s.Slide.SlideId == slideId &&
                string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase));
        }

        return Slots(package).Where(s => string.Equals(s.Name, tag, StringComparison.OrdinalIgnoreCase));
    }
}

/// <summary>One named, fillable text slot.</summary>
internal sealed class SlotRef
{
    public SlotRef(SlideRef slide, string name, TextBody body)
    {
        Slide = slide;
        Name = name;
        Body = body;
    }

    public SlideRef Slide { get; }
    public string Name { get; }
    public TextBody Body { get; }

    /// <summary>The unambiguous form, for disambiguating a duplicated name.</summary>
    public string Qualified => $"slide{Slide.SlideId}/{Name}";

    public string Text => string.Join(" ", Body.Elements<A.Paragraph>()
        .Select(PowerPointModel.TextOf)
        .Where(t => t.Length > 0));
}

/// <summary>
/// Copies direct formatting from one span to another.
/// </summary>
/// <remarks>
/// Only <em>direct</em> properties travel - <c>a:pPr</c> and <c>a:rPr</c>. A deck's
/// inherited look comes from its layout and master, which the destination already has and
/// which copying must not disturb; the point of the verb is to make one line look like
/// another, not to re-parent it.
/// </remarks>
internal sealed class SlideCopyStylesHandler : IOperationHandler
{
    /// <inheritdoc />
    public bool CanHandle(PlanOperation operation) =>
        operation is CopyStylesOp { Target: TextSpanAnchor, Source: TextSpanAnchor };

    /// <inheritdoc />
    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (CopyStylesOp)operation;
        var target = (TextSpanAnchor)op.Target;
        var source = (TextSpanAnchor)op.Source;

        if (SlideStyleScope.Invalid(op.Scope) is { } scopeError)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation, scopeError, target));

        if (PowerPointModel.ResolveParagraph(context, source.ParaId) is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No source paragraph with id '{source.ParaId}'.", source));

        var destination = PowerPointModel.ResolveParagraph(context, target.ParaId);
        if (destination is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No paragraph with id '{target.ParaId}'.", target));

        return OperationPreview.Ok(new ProposedChange
        {
            Target = target,
            Verb = "copyStyles",
            Before = string.Empty,
            After = $"formatting of '{source.ParaId}' ({op.Scope})",
            Context = $"slide {destination.Slide.Number} ({destination.Location})",
            BlastRadius = 1
        });
    }

    /// <inheritdoc />
    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (CopyStylesOp)operation;
        var target = (TextSpanAnchor)op.Target;
        var source = (TextSpanAnchor)op.Source;

        var from = PowerPointModel.ResolveParagraph(context, source.ParaId)!;
        var to = PowerPointModel.ResolveParagraph(context, target.ParaId)!;

        if (SlideStyleScope.Includes(op.Scope, "paragraph"))
        {
            to.Paragraph.ParagraphProperties?.Remove();
            if (from.Paragraph.ParagraphProperties is { } properties)
                to.Paragraph.InsertAt((A.ParagraphProperties)properties.CloneNode(deep: true), 0);
        }

        if (!SlideStyleScope.Includes(op.Scope, "run")) return;

        var donor = from.Paragraph.Elements<A.Run>().FirstOrDefault()?.RunProperties;
        foreach (var run in SlideStyleScope.Runs(to, target))
        {
            run.RunProperties?.Remove();
            if (donor is not null)
                run.InsertAt((A.RunProperties)donor.CloneNode(deep: true), 0);
        }
    }
}

/// <summary>
/// Strips direct formatting so the layout's own styling shows through again - the "reset
/// this to look like the template" verb.
/// </summary>
internal sealed class SlideClearStylesHandler : IOperationHandler
{
    /// <inheritdoc />
    public bool CanHandle(PlanOperation operation) =>
        operation is ClearStylesOp { Target: TextSpanAnchor };

    /// <inheritdoc />
    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (ClearStylesOp)operation;
        var anchor = (TextSpanAnchor)op.Target;

        if (SlideStyleScope.Invalid(op.Scope) is { } scopeError)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation, scopeError, anchor));

        var paragraph = PowerPointModel.ResolveParagraph(context, anchor.ParaId);
        if (paragraph is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No paragraph with id '{anchor.ParaId}'.", anchor));

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "clearStyles",
            Before = PowerPointModel.TextOf(paragraph.Paragraph),
            After = $"direct formatting removed ({op.Scope})",
            Context = $"slide {paragraph.Slide.Number} ({paragraph.Location})",
            BlastRadius = 1
        });
    }

    /// <inheritdoc />
    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (ClearStylesOp)operation;
        var anchor = (TextSpanAnchor)op.Target;

        var paragraph = PowerPointModel.ResolveParagraph(context, anchor.ParaId)!;

        if (SlideStyleScope.Includes(op.Scope, "paragraph"))
            paragraph.Paragraph.ParagraphProperties?.Remove();

        if (!SlideStyleScope.Includes(op.Scope, "run")) return;

        foreach (var run in SlideStyleScope.Runs(paragraph, anchor))
        {
            // The language tag is not formatting - dropping it makes the spell checker
            // treat the text as undetermined, which shows up as spurious squiggles.
            var language = run.RunProperties?.Language?.Value;
            run.RunProperties?.Remove();
            if (language is not null)
                run.InsertAt(new A.RunProperties { Language = language }, 0);
        }
    }
}

/// <summary>Shared scope handling for the two style verbs.</summary>
internal static class SlideStyleScope
{
    public static string? Invalid(string scope) =>
        scope is "run" or "paragraph" or "all"
            ? null
            : $"scope must be 'run', 'paragraph', or 'all'; got '{scope}'.";

    public static bool Includes(string scope, string what) =>
        scope == "all" || scope == what;

    /// <summary>
    /// The runs a span covers, isolating it first so a partial match does not restyle the
    /// whole line. An empty <c>expect</c> means every run in the paragraph.
    /// </summary>
    public static IReadOnlyList<A.Run> Runs(ParagraphRef paragraph, TextSpanAnchor anchor)
    {
        if (string.IsNullOrEmpty(anchor.Expect))
            return paragraph.Paragraph.Elements<A.Run>().ToList();

        var text = PowerPointModel.TextOf(paragraph.Paragraph);
        var start = PowerPointModel.Text.IndexOfOccurrence(
            text, anchor.Expect, anchor.Occurrence, StringComparison.Ordinal);

        return start < 0
            ? Array.Empty<A.Run>()
            : PowerPointModel.Text.IsolateSpan(paragraph.Paragraph, start, anchor.Expect.Length)
                .OfType<A.Run>().ToList();
    }
}
