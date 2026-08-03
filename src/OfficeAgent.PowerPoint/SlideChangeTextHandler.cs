using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeAgent.PowerPoint;

/// <summary>
/// Replaces an anchored, content-verified text span on a slide, in a table cell, or in
/// the slide notes.
/// </summary>
/// <remarks>
/// The span is isolated through the Core <see cref="TextBodyEngine"/>, so text that runs
/// across several <c>a:r</c> elements is replaced as one edit and the character
/// formatting of every run the span does not fully cover survives. PresentationML has no
/// redline vocabulary - <c>ChangeMode.Tracked</c> has no meaning here - so a tracked
/// request is refused rather than silently written as a direct edit.
/// </remarks>
internal sealed class SlideChangeTextHandler : IOperationHandler
{
    /// <inheritdoc />
    public bool CanHandle(PlanOperation operation) =>
        operation is ChangeTextOp { Target: TextSpanAnchor };

    /// <inheritdoc />
    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (ChangeTextOp)operation;
        var anchor = (TextSpanAnchor)op.Target;

        if (op.Mode == ChangeMode.Tracked)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "PowerPoint has no tracked-changes representation, so mode 'Tracked' cannot be honoured. " +
                "Re-issue this operation with mode 'Direct', or add a comment to record the intent.",
                anchor));

        var paragraph = PowerPointModel.ResolveParagraph(context, anchor.ParaId);
        if (paragraph is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No paragraph with id '{anchor.ParaId}'.", anchor));

        var text = PowerPointModel.TextOf(paragraph.Paragraph);

        // An empty 'expect' means "this paragraph is blank; make it say something" - the
        // only way to fill the empty placeholder a new deck starts with, since a slide
        // has no paragraph-inserting verb. It stays content-verified: if the paragraph
        // turns out to hold text, the deck drifted and the operation fails.
        if (string.IsNullOrEmpty(anchor.Expect))
            return text.Length > 0
                ? OperationPreview.Fail(new ValidationError(
                    ValidationErrorCodes.ExpectMismatch,
                    $"Paragraph '{anchor.ParaId}' was expected to be empty but reads '{Shorten(text)}'. " +
                    "Set 'expect' to the text you mean to replace.",
                    anchor))
                : OperationPreview.Ok(new ProposedChange
                {
                    Target = anchor,
                    Verb = "changeText",
                    Before = string.Empty,
                    After = op.With,
                    Context = $"slide {paragraph.Slide.Number} ({paragraph.Location}): empty paragraph",
                    BlastRadius = 1
                });

        var comparison = PowerPointModel.Comparison(caseSensitive: true);
        var occurrences = PowerPointModel.Text.CountOccurrences(text, anchor.Expect, comparison);

        if (occurrences == 0)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.ExpectMismatch,
                $"Expected text '{anchor.Expect}' not found in paragraph '{anchor.ParaId}' (the deck drifted).",
                anchor));

        var start = PowerPointModel.Text.IndexOfOccurrence(text, anchor.Expect, anchor.Occurrence, comparison);
        if (start < 0)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AmbiguousAnchor,
                $"Occurrence {anchor.Occurrence} of '{anchor.Expect}' does not exist ({occurrences} found).",
                anchor));

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "changeText",
            Before = anchor.Expect,
            After = op.With,
            Context = $"slide {paragraph.Slide.Number} ({paragraph.Location}): " +
                      PowerPointModel.Snippet(text, start, anchor.Expect.Length),
            BlastRadius = 1
        });
    }

    /// <inheritdoc />
    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (ChangeTextOp)operation;
        var anchor = (TextSpanAnchor)op.Target;

        var paragraph = PowerPointModel.ResolveParagraph(context, anchor.ParaId)
            ?? throw new InvalidOperationException($"Paragraph '{anchor.ParaId}' vanished before apply.");

        var text = PowerPointModel.TextOf(paragraph.Paragraph);

        if (string.IsNullOrEmpty(anchor.Expect))
        {
            SetParagraphText(paragraph.Paragraph, op.With);
            return;
        }

        var comparison = PowerPointModel.Comparison(caseSensitive: true);
        var start = PowerPointModel.Text.IndexOfOccurrence(text, anchor.Expect, anchor.Occurrence, comparison);
        if (start < 0)
            throw new InvalidOperationException($"Expected text '{anchor.Expect}' not found at apply time.");

        // IsolateSpan splits runs so the covered set is exactly the span: the first
        // covered run keeps its formatting and receives the replacement, and the rest
        // are dropped.
        var covered = PowerPointModel.Text.IsolateSpan(paragraph.Paragraph, start, anchor.Expect.Length);
        if (covered.Count == 0)
            throw new InvalidOperationException($"Span for '{anchor.Expect}' could not be isolated.");

        PowerPointModel.Dialect.SetRunText(covered[0], op.With);
        for (var i = 1; i < covered.Count; i++)
            covered[i].Remove();
    }

    /// <summary>Trims paragraph text for an error message.</summary>
    private static string Shorten(string text) =>
        text.Length <= 40 ? text : text.Substring(0, 40) + "…";

    /// <summary>
    /// Writes text into a paragraph that has none. An empty <c>a:p</c> may still carry an
    /// <c>a:endParaRPr</c> describing how the next typed character should look; the run is
    /// placed before it, both because the schema requires it last and because reusing that
    /// formatting is what an author would expect.
    /// </summary>
    private static void SetParagraphText(A.Paragraph paragraph, string text)
    {
        var existing = paragraph.Elements<A.Run>().FirstOrDefault();
        if (existing is not null)
        {
            PowerPointModel.Dialect.SetRunText(existing, text);
            foreach (var extra in paragraph.Elements<A.Run>().Skip(1).ToList())
                extra.Remove();
            return;
        }

        var endProperties = paragraph.GetFirstChild<A.EndParagraphRunProperties>();
        var run = new A.Run(
            endProperties is null
                ? new A.RunProperties { Language = "en-US" }
                : (A.RunProperties)CloneAsRunProperties(endProperties),
            new A.Text(text));

        if (endProperties is null) paragraph.AppendChild(run);
        else endProperties.InsertBeforeSelf(run);
    }

    /// <summary>
    /// Re-expresses <c>a:endParaRPr</c> as the <c>a:rPr</c> of a real run: the two carry
    /// the same attributes, so the text an agent adds inherits the look the placeholder
    /// was already primed with.
    /// </summary>
    private static A.RunProperties CloneAsRunProperties(A.EndParagraphRunProperties source)
    {
        var properties = new A.RunProperties();
        properties.InnerXml = source.InnerXml;
        foreach (var attribute in source.GetAttributes())
            properties.SetAttribute(attribute);
        return properties;
    }
}
