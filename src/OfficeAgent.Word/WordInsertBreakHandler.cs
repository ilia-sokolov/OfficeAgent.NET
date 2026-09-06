using DocumentFormat.OpenXml;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.Word;

/// <summary>
/// Inserts a page, column, or section break in a paragraph of its own, before or after an
/// anchored paragraph.
/// </summary>
/// <remarks>
/// <para>
/// A page or column break is a <c>w:br</c> in a run. A <em>section</em> break is not a
/// character at all: it is a <c>w:sectPr</c> in a paragraph's properties, and it describes
/// the section that <em>ends</em> there. So the new paragraph carries a copy of the
/// section it splits - the pages before the break keep the geometry, headers, and footers
/// they had - and the body's own <c>w:sectPr</c> goes on governing everything after it.
/// </para>
/// <para>
/// Under <see cref="ChangeMode.Tracked"/> the break paragraph is a redline: rejecting it
/// removes the paragraph and the break with it.
/// </para>
/// </remarks>
internal sealed class WordInsertBreakHandler : IOperationHandler
{
    private readonly TimeProvider _clock;

    public WordInsertBreakHandler(TimeProvider clock) => _clock = clock;

    public bool CanHandle(PlanOperation operation) =>
        operation is InsertBreakOp { Target: TextSpanAnchor };

    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (InsertBreakOp)operation;
        var anchor = (TextSpanAnchor)op.Target;

        var paragraph = WordModel.ResolveParagraph(context, anchor.ParaId);
        if (paragraph is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No paragraph with id '{anchor.ParaId}'.", anchor));

        // A section break belongs to the body's flow: its w:sectPr describes page geometry,
        // which a header, footer, or note part has none of.
        if (IsSectionBreak(op.Kind) && !WordModel.IsInMainBody(paragraph))
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"Paragraph '{anchor.ParaId}' is not in the document body, and only the body has sections. " +
                "Re-inspect and target a paragraph whose location is \"body\".", anchor));

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "insertBreak",
            Before = string.Empty,
            After = $"[{Describe(op.Kind)}]",
            Context = $"{op.Position.ToString().ToLowerInvariant()} paragraph '{anchor.ParaId}'",
            BlastRadius = 1,
            Capability = Capability.Deterministic
        });
    }

    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (InsertBreakOp)operation;
        var anchor = (TextSpanAnchor)op.Target;

        var paragraph = WordModel.ResolveParagraph(context, anchor.ParaId)
            ?? throw new InvalidOperationException($"Paragraph '{anchor.ParaId}' vanished before apply.");

        var carrier = IsSectionBreak(op.Kind)
            ? SectionBreakParagraph(context, op.Kind)
            : CharacterBreakParagraph(op.Kind);

        if (op.Position == InsertPosition.Before)
            paragraph.InsertBeforeSelf(carrier);
        else
            paragraph.InsertAfterSelf(carrier);

        if (WordRevisionMarker.IsTracked(op.Mode))
            new WordRevisionMarker(context.Package, _clock).MarkParagraphInserted(carrier);
    }

    private static Paragraph CharacterBreakParagraph(BreakKind kind) =>
        new(new Run(new Break
        {
            Type = kind == BreakKind.Column ? BreakValues.Column : BreakValues.Page
        }));

    /// <summary>
    /// A paragraph carrying a copy of the section it splits, retyped to the requested
    /// break. Everything the section said about the pages before this point - size,
    /// margins, headers, footers - is preserved by the copy.
    /// </summary>
    private static Paragraph SectionBreakParagraph(ApplyContext context, BreakKind kind)
    {
        var body = WordSections.Require(WordModel.Main(context.Package));
        var section = (SectionProperties)body.CloneNode(deep: true);

        WordSections.Replace(section, new SectionType { Val = TypeOf(kind) });

        return new Paragraph(new ParagraphProperties(section));
    }

    private static SectionMarkValues TypeOf(BreakKind kind) => kind switch
    {
        BreakKind.SectionContinuous => SectionMarkValues.Continuous,
        BreakKind.SectionEvenPage => SectionMarkValues.EvenPage,
        BreakKind.SectionOddPage => SectionMarkValues.OddPage,
        _ => SectionMarkValues.NextPage
    };

    private static bool IsSectionBreak(BreakKind kind) =>
        kind is not BreakKind.Page and not BreakKind.Column;

    private static string Describe(BreakKind kind) => kind switch
    {
        BreakKind.Page => "page break",
        BreakKind.Column => "column break",
        BreakKind.SectionContinuous => "section break (continuous)",
        BreakKind.SectionEvenPage => "section break (even page)",
        BreakKind.SectionOddPage => "section break (odd page)",
        _ => "section break (next page)"
    };
}
