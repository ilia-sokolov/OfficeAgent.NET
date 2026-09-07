using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;

namespace OfficeAgent.Word;

/// <summary>
/// Defines or updates a style in <c>styles.xml</c>, so a look is stated once instead of
/// being copied onto every paragraph that wants it.
/// </summary>
/// <remarks>
/// Defining a style that already exists updates it in place: properties the operation does
/// not mention keep the values they had, which is what makes "make the headings a bit
/// tighter" a one-property plan rather than a full restatement of the style.
/// </remarks>
internal sealed class DefineStyleHandler : IOperationHandler
{
    /// <summary>The style types a plan may ask for, for an error message that helps.</summary>
    internal const string TypeNames = "paragraph, character, table";

    public bool CanHandle(PlanOperation operation) => operation is DefineStyleOp;

    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (DefineStyleOp)operation;

        if (string.IsNullOrWhiteSpace(op.StyleId))
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "defineStyle requires a styleId.", op.Target));

        if (op.StyleId.Any(char.IsWhiteSpace))
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"styleId '{op.StyleId}' cannot contain whitespace; it is an identifier, not the display name. " +
                "Use 'name' for the name shown in Word.", op.Target));

        if (ParseType(op.Type) is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"'{op.Type}' is not a style type. Expected one of: {TypeNames}.", op.Target));

        if (!FormatHandler.AreEdges(op.BorderEdges))
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"'{op.BorderEdges}' is not a border edge list. Expected a comma-separated subset of: {FormatHandler.BorderEdgeNames}.",
                op.Target));

        // WordprocessingML has no style-level highlight: w:highlight is a property of a run,
        // and a style's w:rPr refuses it outright. Dropping it quietly would hand back a
        // style that does not do what the plan asked for.
        if (op.Highlight is not null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "A style cannot carry a highlight; w:highlight belongs to a run. Use 'color' for the " +
                "text colour, or highlight the span itself with the format verb.", op.Target));

        if (op.OutlineLevel is { } level && (level < 1 || level > 9))
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"outlineLevel must be between 1 and 9; got {level}.", op.Target));

        if (string.Equals(op.BasedOn, op.StyleId, StringComparison.Ordinal))
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"Style '{op.StyleId}' cannot be based on itself.", op.Target));

        var existing = Find(context.Package, op.StyleId);
        return OperationPreview.Ok(new ProposedChange
        {
            Target = op.Target,
            Verb = "defineStyle",
            Before = existing is null ? "(undefined)" : Describe(existing),
            After = $"{op.Type} style '{op.StyleId}'",
            Context = op.StyleId,

            // One style, however many paragraphs eventually carry it: the blast radius of
            // the operation is the definition, not the document's future use of it.
            BlastRadius = 1
        });
    }

    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (DefineStyleOp)operation;
        var styles = StylesOf(context.Package);
        var type = ParseType(op.Type)
            ?? throw new InvalidOperationException($"Style type '{op.Type}' is not supported.");

        var style = Find(context.Package, op.StyleId);
        var isNew = style is null;
        if (style is null)
        {
            style = new Style { StyleId = op.StyleId, Type = type };
            styles.AppendChild(style);
        }
        else
        {
            style.Type = type;
        }

        // Only when asked, or when there is nothing to keep. Rewriting the name on every
        // update would rename "Pull Quote" to "Quote" for a plan that only changed a size.
        if (op.Name is { Length: > 0 } name)
            Place(style, new StyleName { Val = name });
        else if (isNew)
            Place(style, new StyleName { Val = op.StyleId });

        if (op.BasedOn is { Length: > 0 } basedOn)
            Place(style, new BasedOn { Val = basedOn });
        if (op.Next is { Length: > 0 } next)
            Place(style, new NextParagraphStyle { Val = next });

        // A style Word does not offer in its gallery is one an author cannot find, and an
        // author who cannot find it reaches for direct formatting instead.
        if (op.Quick) Place(style, new PrimaryStyle());
        else style.GetFirstChild<PrimaryStyle>()?.Remove();

        if (HasParagraphProperties(op))
        {
            var properties = style.GetFirstChild<StyleParagraphProperties>();
            if (properties is null)
            {
                properties = new StyleParagraphProperties();
                Place(style, properties);
            }

            FormatHandler.WriteParagraphFormatting(properties, op);
            if (op.OutlineLevel is { } level)
            {
                properties.GetFirstChild<OutlineLevel>()?.Remove();

                // w:outlineLvl is zero-based; the plan counts headings from 1, the way a
                // document's own numbering does.
                FormatHandler.PlaceParagraphProperty(properties, new OutlineLevel { Val = level - 1 });
            }
        }

        if (HasRunProperties(op))
        {
            var properties = style.GetFirstChild<StyleRunProperties>();
            if (properties is null)
            {
                properties = new StyleRunProperties();
                Place(style, properties);
            }

            FormatHandler.WriteRunProperties(properties, op);
        }
    }

    /// <summary>
    /// Returns the style with this id, whatever its type. Ids are unique across types in
    /// <c>styles.xml</c>, so a lookup that also matched on type would happily create a
    /// second style Word then refuses to open.
    /// </summary>
    internal static Style? Find(IOpenXmlPackage package, string styleId) =>
        WordModel.Doc(package).MainDocumentPart?.StyleDefinitionsPart?.Styles?
            .Elements<Style>()
            .FirstOrDefault(style => string.Equals(style.StyleId?.Value, styleId, StringComparison.Ordinal));

    internal static StyleValues? ParseType(string? type) => type?.Trim().ToLowerInvariant() switch
    {
        null or "" or "paragraph" => StyleValues.Paragraph,
        "character" => StyleValues.Character,
        "table" => StyleValues.Table,
        _ => null
    };

    private static Styles StylesOf(IOpenXmlPackage package)
    {
        var main = WordModel.Doc(package).MainDocumentPart
            ?? throw new InvalidOperationException("The document has no main part.");
        var part = main.StyleDefinitionsPart ?? main.AddNewPart<StyleDefinitionsPart>();
        return part.Styles ??= new Styles();
    }

    private static bool HasRunProperties(DefineStyleOp op) =>
        op.FontFamily is not null || op.SizeHalfPoints is not null || op.Bold is not null
        || op.Italic is not null || op.Underline is not null || op.Highlight is not null
        || op.Color is not null;

    private static bool HasParagraphProperties(DefineStyleOp op) =>
        op.Alignment is not null || op.IndentLeftTwips is not null || op.IndentRightTwips is not null
        || op.IndentFirstLineTwips is not null || op.SpacingBeforeTwips is not null
        || op.SpacingAfterTwips is not null || op.OutlineLevel is not null
        || op.BorderStyle is not null || op.BorderSizeEighths is not null
        || op.BorderColor is not null || op.BorderEdges is not null;

    /// <summary>
    /// The order <c>CT_Style</c> declares the children this handler writes. Appending
    /// instead - <c>w:basedOn</c> after <c>w:pPr</c>, say - produces a style Word offers
    /// to repair, the same way an out-of-order <c>w:pPr</c> does.
    /// </summary>
    private static readonly Type[] StyleChildOrder =
    {
        typeof(StyleName), typeof(BasedOn), typeof(NextParagraphStyle),
        typeof(PrimaryStyle), typeof(StyleParagraphProperties), typeof(StyleRunProperties)
    };

    private static void Place(Style style, OpenXmlElement child)
    {
        var same = style.ChildElements.FirstOrDefault(e => e.GetType() == child.GetType());
        if (ReferenceEquals(same, child)) return;
        same?.Remove();

        var rank = Rank(child);
        if (rank < 0)
        {
            style.AppendChild(child);
            return;
        }

        foreach (var existing in style.ChildElements)
        {
            var existingRank = Rank(existing);
            if (existingRank >= 0 && existingRank > rank)
            {
                style.InsertBefore(child, existing);
                return;
            }
        }

        style.AppendChild(child);
    }

    private static int Rank(OpenXmlElement element)
    {
        for (var index = 0; index < StyleChildOrder.Length; index++)
            if (StyleChildOrder[index] == element.GetType()) return index;
        return -1;
    }

    private static string Describe(Style style)
    {
        var name = style.StyleName?.Val?.Value;
        var basedOn = style.GetFirstChild<BasedOn>()?.Val?.Value;
        return basedOn is null ? name ?? style.StyleId?.Value ?? "(style)" : $"{name} based on {basedOn}";
    }
}
