using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;

namespace OfficeAgent.Word;

/// <summary>
/// Sets the running header and footer, the page number, and whether the first page has its
/// own.
/// </summary>
/// <remarks>
/// A page number is written as a <c>PAGE</c> field rather than a number, so it stays right
/// as the document grows. That is three runs - begin, instruction, end - plus a cached
/// result Word replaces on open; a document that omits the cached value shows an empty
/// header until the reader presses F9.
/// </remarks>
internal sealed class WordHeaderFooterHandler : IOperationHandler
{
    /// <inheritdoc />
    public bool CanHandle(PlanOperation operation) => operation is HeaderFooterOp;

    /// <inheritdoc />
    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (HeaderFooterOp)operation;

        if (Unsupported(op) is { } unsupported)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation, unsupported, op.Target));

        if (op.Scope is not null && ScopeOf(op.Scope) is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"'{op.Scope}' is not a header scope. Expected default, firstPage, or evenPage.",
                op.Target));

        if (op.Alignment is not null && !IsAlignment(op.Alignment))
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"'{op.Alignment}' is not a header alignment. Expected left, center, right, or edges.",
                op.Target));

        if (op.Header is null && op.Footer is null && op.ShowPageNumber is null &&
            op.DifferentFirstPage is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "headerFooter requires at least one of: header, footer, showPageNumber, differentFirstPage.",
                op.Target));

        return OperationPreview.Ok(new ProposedChange
        {
            Target = new NodeAnchor { Kind = "page", Path = "page#headerFooter" },
            Verb = "headerFooter",
            Before = string.Empty,
            After = Describe(op),
            Context = op.Scope ?? "default",
            BlastRadius = 1
        });
    }

    /// <inheritdoc />
    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (HeaderFooterOp)operation;
        var main = WordModel.Doc(context.Package).MainDocumentPart!;
        var section = WordSections.Require(main);

        // Set first, so a header written for the first page in the same plan lands in a
        // section that already knows it has one.
        if (op.DifferentFirstPage is bool distinct)
            WordSections.SetTitlePage(section, distinct);

        var kind = ScopeOf(op.Scope) ?? HeaderFooterValues.Default;

        if (op.Header is not null)
        {
            var header = WordSections.HeaderFor(main, section, kind);
            Rewrite(header, op.Header, pageNumber: false, op.Alignment);
        }

        if (op.Footer is not null || op.ShowPageNumber is not null)
        {
            var footer = WordSections.FooterFor(main, section, kind);
            Rewrite(footer, op.Footer ?? string.Empty, op.ShowPageNumber == true, op.Alignment);
        }
    }

    /// <summary>
    /// Replaces the text of a header or footer, keeping any background drawing already
    /// there. Rewriting the whole part instead would take the background out with the text.
    /// </summary>
    private static void Rewrite(OpenXmlPartRootElement container, string text, bool pageNumber, string? alignment)
    {
        foreach (var paragraph in container.Elements<Paragraph>().ToList())
            if (!paragraph.Descendants<Drawing>().Any())
                paragraph.Remove();

        if (text.Length == 0 && !pageNumber) return;

        var edges = string.Equals(alignment, "edges", StringComparison.OrdinalIgnoreCase);
        var properties = new ParagraphProperties();

        if (edges && pageNumber && text.Length > 0)
            // A right tab at the margin is what puts the number on the far side of the same
            // line as the text, which is the shape of most running heads.
            properties.Append(new Tabs(new TabStop
            {
                Val = TabStopValues.Right,
                Position = RightTabPosition(container)
            }));
        else if (Justification(alignment) is { } justification)
            properties.Append(new Justification { Val = justification });

        var line = new Paragraph(properties);

        if (text.Length > 0)
            line.Append(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));

        if (pageNumber)
        {
            if (edges && text.Length > 0) line.Append(new Run(new TabChar()));
            else if (text.Length > 0)
                line.Append(new Run(new Text("  ") { Space = SpaceProcessingModeValues.Preserve }));

            foreach (var run in PageNumberField()) line.Append(run);
        }

        container.AppendChild(line);
    }

    /// <summary>
    /// The <c>PAGE</c> field, as the run sequence Word writes: begin, instruction, separate,
    /// a cached result, end.
    /// </summary>
    private static IEnumerable<Run> PageNumberField()
    {
        yield return new Run(new FieldChar { FieldCharType = FieldCharValues.Begin });
        yield return new Run(new FieldCode(" PAGE ") { Space = SpaceProcessingModeValues.Preserve });
        yield return new Run(new FieldChar { FieldCharType = FieldCharValues.Separate });
        yield return new Run(new Text("1"));
        yield return new Run(new FieldChar { FieldCharType = FieldCharValues.End });
    }

    /// <summary>
    /// Where the right edge of the text is, in twips, for the tab that pushes the page
    /// number to it.
    /// </summary>
    private static int RightTabPosition(OpenXmlPartRootElement container)
    {
        const int letterWidth = 12240;
        const int defaultMargin = 1440;

        if (container.OpenXmlPart?.OpenXmlPackage is DocumentFormat.OpenXml.Packaging.WordprocessingDocument doc &&
            doc.MainDocumentPart?.Document.Body?.GetFirstChild<SectionProperties>() is { } section)
        {
            var size = section.GetFirstChild<PageSize>()?.Width?.Value ?? letterWidth;
            var margin = section.GetFirstChild<PageMargin>();
            var left = margin?.Left?.Value ?? defaultMargin;
            var right = margin?.Right?.Value ?? defaultMargin;
            return (int)(size - left - right);
        }

        return letterWidth - (defaultMargin * 2);
    }

    private static JustificationValues? Justification(string? alignment) =>
        alignment?.Trim().ToLowerInvariant() switch
        {
            "center" => JustificationValues.Center,
            "right" => JustificationValues.Right,
            "left" => JustificationValues.Left,
            _ => null
        };

    private static bool IsAlignment(string value) =>
        value.Trim().ToLowerInvariant() is "left" or "center" or "right" or "edges";

    private static HeaderFooterValues? ScopeOf(string? scope) =>
        scope?.Trim().ToLowerInvariant() switch
        {
            null or "" or "default" => HeaderFooterValues.Default,
            "firstpage" or "first" => HeaderFooterValues.First,
            "evenpage" or "even" => HeaderFooterValues.Even,
            _ => null
        };

    /// <summary>
    /// Names the deck-only settings rather than dropping them, so a plan written for a deck
    /// and pointed at a document says so instead of half-working.
    /// </summary>
    private static string? Unsupported(HeaderFooterOp op)
    {
        var rejected = new List<string>();
        if (op.ShowSlideNumber is not null)
            rejected.Add("showSlideNumber (a document has pages - use showPageNumber)");
        if (op.ShowFooter is not null)
            rejected.Add("showFooter (a Word footer is shown when it has content; clear it with footer: \"\")");
        if (op.ShowDateTime is not null || op.DateTime is not null)
            rejected.Add("showDateTime/dateTime (put the date in the header or footer text)");

        return rejected.Count == 0
            ? null
            : $"The Word module cannot apply {string.Join(", ", rejected)}.";
    }

    private static string Describe(HeaderFooterOp op)
    {
        var parts = new List<string>();
        if (op.Header is { Length: > 0 } header) parts.Add($"header \"{header}\"");
        else if (op.Header is not null) parts.Add("header cleared");
        if (op.Footer is { Length: > 0 } footer) parts.Add($"footer \"{footer}\"");
        else if (op.Footer is not null) parts.Add("footer cleared");
        if (op.ShowPageNumber is bool number) parts.Add(number ? "page number" : "no page number");
        if (op.DifferentFirstPage is bool first) parts.Add(first ? "distinct first page" : "one header throughout");
        return string.Join(", ", parts);
    }
}
