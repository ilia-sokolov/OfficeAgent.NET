using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.Word;

/// <summary>
/// Turns an anchored, content-verified text span into an external hyperlink: the covered
/// runs are wrapped in a <c>w:hyperlink</c> whose relationship, on the part that holds the
/// paragraph, points at the URL. The runs keep their formatting.
/// </summary>
internal sealed class LinkHandler : IOperationHandler
{
    private static readonly string[] AllowedSchemes = { "http", "https", "mailto" };

    public bool CanHandle(PlanOperation operation) => operation is LinkOp { Target: TextSpanAnchor };

    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (LinkOp)operation;
        var anchor = (TextSpanAnchor)op.Target;

        if (op.Mode == ChangeMode.Tracked)
            return Fail(anchor, ValidationErrorCodes.InvalidOperation,
                "Word records no revision for adding a hyperlink to existing text, so mode 'Tracked' " +
                "cannot be honoured. Send this operation with mode 'Direct'.");

        if (!IsAllowedUrl(op.Url))
            return Fail(anchor, ValidationErrorCodes.InvalidOperation,
                $"'{op.Url}' is not an absolute http, https, or mailto URL. Other schemes, such as file: " +
                "or javascript:, and relative paths are refused: a link written into a document is one a reader will click.");

        if (string.IsNullOrEmpty(anchor.Expect))
            return Fail(anchor, ValidationErrorCodes.InvalidOperation,
                "link needs a non-empty 'expect' naming the text to turn into a link. Write the text first " +
                "with changeText or insert when it does not exist yet.");

        var paragraph = WordModel.ResolveParagraph(context, anchor.ParaId);
        if (paragraph is null)
            return Fail(anchor, ValidationErrorCodes.AnchorNotFound, $"No paragraph with id '{anchor.ParaId}'.");

        var text = WordModel.Text.GetLogicalText(paragraph);
        var comparison = WordModel.Comparison(caseSensitive: true);
        var occurrences = WordModel.Text.CountOccurrences(text, anchor.Expect, comparison);
        if (occurrences == 0)
            return Fail(anchor, ValidationErrorCodes.ExpectMismatch,
                $"Expected text '{anchor.Expect}' not found in paragraph '{anchor.ParaId}' (document drifted).");

        var start = WordModel.Text.IndexOfOccurrence(text, anchor.Expect, anchor.Occurrence, comparison);
        if (start < 0)
            return Fail(anchor, ValidationErrorCodes.AmbiguousAnchor,
                $"Occurrence {anchor.Occurrence} of '{anchor.Expect}' does not exist ({occurrences} found).");

        if (WhyNotLinkable(ChangeTextHandler.CoveredRuns(paragraph, start, anchor.Expect.Length)) is { } reason)
            return Fail(anchor, ValidationErrorCodes.InvalidOperation, $"'{anchor.Expect}' cannot be linked: {reason}.");

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "link",
            Before = anchor.Expect,
            After = $"{anchor.Expect} -> {op.Url}",
            Context = WordModel.Snippet(text, start, anchor.Expect.Length),
            BlastRadius = 1,
            Capability = Capability.Deterministic
        });
    }

    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (LinkOp)operation;
        var anchor = (TextSpanAnchor)op.Target;
        var paragraph = WordModel.ResolveParagraph(context, anchor.ParaId)
            ?? throw new InvalidOperationException($"Paragraph '{anchor.ParaId}' vanished before apply.");

        var text = WordModel.Text.GetLogicalText(paragraph);
        var start = WordModel.Text.IndexOfOccurrence(text, anchor.Expect, anchor.Occurrence, WordModel.Comparison(caseSensitive: true));
        if (start < 0) throw new InvalidOperationException($"Expected text '{anchor.Expect}' not found at apply time.");

        var covered = WordModel.Text.IsolateSpan(paragraph, start, anchor.Expect.Length);
        if (covered.Count == 0) throw new InvalidOperationException("Span isolation produced no runs.");

        var parent = covered[0].Parent ?? throw new InvalidOperationException("Run has no parent.");
        var part = (paragraph.Ancestors().OfType<OpenXmlPartRootElement>().FirstOrDefault()
                    ?? throw new InvalidOperationException("Paragraph is not inside a document part.")).OpenXmlPart
            ?? throw new InvalidOperationException("Paragraph's part is not loaded.");
        var relationship = part.AddHyperlinkRelationship(new Uri(op.Url, UriKind.Absolute), isExternal: true);

        // Everything between the first and last covered run moves with them - bookmark and
        // comment-range markers included - so the link covers exactly the matched text and the
        // markers keep their place relative to it.
        var children = parent.ChildElements.ToList();
        var first = children.IndexOf(covered[0]);
        var last = children.IndexOf(covered[covered.Count - 1]);
        var hyperlink = new Hyperlink { Id = relationship.Id, History = true };
        parent.InsertBefore(hyperlink, covered[0]);
        for (var i = first; i <= last; i++)
        {
            var child = children[i];
            child.Remove();
            hyperlink.AppendChild(child);
        }
    }

    /// <summary>
    /// Why the covered runs cannot be wrapped in one hyperlink, or null when they can. A
    /// hyperlink is paragraph-level content, so the runs must be direct children of the
    /// paragraph (or of an inline content control's content), and must not already sit in one.
    /// </summary>
    private static string? WhyNotLinkable(IReadOnlyList<OpenXmlElement> covered)
    {
        if (covered.Count == 0) return "it covers no text runs";
        if (covered.Any(run => run.Ancestors<Hyperlink>().Any()))
            return "it is already inside a hyperlink, and links are not nested";

        var parents = covered.Select(run => run.Parent).Distinct().ToList();
        if (parents.Count > 1)
            return "it spans more than one container (a tracked change, field, or content control boundary)";
        if (parents[0] is not (Paragraph or SdtContentRun))
            return $"it sits inside {WordRevisions.TagOf(parents[0]!) ?? parents[0]!.LocalName}, which cannot hold a hyperlink; " +
                   "resolve the pending revision or edit outside it first";

        var children = parents[0]!.ChildElements.ToList();
        var first = children.IndexOf(covered[0]);
        var last = children.IndexOf(covered[covered.Count - 1]);
        for (var i = first + 1; i < last; i++)
            if (children[i] is Hyperlink) return "it spans an existing hyperlink";
        return null;
    }

    private static bool IsAllowedUrl(string url) =>
        Uri.TryCreate(url, UriKind.Absolute, out var uri) &&
        AllowedSchemes.Contains(uri.Scheme, StringComparer.OrdinalIgnoreCase) &&
        (uri.Scheme.Equals("mailto", StringComparison.OrdinalIgnoreCase) || !string.IsNullOrEmpty(uri.Host));

    private static OperationPreview Fail(Anchor anchor, string code, string message) =>
        OperationPreview.Fail(new ValidationError(code, message, anchor));
}
