using DocumentFormat.OpenXml;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.Word;

/// <summary>
/// Sets a section's page geometry: paper size, orientation, and margins.
/// </summary>
/// <remarks>
/// <c>w:sectPr</c> is a strict sequence and shared with the header/footer writer, so every
/// write goes through <see cref="WordSections.Replace{T}"/> rather than appending. The
/// section a null target means is the body's - the whole document, until an
/// <c>insertBreak</c> splits it.
/// </remarks>
internal sealed class WordPageSetupHandler : IOperationHandler
{
    /// <summary>Paper sizes in twips, portrait, as Word writes them.</summary>
    private static readonly IReadOnlyDictionary<string, (int Width, int Height)> Papers =
        new Dictionary<string, (int, int)>(StringComparer.OrdinalIgnoreCase)
        {
            ["A3"] = (16838, 23811),
            ["A4"] = (11906, 16838),
            ["A5"] = (8391, 11906),
            ["Letter"] = (12240, 15840),
            ["Legal"] = (12240, 20160),
            ["Tabloid"] = (15840, 24480),
            ["Executive"] = (10440, 15120)
        };

    private static string PaperNames => string.Join(", ", Papers.Keys.OrderBy(k => k, StringComparer.Ordinal));

    public bool CanHandle(PlanOperation operation) => operation is PageSetupOp;

    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (PageSetupOp)operation;

        if (op.Target is not null and not TextSpanAnchor)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.UnsupportedOperation,
                "pageSetup takes no target - which sets the document's final section - or a " +
                "paragraph anchor naming the section to set.", op.Target));

        if (op.PaperSize is { Length: > 0 } paper)
        {
            if (!Papers.ContainsKey(paper))
                return OperationPreview.Fail(new ValidationError(
                    ValidationErrorCodes.InvalidOperation,
                    $"'{paper}' is not a paper size. Expected one of: {PaperNames}. " +
                    "For anything else give pageWidthTwips and pageHeightTwips.", op.Target));

            if (op.PageWidthTwips is not null || op.PageHeightTwips is not null)
                return OperationPreview.Fail(new ValidationError(
                    ValidationErrorCodes.InvalidOperation,
                    "paperSize and pageWidthTwips/pageHeightTwips both set the page size; name one.",
                    op.Target));
        }

        foreach (var (name, value) in Measurements(op))
            if (value is <= 0)
                return OperationPreview.Fail(new ValidationError(
                    ValidationErrorCodes.InvalidOperation,
                    $"{name} must be a positive number of twips; got {value}.", op.Target));

        if (!HasAnything(op))
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "pageSetup requires at least one of: paperSize, pageWidthTwips, pageHeightTwips, " +
                "orientation, marginTopTwips, marginBottomTwips, marginLeftTwips, marginRightTwips, " +
                "headerDistanceTwips, footerDistanceTwips, gutterTwips.", op.Target));

        if (op.Target is TextSpanAnchor anchor && WordModel.ResolveParagraph(context, anchor.ParaId) is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No paragraph with id '{anchor.ParaId}'.", anchor));

        var section = Section(context, op);
        var before = section.GetFirstChild<PageSize>() is { } size
            ? $"{size.Width?.Value ?? 0}×{size.Height?.Value ?? 0} twips"
            : "default page size";

        return OperationPreview.Ok(new ProposedChange
        {
            Target = op.Target,
            Verb = "pageSetup",
            Before = before,
            After = Describe(op),
            Context = op.Target is TextSpanAnchor ts ? $"section of paragraph '{ts.ParaId}'" : "final section",
            BlastRadius = 1,
            Capability = Capability.Deterministic
        });
    }

    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (PageSetupOp)operation;
        var section = Section(context, op);

        ApplyPageSize(section, op);
        ApplyMargins(section, op);
    }

    private static void ApplyPageSize(SectionProperties section, PageSetupOp op)
    {
        var existing = section.GetFirstChild<PageSize>();
        var orientation = op.Orientation ?? Current(existing);

        uint width = existing?.Width?.Value ?? (uint)Papers["Letter"].Width;
        uint height = existing?.Height?.Value ?? (uint)Papers["Letter"].Height;

        if (op.PaperSize is { Length: > 0 } paper)
        {
            var (w, h) = Papers[paper];
            (width, height) = ((uint)w, (uint)h);
        }

        if (op.PageWidthTwips is int explicitWidth) width = (uint)explicitWidth;
        if (op.PageHeightTwips is int explicitHeight) height = (uint)explicitHeight;

        // Orientation is the last word on which dimension is which: a caller asking for
        // landscape A4 gives a paper size in portrait and expects the pages turned.
        if (orientation == PageOrientation.Landscape && width < height) (width, height) = (height, width);
        if (orientation == PageOrientation.Portrait && width > height) (width, height) = (height, width);

        if (op.PaperSize is null && op.PageWidthTwips is null && op.PageHeightTwips is null &&
            op.Orientation is null && existing is null)
            return;

        WordSections.Replace(section, new PageSize
        {
            Width = width,
            Height = height,
            Orient = orientation == PageOrientation.Landscape
                ? PageOrientationValues.Landscape
                : PageOrientationValues.Portrait
        });
    }

    private static void ApplyMargins(SectionProperties section, PageSetupOp op)
    {
        if (op.MarginTopTwips is null && op.MarginBottomTwips is null &&
            op.MarginLeftTwips is null && op.MarginRightTwips is null &&
            op.HeaderDistanceTwips is null && op.FooterDistanceTwips is null &&
            op.GutterTwips is null)
            return;

        // Cloned, so setting one edge does not reset the three the caller did not mention.
        var existing = section.GetFirstChild<PageMargin>();
        var margin = existing is null ? DefaultMargin() : (PageMargin)existing.CloneNode(deep: true);

        if (op.MarginTopTwips is int top) margin.Top = top;
        if (op.MarginBottomTwips is int bottom) margin.Bottom = bottom;
        if (op.MarginLeftTwips is int left) margin.Left = (uint)left;
        if (op.MarginRightTwips is int right) margin.Right = (uint)right;
        if (op.HeaderDistanceTwips is int header) margin.Header = (uint)header;
        if (op.FooterDistanceTwips is int footer) margin.Footer = (uint)footer;
        if (op.GutterTwips is int gutter) margin.Gutter = (uint)gutter;

        WordSections.Replace(section, margin);
    }

    /// <summary>Word's own defaults - one inch all round - for a section that states none.</summary>
    private static PageMargin DefaultMargin() => new()
    {
        Top = 1440,
        Bottom = 1440,
        Left = 1440,
        Right = 1440,
        Header = 720,
        Footer = 720,
        Gutter = 0
    };

    private static PageOrientation Current(PageSize? size) =>
        size?.Orient?.Value == PageOrientationValues.Landscape ||
        (size?.Width?.Value is { } w && size.Height?.Value is { } h && w > h)
            ? PageOrientation.Landscape
            : PageOrientation.Portrait;

    /// <summary>
    /// The section the operation addresses: the one governing the anchored paragraph, or
    /// the body's when there is no target.
    /// </summary>
    private static SectionProperties Section(ApplyContext context, PageSetupOp op)
    {
        var main = WordModel.Main(context.Package);
        if (op.Target is not TextSpanAnchor anchor) return WordSections.Require(main);

        var paragraph = WordModel.ResolveParagraph(context, anchor.ParaId);
        if (paragraph is null) return WordSections.Require(main);

        // A paragraph belongs to the first section that ends at or after it, which is the
        // next paragraph-level w:sectPr in document order - or the body's if there is none.
        foreach (var candidate in FollowingParagraphs(paragraph))
            if (candidate.ParagraphProperties?.SectionProperties is { } section)
                return section;

        return WordSections.Require(main);
    }

    private static IEnumerable<Paragraph> FollowingParagraphs(Paragraph from)
    {
        yield return from;

        var body = from.Ancestors<Body>().FirstOrDefault();
        if (body is null) yield break;

        var seen = false;
        foreach (var paragraph in body.Descendants<Paragraph>())
        {
            if (!seen)
            {
                seen = ReferenceEquals(paragraph, from);
                continue;
            }
            yield return paragraph;
        }
    }

    private static IEnumerable<(string Name, int? Value)> Measurements(PageSetupOp op)
    {
        yield return (nameof(op.PageWidthTwips), op.PageWidthTwips);
        yield return (nameof(op.PageHeightTwips), op.PageHeightTwips);
        yield return (nameof(op.MarginTopTwips), op.MarginTopTwips);
        yield return (nameof(op.MarginBottomTwips), op.MarginBottomTwips);
        yield return (nameof(op.MarginLeftTwips), op.MarginLeftTwips);
        yield return (nameof(op.MarginRightTwips), op.MarginRightTwips);
        yield return (nameof(op.HeaderDistanceTwips), op.HeaderDistanceTwips);
        yield return (nameof(op.FooterDistanceTwips), op.FooterDistanceTwips);
    }

    private static bool HasAnything(PageSetupOp op) =>
        op.PaperSize is not null || op.Orientation is not null || op.GutterTwips is not null ||
        Measurements(op).Any(m => m.Value is not null);

    private static string Describe(PageSetupOp op)
    {
        var parts = new List<string>();
        if (op.PaperSize is { Length: > 0 } paper) parts.Add(paper);
        if (op.PageWidthTwips is int w) parts.Add($"width={w}");
        if (op.PageHeightTwips is int h) parts.Add($"height={h}");
        if (op.Orientation is { } orientation) parts.Add(orientation.ToString().ToLowerInvariant());
        if (op.MarginTopTwips is int t) parts.Add($"top={t}");
        if (op.MarginBottomTwips is int b) parts.Add($"bottom={b}");
        if (op.MarginLeftTwips is int l) parts.Add($"left={l}");
        if (op.MarginRightTwips is int r) parts.Add($"right={r}");
        if (op.HeaderDistanceTwips is int hd) parts.Add($"header={hd}");
        if (op.FooterDistanceTwips is int fd) parts.Add($"footer={fd}");
        if (op.GutterTwips is int g) parts.Add($"gutter={g}");
        return string.Join(", ", parts);
    }
}
