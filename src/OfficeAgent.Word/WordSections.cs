using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.Word;

/// <summary>
/// The section: page size, and the headers and footers hung off it.
/// </summary>
/// <remarks>
/// <c>w:sectPr</c> is a sequence, and a strict one - the header and footer references open
/// it, and <c>w:titlePg</c> comes near the end, after the column and alignment settings.
/// Appending any of them, which is the obvious thing to do, produces a document Word offers
/// to repair. Every write here goes through <see cref="Place"/> for that reason.
/// </remarks>
internal static class WordSections
{
    /// <summary>A twip is 1/1440 inch; an EMU is 1/914400. Word measures pages in twips.</summary>
    public const long EmuPerTwip = 635L;

    /// <summary>US Letter, which is what a section with no explicit page size means.</summary>
    private const long DefaultPageWidthTwips = 12240L;
    private const long DefaultPageHeightTwips = 15840L;

    /// <summary>
    /// The order <c>CT_SectPr</c> declares its children in. Anything not listed sorts after
    /// everything listed, which is where the printer settings and the like belong anyway.
    /// </summary>
    private static readonly Type[] Order =
    {
        typeof(HeaderReference), typeof(FooterReference), typeof(FootnoteProperties),
        typeof(EndnoteProperties), typeof(SectionType), typeof(PageSize), typeof(PageMargin),
        typeof(PaperSource), typeof(PageBorders), typeof(LineNumberType), typeof(PageNumberType),
        typeof(Columns), typeof(FormProtection), typeof(VerticalTextAlignmentOnPage),
        typeof(NoEndnote), typeof(TitlePage), typeof(TextDirection), typeof(BiDi),
        typeof(GutterOnRight), typeof(DocGrid)
    };

    /// <summary>
    /// The body's section properties, created when the document has none. <c>w:sectPr</c> is
    /// the last child of <c>w:body</c> - after every paragraph, not before.
    /// </summary>
    public static SectionProperties Require(MainDocumentPart main)
    {
        var body = main.Document.Body
            ?? throw new InvalidOperationException("Document has no body.");

        var existing = body.GetFirstChild<SectionProperties>();
        if (existing is not null) return existing;

        var properties = new SectionProperties();
        body.AppendChild(properties);
        return properties;
    }

    /// <summary>The page's size in EMUs, falling back to Letter when the section is silent.</summary>
    public static (long Width, long Height) PageSizeEmu(SectionProperties section)
    {
        var size = section.GetFirstChild<PageSize>();
        var width = size?.Width?.Value ?? (uint)DefaultPageWidthTwips;
        var height = size?.Height?.Value ?? (uint)DefaultPageHeightTwips;

        // A landscape section states its orientation as well as swapped dimensions, but not
        // every producer swaps them, so trust the orientation when it disagrees.
        if (size?.Orient?.Value == PageOrientationValues.Landscape && width < height)
            (width, height) = (height, width);

        return (width * EmuPerTwip, height * EmuPerTwip);
    }

    /// <summary>Inserts a child of <c>w:sectPr</c> at the position the schema requires.</summary>
    public static void Place(SectionProperties section, OpenXmlElement child)
    {
        var rank = RankOf(child);

        OpenXmlElement? previous = null;
        foreach (var existing in section.ChildElements)
        {
            if (RankOf(existing) > rank) break;
            previous = existing;
        }

        if (previous is null) section.InsertAt(child, 0);
        else section.InsertAfter(child, previous);
    }

    /// <summary>Replaces a single-valued section setting, or adds it in schema order.</summary>
    public static void Replace<T>(SectionProperties section, T child) where T : OpenXmlElement
    {
        section.GetFirstChild<T>()?.Remove();
        Place(section, child);
    }

    /// <summary>
    /// The header of the given kind, created and referenced when the section has none.
    /// </summary>
    public static Header HeaderFor(MainDocumentPart main, SectionProperties section, HeaderFooterValues kind)
    {
        var reference = section.Elements<HeaderReference>()
            .FirstOrDefault(r => r.Type is not null && r.Type.Value == kind);

        if (reference?.Id?.Value is { Length: > 0 } id &&
            main.GetPartById(id) is HeaderPart existing)
            return existing.Header ??= new Header();

        var part = main.AddNewPart<HeaderPart>();
        part.Header = new Header();

        reference?.Remove();
        Place(section, new HeaderReference { Type = kind, Id = main.GetIdOfPart(part) });
        return part.Header;
    }

    /// <summary>The footer of the given kind, created and referenced when there is none.</summary>
    public static Footer FooterFor(MainDocumentPart main, SectionProperties section, HeaderFooterValues kind)
    {
        var reference = section.Elements<FooterReference>()
            .FirstOrDefault(r => r.Type is not null && r.Type.Value == kind);

        if (reference?.Id?.Value is { Length: > 0 } id &&
            main.GetPartById(id) is FooterPart existing)
            return existing.Footer ??= new Footer();

        var part = main.AddNewPart<FooterPart>();
        part.Footer = new Footer();

        reference?.Remove();
        Place(section, new FooterReference { Type = kind, Id = main.GetIdOfPart(part) });
        return part.Footer;
    }

    /// <summary>Whether the first page uses headers and footers of its own.</summary>
    public static void SetTitlePage(SectionProperties section, bool enabled)
    {
        section.GetFirstChild<TitlePage>()?.Remove();
        if (enabled) Place(section, new TitlePage());
    }

    public static bool HasTitlePage(SectionProperties section)
    {
        var page = section.GetFirstChild<TitlePage>();
        // w:titlePg with no w:val means on; with one, the value decides.
        return page is not null && (page.Val is null || page.Val.Value);
    }

    private static int RankOf(OpenXmlElement element)
    {
        var index = Array.IndexOf(Order, element.GetType());
        return index < 0 ? Order.Length : index;
    }
}
