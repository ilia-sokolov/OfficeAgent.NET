using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeAgent.PowerPoint;

/// <summary>
/// Mints the empty but valid .pptx a new deck starts from.
/// </summary>
/// <remarks>
/// A presentation needs more scaffolding than a Word document to be openable at all: a
/// slide master, a layout the slide points at, and a theme the master points at. All
/// three are built here so a created deck opens in PowerPoint rather than being reported
/// as needing repair. The single slide carries one empty title shape, which inspection
/// addresses as <c>slide256/shape2/p0</c> - the anchor an initial plan targets, matching
/// the role <c>auto-0000</c> plays for a blank Word document.
/// </remarks>
internal static class PowerPointBlankDocument
{
    /// <summary>The slide id the one starting slide is given.</summary>
    internal const uint FirstSlideId = 256U;

    /// <summary>The paragraph id an initial plan targets in a freshly created deck.</summary>
    internal const string FirstAnchor = "slide256/shape2/p0";

    /// <summary>Returns a minimal, schema-valid presentation with one empty slide.</summary>
    public static byte[] Create()
    {
        using var buffer = new MemoryStream();
        using (var document = PresentationDocument.Create(buffer, PresentationDocumentType.Presentation))
        {
            var presentationPart = document.AddPresentationPart();
            presentationPart.Presentation = new Presentation();

            var masterPart = presentationPart.AddNewPart<SlideMasterPart>();
            var themePart = masterPart.AddNewPart<ThemePart>();

            // The master carries the title and body placeholders every layout inherits
            // from, so a slide that states only its text still lands where the template
            // says it should.
            masterPart.SlideMaster = new SlideMaster(
                new CommonSlideData(Tree(
                    SlideLayouts.Placeholder(2U, "Title Placeholder 1", PlaceholderValues.Title, null,
                        SlideLayouts.TitleBox, majorFont: true),
                    SlideLayouts.Placeholder(3U, "Text Placeholder 2", PlaceholderValues.Body, 1U,
                        SlideLayouts.BodyBox, majorFont: false))),
                DefaultColorMap(),
                new SlideLayoutIdList());

            themePart.Theme = DefaultTheme();

            // One layout per shape a deck actually needs. A generated slide picks one by
            // name; without them every new slide would be a bare text box with hand-placed
            // geometry rather than something the template governs.
            var layoutParts = new Dictionary<string, SlideLayoutPart>(StringComparer.Ordinal);
            uint layoutId = 2147483649U;
            foreach (var definition in SlideLayouts.All)
            {
                var layoutPart = masterPart.AddNewPart<SlideLayoutPart>();

                // The layout must point back at its master. Creating the layout under the
                // master only writes the master→layout relationship; without the reverse
                // one PowerPoint reports the package as corrupt and offers to repair it,
                // even though the schema validator is satisfied.
                layoutPart.AddPart(masterPart);
                layoutPart.SlideLayout = definition.Build();

                masterPart.SlideMaster.SlideLayoutIdList!.Append(new SlideLayoutId
                {
                    Id = layoutId++,
                    RelationshipId = masterPart.GetIdOfPart(layoutPart)
                });
                layoutParts[definition.Name] = layoutPart;
            }

            var slidePart = presentationPart.AddNewPart<SlidePart>();
            slidePart.Slide = new Slide(new CommonSlideData(Tree(TitleShape())));
            slidePart.AddPart(layoutParts[SlideLayouts.Title]);

            presentationPart.Presentation.Append(
                new SlideMasterIdList(new SlideMasterId
                {
                    Id = 2147483648U,
                    RelationshipId = presentationPart.GetIdOfPart(masterPart)
                }),
                new SlideIdList(new SlideId
                {
                    Id = FirstSlideId,
                    RelationshipId = presentationPart.GetIdOfPart(slidePart)
                }),
                // 16:9, the default PowerPoint has used since 2013.
                new SlideSize { Cx = 12192000, Cy = 6858000 },
                new NotesSize { Cx = 6858000, Cy = 9144000 });

            presentationPart.Presentation.Save();
        }

        return buffer.ToArray();
    }

    /// <summary>
    /// A shape tree. p:spTree needs both p:nvGrpSpPr and p:grpSpPr, in that order, before
    /// any shape; without the second the part does not validate.
    /// </summary>
    private static ShapeTree Tree(params OpenXmlElement[] shapes)
    {
        var tree = new ShapeTree(
            new NonVisualGroupShapeProperties(
                new NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                new NonVisualGroupShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new GroupShapeProperties());

        foreach (var shape in shapes) tree.Append(shape);
        return tree;
    }

    /// <summary>The empty title placeholder that gives a new deck its first anchor.</summary>
    private static Shape TitleShape() => new(
        new NonVisualShapeProperties(
            new NonVisualDrawingProperties { Id = 2U, Name = "Title 1" },
            new NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
            new ApplicationNonVisualDrawingProperties(
                new PlaceholderShape { Type = PlaceholderValues.Title })),
        new ShapeProperties(),
        // One empty paragraph: something for an initial plan to target, with no text to
        // have to remove first.
        new TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph()));

    private static ColorMap DefaultColorMap() => new()
    {
        Background1 = A.ColorSchemeIndexValues.Light1,
        Text1 = A.ColorSchemeIndexValues.Dark1,
        Background2 = A.ColorSchemeIndexValues.Light2,
        Text2 = A.ColorSchemeIndexValues.Dark2,
        Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
        FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink,
        Accent1 = A.ColorSchemeIndexValues.Accent1,
        Accent2 = A.ColorSchemeIndexValues.Accent2,
        Accent3 = A.ColorSchemeIndexValues.Accent3,
        Accent4 = A.ColorSchemeIndexValues.Accent4,
        Accent5 = A.ColorSchemeIndexValues.Accent5,
        Accent6 = A.ColorSchemeIndexValues.Accent6
    };

    /// <summary>The Office theme, which a master must reference for the deck to open.</summary>
    private static A.Theme DefaultTheme() => new(
        new A.ThemeElements(
            new A.ColorScheme(
                new A.Dark1Color(new A.SystemColor { Val = A.SystemColorValues.WindowText }),
                new A.Light1Color(new A.SystemColor { Val = A.SystemColorValues.Window }),
                new A.Dark2Color(new A.RgbColorModelHex { Val = "44546A" }),
                new A.Light2Color(new A.RgbColorModelHex { Val = "E7E6E6" }),
                new A.Accent1Color(new A.RgbColorModelHex { Val = "4472C4" }),
                new A.Accent2Color(new A.RgbColorModelHex { Val = "ED7D31" }),
                new A.Accent3Color(new A.RgbColorModelHex { Val = "A5A5A5" }),
                new A.Accent4Color(new A.RgbColorModelHex { Val = "FFC000" }),
                new A.Accent5Color(new A.RgbColorModelHex { Val = "5B9BD5" }),
                new A.Accent6Color(new A.RgbColorModelHex { Val = "70AD47" }),
                new A.Hyperlink(new A.RgbColorModelHex { Val = "0563C1" }),
                new A.FollowedHyperlinkColor(new A.RgbColorModelHex { Val = "954F72" }))
            { Name = "Office" },
            new A.FontScheme(
                new A.MajorFont(
                    new A.LatinFont { Typeface = "Calibri Light" },
                    new A.EastAsianFont { Typeface = string.Empty },
                    new A.ComplexScriptFont { Typeface = string.Empty }),
                new A.MinorFont(
                    new A.LatinFont { Typeface = "Calibri" },
                    new A.EastAsianFont { Typeface = string.Empty },
                    new A.ComplexScriptFont { Typeface = string.Empty }))
            { Name = "Office" },
            new A.FormatScheme(
                new A.FillStyleList(
                    Solid(), Solid(), Solid()),
                new A.LineStyleList(
                    new A.Outline(Solid()) { Width = 6350 },
                    new A.Outline(Solid()) { Width = 12700 },
                    new A.Outline(Solid()) { Width = 19050 }),
                new A.EffectStyleList(
                    new A.EffectStyle(new A.EffectList()),
                    new A.EffectStyle(new A.EffectList()),
                    new A.EffectStyle(new A.EffectList())),
                new A.BackgroundFillStyleList(
                    Solid(), Solid(), Solid()))
            { Name = "Office" }))
    { Name = "Office Theme" };

    private static A.SolidFill Solid() =>
        new(new A.SchemeColor { Val = A.SchemeColorValues.PhColor });
}
