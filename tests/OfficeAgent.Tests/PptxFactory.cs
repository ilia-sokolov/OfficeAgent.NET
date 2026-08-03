using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeAgent.Tests;

/// <summary>Builds small, valid .pptx fixtures in memory for the PowerPoint tests.</summary>
internal static class PptxFactory
{
    public const string TitleText = "Quarterly Review";
    public const string BulletText = "Acme Corp revenue grew in every region except Acme Corp EMEA.";
    public const string SecondSlideTitle = "Outlook";
    public const string NotesText = "Remember to mention the Acme Corp merger.";

    /// <summary>
    /// A two-slide deck: a title plus a bullet on slide 1 (with speaker notes), and a
    /// title on slide 2. Slide 1's bullet deliberately repeats "Acme Corp" so occurrence
    /// handling is exercised, and its text is split across runs so run-spanning
    /// replacement is too.
    /// </summary>
    public static byte[] Deck()
    {
        using var stream = new MemoryStream();
        using (var document = PresentationDocument.Create(stream, PresentationDocumentType.Presentation))
        {
            var presentationPart = document.AddPresentationPart();
            presentationPart.Presentation = new Presentation();

            var (masterPart, layoutPart) = AddMasterAndLayout(presentationPart);

            var slide1 = presentationPart.AddNewPart<SlidePart>();
            slide1.Slide = new Slide(new CommonSlideData(Tree(
                TitleShape(2, TitleText),
                // Split across three runs: "Acme Corp" spans a run boundary in the
                // middle, which is what makes IsolateSpan worth testing.
                BodyShape(3,
                    "Acme ", "Corp revenue grew in every region except ", "Acme Corp EMEA."))));

            var notes = slide1.AddNewPart<NotesSlidePart>();
            notes.NotesSlide = new NotesSlide(new CommonSlideData(Tree(
                BodyShape(2, NotesText))));

            var slide2 = presentationPart.AddNewPart<SlidePart>();
            slide2.Slide = new Slide(new CommonSlideData(Tree(
                TitleShape(2, SecondSlideTitle))));

            foreach (var part in new[] { slide1, slide2 })
                part.AddPart(layoutPart);

            presentationPart.Presentation.Append(
                new SlideMasterIdList(new SlideMasterId
                {
                    Id = 2147483648U,
                    RelationshipId = presentationPart.GetIdOfPart(masterPart)
                }),
                new SlideIdList(
                    new SlideId { Id = 256U, RelationshipId = presentationPart.GetIdOfPart(slide1) },
                    new SlideId { Id = 257U, RelationshipId = presentationPart.GetIdOfPart(slide2) }),
                new SlideSize { Cx = 12192000, Cy = 6858000 },
                new NotesSize { Cx = 6858000, Cy = 9144000 });

            presentationPart.Presentation.Save();
        }

        return stream.ToArray();
    }

    /// <summary>
    /// A one-slide deck whose body shape splits a sentence across three runs, the middle
    /// one bold. Replacing text that starts inside the bold run is what proves character
    /// formatting on partially covered runs survives.
    /// </summary>
    public static byte[] DeckWithBoldRun()
    {
        using var stream = new MemoryStream();
        using (var document = PresentationDocument.Create(stream, PresentationDocumentType.Presentation))
        {
            var presentationPart = document.AddPresentationPart();
            presentationPart.Presentation = new Presentation();
            var (masterPart, layoutPart) = AddMasterAndLayout(presentationPart);

            var body = new Shape(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties { Id = 3U, Name = "Body 3" },
                    new NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties()),
                new ShapeProperties(),
                new TextBody(new A.BodyProperties(), new A.ListStyle(),
                    new A.Paragraph(
                        new A.Run(new A.RunProperties { Language = "en-US" }, new A.Text("plain ")),
                        new A.Run(new A.RunProperties { Language = "en-US", Bold = true }, new A.Text("BOLD text")),
                        new A.Run(new A.RunProperties { Language = "en-US" }, new A.Text(" tail")))));

            var slide = presentationPart.AddNewPart<SlidePart>();
            slide.Slide = new Slide(new CommonSlideData(Tree(TitleShape(2, TitleText), body)));
            slide.AddPart(layoutPart);

            presentationPart.Presentation.Append(
                new SlideMasterIdList(new SlideMasterId
                {
                    Id = 2147483648U,
                    RelationshipId = presentationPart.GetIdOfPart(masterPart)
                }),
                new SlideIdList(new SlideId { Id = 256U, RelationshipId = presentationPart.GetIdOfPart(slide) }),
                new SlideSize { Cx = 12192000, Cy = 6858000 },
                new NotesSize { Cx = 6858000, Cy = 9144000 });
            presentationPart.Presentation.Save();
        }

        return stream.ToArray();
    }

    /// <summary>A one-slide deck whose only shape holds a table, for table tests.</summary>
    public static byte[] DeckWithTable()
    {
        using var stream = new MemoryStream();
        using (var document = PresentationDocument.Create(stream, PresentationDocumentType.Presentation))
        {
            var presentationPart = document.AddPresentationPart();
            presentationPart.Presentation = new Presentation();
            var (masterPart, layoutPart) = AddMasterAndLayout(presentationPart);

            var slide = presentationPart.AddNewPart<SlidePart>();
            slide.Slide = new Slide(new CommonSlideData(Tree(
                TitleShape(2, TitleText),
                TableFrame(3))));
            slide.AddPart(layoutPart);

            presentationPart.Presentation.Append(
                new SlideMasterIdList(new SlideMasterId
                {
                    Id = 2147483648U,
                    RelationshipId = presentationPart.GetIdOfPart(masterPart)
                }),
                new SlideIdList(new SlideId { Id = 256U, RelationshipId = presentationPart.GetIdOfPart(slide) }),
                new SlideSize { Cx = 12192000, Cy = 6858000 },
                new NotesSize { Cx = 6858000, Cy = 9144000 });
            presentationPart.Presentation.Save();
        }

        return stream.ToArray();
    }

    private static (SlideMasterPart Master, SlideLayoutPart Layout) AddMasterAndLayout(
        PresentationPart presentationPart)
    {
        var masterPart = presentationPart.AddNewPart<SlideMasterPart>();
        var layoutPart = masterPart.AddNewPart<SlideLayoutPart>();

        // Creating the layout under the master writes only master→layout. PowerPoint also
        // requires the reverse relationship, and refuses the whole package without it -
        // a fixture missing this would be corrupt in a way no schema check reveals.
        layoutPart.AddPart(masterPart);

        layoutPart.SlideLayout = new SlideLayout(
            new CommonSlideData(Tree()),
            new ColorMapOverride(new A.MasterColorMapping()))
        { Type = SlideLayoutValues.Title };

        masterPart.SlideMaster = new SlideMaster(
            new CommonSlideData(Tree()),
            new ColorMap
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
            },
            new SlideLayoutIdList(new SlideLayoutId
            {
                Id = 2147483649U,
                RelationshipId = masterPart.GetIdOfPart(layoutPart)
            }));

        var themePart = masterPart.AddNewPart<ThemePart>();
        themePart.Theme = MinimalTheme();

        return (masterPart, layoutPart);
    }

    /// <summary>
    /// Builds a shape tree. p:spTree requires both p:nvGrpSpPr and p:grpSpPr, in that
    /// order, before any shape - omitting the second makes the part fail validation.
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

    private static Shape TitleShape(uint id, string text) =>
        new(new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = id, Name = $"Title {id}" },
                new NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(
                    new PlaceholderShape { Type = PlaceholderValues.Title })),
            new ShapeProperties(),
            new TextBody(new A.BodyProperties(), new A.ListStyle(), Paragraph(text)));

    private static Shape BodyShape(uint id, params string[] runTexts) =>
        new(new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = id, Name = $"Body {id}" },
                new NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties()),
            new ShapeProperties(),
            new TextBody(new A.BodyProperties(), new A.ListStyle(), Paragraph(runTexts)));

    private static GraphicFrame TableFrame(uint id)
    {
        var table = new A.Table(
            new A.TableProperties { FirstRow = true, BandRow = true },
            // Deliberately unequal, so a column edit that desynchronises the grid from the
            // cells is observable rather than hidden behind identical widths.
            new A.TableGrid(
                new A.GridColumn { Width = 3000000L },
                new A.GridColumn { Width = 1000000L }),
            Row("Region", "Revenue"),
            Row("EMEA", "41850"));

        return new GraphicFrame(
            new NonVisualGraphicFrameProperties(
                new NonVisualDrawingProperties { Id = id, Name = $"Table {id}" },
                new NonVisualGraphicFrameDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new Transform(
                new A.Offset { X = 0L, Y = 0L },
                new A.Extents { Cx = 6000000L, Cy = 1000000L }),
            new A.Graphic(new A.GraphicData(table)
            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/table" }));
    }

    private static A.TableRow Row(params string[] cells)
    {
        var row = new A.TableRow { Height = 370840L };
        foreach (var cell in cells)
            row.Append(new A.TableCell(
                new A.TextBody(new A.BodyProperties(), new A.ListStyle(), Paragraph(cell)),
                new A.TableCellProperties()));
        return row;
    }

    private static A.Paragraph Paragraph(params string[] runTexts)
    {
        var paragraph = new A.Paragraph();
        foreach (var text in runTexts)
            paragraph.Append(new A.Run(
                new A.RunProperties { Language = "en-US", Dirty = false },
                new A.Text(text)));
        return paragraph;
    }

    /// <summary>The smallest theme a deck needs to satisfy schema validation.</summary>
    private static A.Theme MinimalTheme() => new(
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
                new A.MajorFont(new A.LatinFont { Typeface = "Calibri Light" },
                    new A.EastAsianFont { Typeface = string.Empty },
                    new A.ComplexScriptFont { Typeface = string.Empty }),
                new A.MinorFont(new A.LatinFont { Typeface = "Calibri" },
                    new A.EastAsianFont { Typeface = string.Empty },
                    new A.ComplexScriptFont { Typeface = string.Empty }))
            { Name = "Office" },
            new A.FormatScheme(
                new A.FillStyleList(
                    new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                    new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                    new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })),
                new A.LineStyleList(
                    new A.Outline(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })) { Width = 6350 },
                    new A.Outline(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })) { Width = 12700 },
                    new A.Outline(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })) { Width = 19050 }),
                new A.EffectStyleList(
                    new A.EffectStyle(new A.EffectList()),
                    new A.EffectStyle(new A.EffectList()),
                    new A.EffectStyle(new A.EffectList())),
                new A.BackgroundFillStyleList(
                    new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                    new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                    new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })))
            { Name = "Office" }))
    { Name = "Office Theme" };
}
