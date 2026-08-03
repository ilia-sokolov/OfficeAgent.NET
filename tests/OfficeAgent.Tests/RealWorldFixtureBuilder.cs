using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.Tests;

/// <summary>
/// Writes documents with the length, structure and language of real ones - a signed
/// statement of work and a quarterly business review - so the MCP server can be driven
/// against them the way a client such as Claude Code would, and the result judged by
/// opening it in Office.
/// </summary>
public class RealWorldFixtureBuilder
{
    public static string Workspace =>
        Environment.GetEnvironmentVariable("OFFICEAGENT_REALWORLD_ROOT")
        ?? Path.Combine(Path.GetTempPath(), "officeagent-realworld");

    [Fact]
    [Trait("Category", "Fixture")]
    public void Build_the_real_world_workspace()
    {
        Directory.CreateDirectory(Workspace);
        File.WriteAllBytes(Path.Combine(Workspace, "statement-of-work.docx"), StatementOfWork());
        File.WriteAllBytes(Path.Combine(Workspace, "qbr-fy26q3.pptx"), QuarterlyReview());
        File.WriteAllBytes(Path.Combine(Workspace, "_blank.pptx"),
            new OfficeAgent.PowerPoint.PowerPointModule().CreateBlank());

        Assert.True(File.Exists(Path.Combine(Workspace, "statement-of-work.docx")));
        Assert.True(File.Exists(Path.Combine(Workspace, "qbr-fy26q3.pptx")));
    }

    /// <summary>
    /// A consulting statement of work: parties, numbered sections in the language such
    /// documents actually use, a rate card, a milestone schedule, and a signature block.
    /// </summary>
    public static byte[] StatementOfWork()
    {
        using var buffer = new MemoryStream();
        using (var document = WordprocessingDocument.Create(buffer, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            var body = new Body();

            body.Append(H("Statement of Work", "Heading1"));
            body.Append(P("SOW-2026-0148 under the Master Services Agreement dated 4 January 2026."));
            body.Append(P("This Statement of Work is made between Northwind Traders Limited (the \"Client\") " +
                          "and Contoso Consulting BV (the \"Supplier\")."));

            body.Append(H("1. Background", "Heading2"));
            body.Append(P("The Client operates a billing platform that has reached the limits of its current " +
                          "architecture. The Supplier has been engaged to migrate the platform to a " +
                          "service-oriented design without interrupting month-end invoicing."));

            body.Append(H("2. Scope of Work", "Heading2"));
            body.Append(P("The Supplier shall deliver the following: discovery and architecture assessment; " +
                          "migration of the rating engine; migration of the invoicing service; and a period " +
                          "of hypercare following go-live."));
            body.Append(P("Anything not expressly listed in this section is out of scope. Changes to scope " +
                          "require a written change request signed by both parties."));

            body.Append(H("3. Fees and Rate Card", "Heading2"));
            body.Append(P("Services are provided on a time-and-materials basis at the following daily rates, " +
                          "exclusive of VAT."));
            body.Append(RateCard());
            body.Append(P(string.Empty));
            body.Append(P("The Supplier shall invoice monthly in arrears. Invoices are payable within 30 days " +
                          "of receipt. Late payment accrues interest at 8% above base rate."));

            body.Append(H("4. Milestones", "Heading2"));
            body.Append(MilestoneTable());
            body.Append(P(string.Empty));

            body.Append(H("5. Liability", "Heading2"));
            body.Append(P("The Supplier's total aggregate liability under this Statement of Work shall not " +
                          "exceed the total fees paid in the twelve months preceding the claim."));

            body.Append(H("6. Termination", "Heading2"));
            body.Append(P("Either party may terminate this Statement of Work on 60 days written notice. " +
                          "The Client shall pay for all work performed up to the effective date of termination."));

            body.Append(H("Signatures", "Heading2"));
            var signature = P("Signed for and on behalf of the Client: ");
            signature.AppendChild(new SdtRun(
                new SdtProperties(new W.Tag { Val = "ClientSignatory" }, new SdtId { Val = 4001 }),
                new SdtContentRun(new Run(new W.Text("[NAME]")))));
            body.Append(signature);
            body.Append(P("Signed for and on behalf of Contoso Consulting BV: ______________________"));

            main.Document = new Document(body);

            var styles = main.AddNewPart<StyleDefinitionsPart>();
            styles.Styles = new Styles(S("Heading1", "heading 1"), S("Heading2", "heading 2"), S("Normal", "Normal"));
            styles.Styles.Save();
            main.Document.Save();
        }

        return buffer.ToArray();
    }

    /// <summary>
    /// A quarterly business review deck with the slides such a deck really has, and
    /// speaker notes on the three that a presenter would annotate.
    /// </summary>
    public static byte[] QuarterlyReview()
    {
        using var buffer = new MemoryStream();
        using (var document = PresentationDocument.Create(buffer, PresentationDocumentType.Presentation))
        {
            var presentationPart = document.AddPresentationPart();
            presentationPart.Presentation = new Presentation();

            var masterPart = presentationPart.AddNewPart<SlideMasterPart>();
            var layoutPart = masterPart.AddNewPart<SlideLayoutPart>();
            var themePart = masterPart.AddNewPart<ThemePart>();
            layoutPart.AddPart(masterPart);

            layoutPart.SlideLayout = new SlideLayout(
                new CommonSlideData(Tree()),
                new ColorMapOverride(new A.MasterColorMapping())) { Type = SlideLayoutValues.Title };
            masterPart.SlideMaster = new SlideMaster(
                new CommonSlideData(Tree()), Mapping(),
                new SlideLayoutIdList(new SlideLayoutId
                {
                    Id = 2147483649U,
                    RelationshipId = masterPart.GetIdOfPart(layoutPart)
                }));
            themePart.Theme = Theme();

            var slides = new List<SlidePart>
            {
                Slide(presentationPart, layoutPart,
                    Title(2, "FY26 Q3 Business Review"),
                    Body(3, "Northwind Traders  |  Billing Platform Programme")),

                Slide(presentationPart, layoutPart,
                    Title(2, "Agenda"),
                    Body(3, "Financial performance against plan"),
                    Body(4, "Regional revenue and pipeline"),
                    Body(5, "Delivery risks and mitigations"),
                    Body(6, "Asks of the board")),

                Slide(presentationPart, layoutPart,
                    Title(2, "Revenue by Region"),
                    RegionTable(3)),

                Slide(presentationPart, layoutPart,
                    Title(2, "Delivery Risks"),
                    Body(3, "Rating engine migration is four weeks behind plan."),
                    Body(4, "Two senior engineers roll off at the end of the quarter."),
                    Body(5, "The Contoso Consulting BV renewal is still unsigned.")),

                Slide(presentationPart, layoutPart,
                    Title(2, "Asks of the Board"),
                    Body(3, "Approve the additional hypercare budget of EUR 85,000."),
                    Body(4, "Confirm the go-live date of 31 March 2026."))
            };

            Notes(slides[0], "Thank the board for the extra time. Keep the opening to two minutes.");
            Notes(slides[2], "If asked about APAC: the decline is one lost account, not a trend.");
            Notes(slides[3], "Do not commit to a recovery date here. Say we will come back in two weeks.");

            var ids = new SlideIdList();
            uint id = 256;
            foreach (var slide in slides)
                ids.Append(new SlideId { Id = id++, RelationshipId = presentationPart.GetIdOfPart(slide) });

            presentationPart.Presentation.Append(
                new SlideMasterIdList(new SlideMasterId
                {
                    Id = 2147483648U,
                    RelationshipId = presentationPart.GetIdOfPart(masterPart)
                }),
                ids,
                new SlideSize { Cx = 12192000, Cy = 6858000 },
                new NotesSize { Cx = 6858000, Cy = 9144000 });
            presentationPart.Presentation.Save();
        }

        return buffer.ToArray();
    }

    // ── Word building blocks ──────────────────────────────────────────────────

    private static Paragraph H(string text, string styleId) =>
        new(new ParagraphProperties(new ParagraphStyleId { Val = styleId }),
            new Run(new W.Text(text)));

    private static Paragraph P(string text) =>
        new(new Run(new W.Text(text) { Space = SpaceProcessingModeValues.Preserve }));

    private static Table RateCard() => WordTable(
        new[] { "4200", "2600", "2600" },
        new[] { "Role", "Daily rate (EUR)", "Allocation" },
        new[] { "Engagement lead", "1,450", "0.2 FTE" },
        new[] { "Solution architect", "1,250", "1.0 FTE" },
        new[] { "Senior engineer", "1,050", "2.0 FTE" });

    private static Table MilestoneTable() => WordTable(
        new[] { "4200", "2600", "2600" },
        new[] { "Milestone", "Fee (EUR)", "Due" },
        new[] { "Architecture assessment signed off", "38,000", "On signature" },
        new[] { "Rating engine migrated", "96,000", "Net 30" },
        new[] { "Invoicing service migrated", "112,000", "Net 30" });

    private static Table WordTable(string[] widths, params string[][] rows)
    {
        // w:tblBorders is a sequence, and w:tblGrid must precede the rows.
        var table = new Table(new TableProperties(
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4 },
                new LeftBorder { Val = BorderValues.Single, Size = 4 },
                new BottomBorder { Val = BorderValues.Single, Size = 4 },
                new RightBorder { Val = BorderValues.Single, Size = 4 },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 })));

        var grid = new TableGrid();
        foreach (var width in widths) grid.Append(new GridColumn { Width = width });
        table.Append(grid);

        foreach (var row in rows)
        {
            var tr = new TableRow();
            foreach (var cell in row)
                tr.Append(new TableCell(
                    new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Auto }),
                    new Paragraph(new Run(new W.Text(cell)))));
            table.Append(tr);
        }
        return table;
    }

    private static Style S(string id, string name) =>
        new(new StyleName { Val = name }) { StyleId = id, Type = StyleValues.Paragraph };

    // ── PowerPoint building blocks ────────────────────────────────────────────

    private static void Notes(SlidePart slide, string text)
    {
        var notes = slide.AddNewPart<NotesSlidePart>();
        notes.NotesSlide = new NotesSlide(new CommonSlideData(Tree(Body(2, text))));
    }

    private static SlidePart Slide(
        PresentationPart presentationPart, SlideLayoutPart layoutPart, params OpenXmlElement[] shapes)
    {
        var part = presentationPart.AddNewPart<SlidePart>();
        part.Slide = new Slide(new CommonSlideData(Tree(shapes)));
        part.AddPart(layoutPart);
        return part;
    }

    private static ShapeTree Tree(params OpenXmlElement[] shapes)
    {
        var tree = new ShapeTree(
            new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                new P.NonVisualGroupShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.GroupShapeProperties());
        foreach (var shape in shapes) tree.Append(shape);
        return tree;
    }

    private static Shape Title(uint id, string text) => new(
        new P.NonVisualShapeProperties(
            new P.NonVisualDrawingProperties { Id = id, Name = $"Title {id}" },
            new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
            new P.ApplicationNonVisualDrawingProperties(
                new PlaceholderShape { Type = PlaceholderValues.Title })),
        new P.ShapeProperties(),
        new P.TextBody(new A.BodyProperties(), new A.ListStyle(), Para(text)));

    private static Shape Body(uint id, string text) => new(
        new P.NonVisualShapeProperties(
            new P.NonVisualDrawingProperties { Id = id, Name = $"Body {id}" },
            new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
            new P.ApplicationNonVisualDrawingProperties()),
        new P.ShapeProperties(),
        new P.TextBody(new A.BodyProperties(), new A.ListStyle(), Para(text)));

    private static A.Paragraph Para(string text) =>
        new(new A.Run(new A.RunProperties { Language = "en-GB" }, new A.Text(text)));

    private static GraphicFrame RegionTable(uint id)
    {
        var table = new A.Table(
            new A.TableProperties { FirstRow = true, BandRow = true },
            new A.TableGrid(
                new A.GridColumn { Width = 2600000L },
                new A.GridColumn { Width = 2000000L },
                new A.GridColumn { Width = 2000000L },
                new A.GridColumn { Width = 1800000L }),
            Tr("Region", "FY26 Q2", "FY26 Q3", "Variance"),
            Tr("EMEA", "41,850", "44,120", "+5.4%"),
            Tr("Americas", "62,400", "67,900", "+8.8%"),
            Tr("APAC", "22,100", "19,750", "-10.6%"));

        return new GraphicFrame(
            new P.NonVisualGraphicFrameProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = $"Table {id}" },
                new P.NonVisualGraphicFrameDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.Transform(
                new A.Offset { X = 838200L, Y = 2000000L },
                new A.Extents { Cx = 8400000L, Cy = 1500000L }),
            new A.Graphic(new A.GraphicData(table)
            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/table" }));
    }

    private static A.TableRow Tr(params string[] cells)
    {
        var row = new A.TableRow { Height = 370840L };
        foreach (var cell in cells)
            row.Append(new A.TableCell(
                new A.TextBody(new A.BodyProperties(), new A.ListStyle(), Para(cell)),
                new A.TableCellProperties()));
        return row;
    }

    private static ColorMap Mapping() => new()
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

    private static A.Theme Theme() => new(
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
                new A.FollowedHyperlinkColor(new A.RgbColorModelHex { Val = "954F72" })) { Name = "Office" },
            new A.FontScheme(
                new A.MajorFont(new A.LatinFont { Typeface = "Calibri Light" },
                    new A.EastAsianFont { Typeface = string.Empty },
                    new A.ComplexScriptFont { Typeface = string.Empty }),
                new A.MinorFont(new A.LatinFont { Typeface = "Calibri" },
                    new A.EastAsianFont { Typeface = string.Empty },
                    new A.ComplexScriptFont { Typeface = string.Empty })) { Name = "Office" },
            new A.FormatScheme(
                new A.FillStyleList(Fill(), Fill(), Fill()),
                new A.LineStyleList(
                    new A.Outline(Fill()) { Width = 6350 },
                    new A.Outline(Fill()) { Width = 12700 },
                    new A.Outline(Fill()) { Width = 19050 }),
                new A.EffectStyleList(
                    new A.EffectStyle(new A.EffectList()),
                    new A.EffectStyle(new A.EffectList()),
                    new A.EffectStyle(new A.EffectList())),
                new A.BackgroundFillStyleList(Fill(), Fill(), Fill())) { Name = "Office" }))
    { Name = "Office Theme" };

    private static A.SolidFill Fill() => new(new A.SchemeColor { Val = A.SchemeColorValues.PhColor });
}
