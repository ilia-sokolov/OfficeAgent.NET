using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.Samples.RequirementsReview;

/// <summary>Variations of the sample method statement used by the tests.</summary>
public sealed record MethodStatementOptions
{
    /// <summary>Gets the installation-window line.</summary>
    public string InstallationWindow { get; init; } = MethodStatementFixture.InstallationWindowText;

    /// <summary>Gets the first recovery paragraph.</summary>
    public string Recovery { get; init; } = MethodStatementFixture.RecoveryText;

    /// <summary>Gets a value indicating whether the recovery paragraph carries a pending tracked insertion.</summary>
    public bool PendingRevisionInRecovery { get; init; }

    /// <summary>Gets the responsibilities paragraph, or null to omit the whole section.</summary>
    public string? Responsibilities { get; init; } = MethodStatementFixture.ResponsibilitiesText;

    /// <summary>Gets a value indicating whether the inspections heading appears twice.</summary>
    public bool DuplicateInspectionsHeading { get; init; }

    /// <summary>Gets the scope paragraph's lead text.</summary>
    public string Scope { get; init; } = "This method statement covers replacing air-handling unit AHU-3 on the roof of Building B";
}

/// <summary>
/// A supplier's installation method statement, as a customer receives it for review: headings,
/// an approvals table, the customer's open comment from an earlier round, and the supplier's
/// own pending tracked insertion. Names and project details are fictional.
/// </summary>
public static class MethodStatementFixture
{
    /// <summary>Document file name used by the sample.</summary>
    public const string FileName = "method-statement.docx";

    /// <summary>Default installation-window line.</summary>
    public const string InstallationWindowText = "Installation window: 2026-11-09 to 2026-11-13.";

    /// <summary>First inspections paragraph.</summary>
    public const string HoldPoint1 =
        "Hold point 1: after the new unit is set on its frame, the customer's structural engineer checks that every fixing bolt is torqued to 45 N·m, as specified on drawing S-104.";

    /// <summary>Second inspections paragraph.</summary>
    public const string HoldPoint2 =
        "Hold point 2: before the unit is energised, the electrical inspector measures insulation resistance; the acceptance value is at least 1 megohm. " +
        "Work does not continue past a hold point until the inspector signs it off.";

    /// <summary>Default first recovery paragraph: the gap the review should find.</summary>
    public const string RecoveryText = "If installation fails, the team will assess the situation and agree next steps.";

    /// <summary>Second recovery paragraph.</summary>
    public const string RecoveryStorageText = "The existing unit is stored on site until the new unit is commissioned.";

    /// <summary>Default responsibilities paragraph: a reference, not evidence.</summary>
    public const string ResponsibilitiesText = "Responsibilities are as agreed at the kick-off meeting (minutes KM-12).";

    /// <summary>Text of the customer's existing comment.</summary>
    public const string ExistingComment = "Confirm the crane exclusion zone with site security.";

    /// <summary>Text of the supplier's existing tracked insertion.</summary>
    public const string ExistingInsertion = " and reconnecting it to the existing ductwork";

    /// <summary>Text of the pending insertion used by the unsupported-state variant.</summary>
    public const string PendingRecoveryInsertion = " The site supervisor reinstates the existing unit.";

    /// <summary>Paragraph id of the first recovery paragraph.</summary>
    public const string RecoveryParaId = "10000010";

    /// <summary>Builds the method statement.</summary>
    /// <param name="options">Variations, or null for the default document.</param>
    /// <returns>The .docx bytes.</returns>
    public static byte[] Create(MethodStatementOptions? options = null)
    {
        options ??= new MethodStatementOptions();
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            AddStyles(main);

            var comments = main.AddNewPart<WordprocessingCommentsPart>();
            comments.Comments = new Comments(new Comment(
                new Paragraph(new Run(new Text(ExistingComment))))
            {
                Id = "0",
                Author = "Maria Jensen",
                Initials = "MJ",
                Date = new DateTime(2026, 10, 12, 9, 30, 0, DateTimeKind.Utc),
            });
            comments.Comments.Save();

            var body = new Body(
                Para("10000001", "Title", "Installation method statement: rooftop air-handling unit AHU-3"),
                Para("10000002", "Heading1", "Scope"),
                ScopeParagraph(options.Scope),
                Para("10000004", "Heading1", "Programme and approvals"),
                Para("10000005", null, options.InstallationWindow),
                ApprovalsTable(),
                Para("10000006", "Heading1", "Lifting and access"),
                CommentedParagraph("10000007", "A 60-tonne mobile crane lifts the old and new units; the lift plan is in Appendix C."),
                Para("1000000C", "Heading1", "Inspections"),
                Para("1000000D", null, HoldPoint1),
                Para("1000000E", null, HoldPoint2),
                Para("1000000F", "Heading1", "Recovery procedure"),
                options.PendingRevisionInRecovery ? PendingParagraph(RecoveryParaId, options.Recovery) : Para(RecoveryParaId, null, options.Recovery),
                Para("10000011", null, RecoveryStorageText));

            if (options.DuplicateInspectionsHeading)
            {
                body.Append(
                    Para("10000012", "Heading1", "Inspections"),
                    Para("10000013", null, "Further inspections are listed in the quality plan."));
            }

            if (options.Responsibilities is { } responsibilities)
            {
                body.Append(
                    Para("10000014", "Heading1", "Responsibilities"),
                    Para("10000015", null, responsibilities));
            }

            main.Document = new Document(body);
            main.Document.Save();
        }

        return ms.ToArray();
    }

    private static Paragraph ScopeParagraph(string lead)
    {
        var paragraph = Para("10000003", null, lead);
        paragraph.AppendChild(new InsertedRun(new Run(new Text(ExistingInsertion) { Space = SpaceProcessingModeValues.Preserve }))
        {
            Id = "101",
            Author = "Tom Berger",
            Date = new DateTime(2026, 10, 14, 14, 5, 0, DateTimeKind.Utc),
        });
        paragraph.AppendChild(new Run(new Text(".")));
        return paragraph;
    }

    private static Paragraph PendingParagraph(string paraId, string text)
    {
        var paragraph = Para(paraId, null, text);
        paragraph.AppendChild(new InsertedRun(new Run(new Text(PendingRecoveryInsertion) { Space = SpaceProcessingModeValues.Preserve }))
        {
            Id = "102",
            Author = "Tom Berger",
            Date = new DateTime(2026, 10, 15, 10, 0, 0, DateTimeKind.Utc),
        });
        return paragraph;
    }

    private static Paragraph CommentedParagraph(string paraId, string text)
    {
        var paragraph = new Paragraph();
        SetParaId(paragraph, paraId);
        paragraph.Append(
            new CommentRangeStart { Id = "0" },
            new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }),
            new CommentRangeEnd { Id = "0" },
            new Run(new RunProperties(new RunStyle { Val = "CommentReference" }), new CommentReference { Id = "0" }));
        return paragraph;
    }

    private static Table ApprovalsTable()
    {
        var border = new TableBorders(
            new TopBorder { Val = BorderValues.Single, Size = 4 },
            new BottomBorder { Val = BorderValues.Single, Size = 4 },
            new LeftBorder { Val = BorderValues.Single, Size = 4 },
            new RightBorder { Val = BorderValues.Single, Size = 4 },
            new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
            new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 });

        var table = new Table(
            new TableProperties(border, new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct }),
            new TableGrid(new GridColumn { Width = "3000" }, new GridColumn { Width = "3000" }, new GridColumn { Width = "3000" }));

        var rows = new[]
        {
            ("20000001", "Role", "20000002", "Name", "20000003", "Status"),
            ("20000004", "Supplier project manager", "20000005", "Jonas Weber", "20000006", "Signed"),
            ("20000007", "Site supervisor", "20000008", "Priya Raman", "20000009", "Signed"),
        };

        foreach (var (id1, c1, id2, c2, id3, c3) in rows)
        {
            table.AppendChild(new TableRow(Cell(id1, c1), Cell(id2, c2), Cell(id3, c3)));
        }

        return table;
    }

    private static TableCell Cell(string paraId, string text) =>
        new(new TableCellProperties(new TableCellWidth { Width = "3000", Type = TableWidthUnitValues.Dxa }), Para(paraId, null, text));

    private static Paragraph Para(string paraId, string? styleId, string text)
    {
        var paragraph = new Paragraph();
        SetParaId(paragraph, paraId);
        if (styleId is not null)
        {
            paragraph.ParagraphProperties = new ParagraphProperties(new ParagraphStyleId { Val = styleId });
        }

        paragraph.AppendChild(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
        return paragraph;
    }

    private static void SetParaId(Paragraph paragraph, string paraId) =>
        paragraph.SetAttribute(new OpenXmlAttribute("w14", "paraId", "http://schemas.microsoft.com/office/word/2010/wordml", paraId));

    private static void AddStyles(MainDocumentPart main)
    {
        var styles = main.AddNewPart<StyleDefinitionsPart>();
        styles.Styles = new Styles(
            new Style(
                new StyleName { Val = "Normal" },
                new StyleRunProperties(new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri" }, new FontSize { Val = "22" }))
            { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true },
            new Style(
                new StyleName { Val = "Title" },
                new BasedOn { Val = "Normal" },
                new StyleParagraphProperties(new SpacingBetweenLines { After = "240" }),
                new StyleRunProperties(new Bold(), new FontSize { Val = "36" }))
            { Type = StyleValues.Paragraph, StyleId = "Title" },
            new Style(
                new StyleName { Val = "heading 1" },
                new BasedOn { Val = "Normal" },
                new NextParagraphStyle { Val = "Normal" },
                new StyleParagraphProperties(new KeepNext(), new SpacingBetweenLines { Before = "240", After = "80" }, new OutlineLevel { Val = 0 }),
                new StyleRunProperties(new Bold(), new FontSize { Val = "28" }))
            { Type = StyleValues.Paragraph, StyleId = "Heading1" },
            new Style(
                new StyleName { Val = "annotation reference" },
                new StyleRunProperties(new FontSize { Val = "16" }))
            { Type = StyleValues.Character, StyleId = "CommentReference" });
        styles.Styles.Save();
    }
}
