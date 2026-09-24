using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.Samples.RequirementsReview;

/// <summary>A fictional supplier proposal whose data-privacy terms need review.</summary>
public static class PrivacyProposalFixture
{
    public const string FileName = "supplier-proposal.docx";
    public const string Section = "Data privacy";
    public const string Passage =
        "The supplier will upload site access logs, including names and mobile numbers, " +
        "to the shared project folder for all project members. The logs will be kept " +
        "after handover in case they are needed.";

    public static byte[] Create()
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            var styles = main.AddNewPart<StyleDefinitionsPart>();
            styles.Styles = new Styles(
                new Style(new StyleName { Val = "Normal" })
                { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true },
                new Style(new StyleName { Val = "Title" }, new BasedOn { Val = "Normal" })
                { Type = StyleValues.Paragraph, StyleId = "Title" },
                new Style(new StyleName { Val = "heading 1" }, new BasedOn { Val = "Normal" },
                    new StyleParagraphProperties(new OutlineLevel { Val = 0 }))
                { Type = StyleValues.Paragraph, StyleId = "Heading1" });
            styles.Styles.Save();

            main.Document = new Document(new Body(
                Paragraph("30000001", "Title", "Supplier proposal"),
                Paragraph("30000002", "Heading1", Section),
                Paragraph("30000003", null, Passage)));
            main.Document.Save();
        }

        return stream.ToArray();
    }

    private static Paragraph Paragraph(string id, string? style, string text)
    {
        var paragraph = new Paragraph();
        paragraph.SetAttribute(new OpenXmlAttribute("w14", "paraId", "http://schemas.microsoft.com/office/word/2010/wordml", id));
        if (style is not null)
        {
            paragraph.ParagraphProperties = new ParagraphProperties(new ParagraphStyleId { Val = style });
        }

        paragraph.AppendChild(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
        return paragraph;
    }
}
