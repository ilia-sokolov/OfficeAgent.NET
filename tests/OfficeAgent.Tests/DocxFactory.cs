using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.Tests;

/// <summary>Builds small, valid .docx fixtures in memory for the flow tests.</summary>
internal static class DocxFactory
{
    public const string HeadingText = "Service Agreement";
    public const string ClauseText = "Acme Corp shall provide services to Acme Corp.";
    public const string ClientControlTag = "ClientName";
    public const string ClientPlaceholder = "PLACEHOLDER";
    public const string DateText = "Effective date: 2020-01-01.";






    public static byte[] Contract()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();

            var heading = Para("00000001", "Heading1", new Run(new Text(HeadingText)));

            var clause = Para("00000002", null,
                new Run(new Text("Acme ") { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new Text("Corp") { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new Text(" shall provide services to ") { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new Text("Acme Corp") { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new Text(".") { Space = SpaceProcessingModeValues.Preserve }));

            var clientControl = new SdtRun(
                new SdtProperties(new Tag { Val = ClientControlTag }, new SdtId { Val = 101 }),
                new SdtContentRun(new Run(new Text(ClientPlaceholder))));
            var clientPara = Para("00000003", null, new Run(new Text("Client: ") { Space = SpaceProcessingModeValues.Preserve }));
            clientPara.AppendChild(clientControl);

            var date = Para("00000004", null, new Run(new Text(DateText)));

            main.Document = new Document(new Body(heading, clause, clientPara, date));

            var stylePart = main.AddNewPart<StyleDefinitionsPart>();
            stylePart.Styles = new Styles(
                Style("Heading1", "heading 1", custom: false),
                Style("Quote", "Quote", custom: false));

            main.Document.Save();
            stylePart.Styles.Save();
        }

        return ms.ToArray();
    }

    private static Paragraph Para(string paraId, string? styleId, params OpenXmlElement[] runs)
    {
        var paragraph = new Paragraph { ParagraphId = paraId };
        if (styleId is not null)
            paragraph.ParagraphProperties = new ParagraphProperties(new ParagraphStyleId { Val = styleId });
        foreach (var run in runs)
            paragraph.AppendChild(run);
        return paragraph;
    }

    private static Style Style(string id, string name, bool custom) => new()
    {
        Type = StyleValues.Paragraph,
        StyleId = id,
        StyleName = new StyleName { Val = name },
        CustomStyle = custom
    };
}
