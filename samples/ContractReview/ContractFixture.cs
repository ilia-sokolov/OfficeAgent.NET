using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.Samples.ContractReview;

/// <summary>
/// A small supplier agreement with wording the sample playbook is meant to catch:
/// 60-day payment terms, uncapped liability, automatic renewal, and a governing
/// law outside the EU.
/// </summary>
public static class ContractFixture
{
    /// <summary>Text of the payment clause.</summary>
    public const string PaymentClause =
        "The Customer shall pay each undisputed invoice within 60 days of receipt.";

    /// <summary>Text of the liability clause.</summary>
    public const string LiabilityClause =
        "The Supplier accepts unlimited liability for any loss arising under this Agreement.";

    /// <summary>Text of the renewal clause.</summary>
    public const string RenewalClause =
        "This Agreement shall automatically renew for successive periods of twelve months.";

    /// <summary>Text of the governing-law clause.</summary>
    public const string GoverningLawClause =
        "This Agreement is governed by the laws of the State of New York.";

    /// <summary>Builds the fixture.</summary>
    /// <param name="paymentDays">Number of days in the payment clause.</param>
    /// <returns>The .docx bytes.</returns>
    /// <param name="governingLaw">Text of the governing-law clause tail.</param>
    public static byte[] Create(int paymentDays = 60, string governingLaw = "the laws of the State of New York")
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body(
                Para("10000001", "Heading1", "Supply Agreement"),
                Para("10000002", null, "1. Payment"),
                Para("10000003", null, PaymentClause.Replace("60 days", $"{paymentDays} days", StringComparison.Ordinal)),
                Para("10000004", null, "2. Liability"),
                Para("10000005", null, LiabilityClause),
                Para("10000006", null, "3. Term"),
                Para("10000007", null, RenewalClause),
                Para("10000008", null, "4. Governing law"),
                Para("10000009", null, $"This Agreement is governed by {governingLaw}.")));
            main.Document.Save();
        }

        return ms.ToArray();
    }

    private static Paragraph Para(string paraId, string? styleId, string text)
    {
        var paragraph = new Paragraph();
        paragraph.SetAttribute(new OpenXmlAttribute(
            "w14", "paraId", "http://schemas.microsoft.com/office/word/2010/wordml", paraId));

        if (styleId is not null)
        {
            paragraph.ParagraphProperties = new ParagraphProperties(new ParagraphStyleId { Val = styleId });
        }

        paragraph.AppendChild(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
        return paragraph;
    }
}
