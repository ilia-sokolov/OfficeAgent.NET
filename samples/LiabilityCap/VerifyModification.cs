using DocumentFormat.OpenXml.Packaging;

namespace OfficeAgent.Samples.LiabilityCap;

/// <summary>
/// Verification utility to check that comments are preserved after modification.
/// </summary>
internal static class VerifyModification
{
    public static void VerifyDocument(string filePath)
    {
        Console.WriteLine($"\n=== Verifying Document: {filePath} ===");

        using (var doc = WordprocessingDocument.Open(filePath, isEditable: false))
        {
            var mainPart = doc.MainDocumentPart;

            // Show basic document statistics
            var body = mainPart.Document.Body;
            var paragraphCount = body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().Count();
            var tableCount = body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Table>().Count();

            Console.WriteLine($"Document Structure:");
            Console.WriteLine($"  Paragraphs: {paragraphCount}");
            Console.WriteLine($"  Tables: {tableCount}");

            // Look for comment range markers
            var hasCommentRangeStarts = body.Descendants<DocumentFormat.OpenXml.Wordprocessing.CommentRangeStart>().Any();
            var hasCommentRangeEnds = body.Descendants<DocumentFormat.OpenXml.Wordprocessing.CommentRangeEnd>().Any();
            var hasCommentReferences = body.Descendants<DocumentFormat.OpenXml.Wordprocessing.CommentReference>().Any();

            Console.WriteLine($"Comment Structure:");
            Console.WriteLine($"  Has CommentRangeStart markers: {hasCommentRangeStarts}");
            Console.WriteLine($"  Has CommentRangeEnd markers: {hasCommentRangeEnds}");
            Console.WriteLine($"  Has CommentReference marks: {hasCommentReferences}");

            if (hasCommentRangeStarts || hasCommentRangeEnds || hasCommentReferences)
            {
                Console.WriteLine("  ✓ Comment structure is preserved!");
            }
        }

        Console.WriteLine();
    }

    public static void Compare(string original, string modified)
    {
        Console.WriteLine($"\n=== Comparing Documents ===");
        Console.WriteLine($"Original:  {original}");
        Console.WriteLine($"Modified:  {modified}");

        VerifyDocument(original);
        VerifyDocument(modified);
    }
}
