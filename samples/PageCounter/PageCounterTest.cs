using System.Text;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeAgent.Samples.PageCounter;

/// <summary>
/// Demonstrates and tests different approaches to page counting in Word documents.
/// </summary>
public static class PageCounterTest
{
    public static void Run()
    {
        var testDocs = new[]
        {
            "/samples/documents/services-agreement.docx",
            "/tests/OfficeAgent.Tests/Corpus/v0.8.0/assembly-source-a.docx",
        };

        foreach (var doc in testDocs)
        {
            var fullPath = Path.Combine(
                Path.GetDirectoryName(typeof(PageCounterTest).Assembly.Location)
                ?? Environment.CurrentDirectory,
                "..",
                "..",
                "..",
                "..",
                "..",
                doc.TrimStart('/'));

            if (!File.Exists(fullPath))
            {
                Console.WriteLine($"Test document not found: {fullPath}");
                continue;
            }

            Console.WriteLine($"\n=== Testing: {Path.GetFileName(fullPath)} ===");

            try
            {
                // Approach 1: Read core properties
                var corePageCount = GetPageCountFromCoreProperties(fullPath);
                Console.WriteLine($"Core properties page count: {corePageCount ?? 0}");

                // Approach 2: Estimate from structure
                var estimatedCount = EstimatePageCountFromStructure(fullPath);
                Console.WriteLine($"Estimated page count: {estimatedCount}");

                // Analyze document structure
                AnalyzeStructure(fullPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    private static int? GetPageCountFromCoreProperties(string filePath)
    {
        using (var doc = WordprocessingDocument.Open(filePath, isEditable: false))
        {
            var coreProperties = doc.CoreFilePropertiesPart;
            if (coreProperties == null) return null;

            using var stream = coreProperties.GetStream();
            using var reader = new StreamReader(stream);
            var content = reader.ReadToEnd();

            // Look for Pages property in core.xml
            var pagesMatch = System.Text.RegularExpressions.Regex.Match(
                content, @"<.*?Pages>(\d+)<");

            if (pagesMatch.Success && int.TryParse(pagesMatch.Groups[1].Value, out var count))
            {
                return count;
            }

            return null;
        }
    }

    private static int EstimatePageCountFromStructure(string filePath)
    {
        using (var doc = WordprocessingDocument.Open(filePath, isEditable: false))
        {
            var body = doc.MainDocumentPart?.Document?.Body;
            if (body == null) return 1;

            var paragraphs = body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().Count();
            var tables = body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Table>().Count();

            // Rough estimate
            var estimate = Math.Max(1, (paragraphs / 50) + (tables / 2));
            return estimate;
        }
    }

    private static void AnalyzeStructure(string filePath)
    {
        using (var doc = WordprocessingDocument.Open(filePath, isEditable: false))
        {
            var body = doc.MainDocumentPart?.Document?.Body;
            if (body == null) return;

            var paragraphs = body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().Count();
            var tables = body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Table>().Count();
            var sections = body.Descendants<DocumentFormat.OpenXml.Wordprocessing.SectionProperties>().Count();

            Console.WriteLine($"  Paragraphs: {paragraphs}");
            Console.WriteLine($"  Tables: {tables}");
            Console.WriteLine($"  Sections: {sections}");
        }
    }
}
