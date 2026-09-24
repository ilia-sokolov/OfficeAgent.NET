using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeAgent.Samples.PageCounter;

/// <summary>
/// Simple page counter that inspects Word document properties without requiring rendering.
/// Uses the core.xml properties if available, or estimates based on document structure.
/// </summary>
public static class WordPageCounterSimple
{
    /// <summary>
    /// Gets the page count from a Word document by inspecting the core properties.
    /// </summary>
    /// <remarks>
    /// This approach:
    /// - Reads the Word document without external dependencies
    /// - Extracts page count from core.xml if Word has calculated it
    /// - Falls back to counting section breaks if properties are unavailable
    ///
    /// Limitations:
    /// - Requires Word to have opened and saved the document (core.xml may be outdated)
    /// - Does not account for dynamic formatting changes that affect pagination
    /// - Estimates may be inaccurate for complex documents
    /// - Does not handle all Office features (e.g., floating objects, complex tables)
    /// </remarks>
    public static int? GetPageCount(string filePath)
    {
        try
        {
            using (var doc = WordprocessingDocument.Open(filePath, isEditable: false))
            {
                // Try to get page count from core properties
                var coreProperties = doc.CoreFilePropertiesPart;
                if (coreProperties != null)
                {
                    using var stream = coreProperties.GetStream();
                    var coreXml = XDocument.Load(stream);
                    var ns = XNamespace.Get("http://schemas.openxmlformats.org/officeDocument/2006/extended-properties");

                    var pages = coreXml.Descendants(ns + "Pages").FirstOrDefault();
                    if (pages?.Value != null && int.TryParse(pages.Value, out var pageCount))
                    {
                        return pageCount;
                    }
                }

                // Fallback: estimate from document structure
                return EstimatePageCount(doc);
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to count pages in {filePath}", ex);
        }
    }

    /// <summary>
    /// Estimates page count by analyzing document structure.
    /// This is a rough estimate and not always accurate.
    /// </summary>
    private static int EstimatePageCount(WordprocessingDocument doc)
    {
        var body = doc.MainDocumentPart?.Document?.Body;
        if (body == null) return 1;

        // Very rough estimation: count paragraphs and divide by typical lines per page
        var paragraphs = body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().Count();
        var tables = body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Table>().Count();

        // Rough estimate: 40-60 paragraphs per page depending on content
        // Tables take more space, so add extra
        var estimatedPages = Math.Max(1, (paragraphs / 50) + (tables / 2));
        return estimatedPages;
    }
}
