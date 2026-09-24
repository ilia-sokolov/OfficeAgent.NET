using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeAgent.Samples.PageCounter;

/// <summary>
/// Quick test using just the OpenXML SDK without external rendering tools.
/// </summary>
class SimplePageTest
{
    static void Main(string[] args)
    {
        if (args.Length == 0)
        {
            Console.WriteLine("Usage: SimplePageTest document.docx");
            return;
        }

        var docPath = args[0];
        if (!File.Exists(docPath))
        {
            Console.WriteLine($"File not found: {docPath}");
            return;
        }

        Console.WriteLine($"Analyzing: {Path.GetFileName(docPath)}");
        Console.WriteLine();

        try
        {
            // Open without rendering
            using (var doc = WordprocessingDocument.Open(docPath, isEditable: false))
            {
                // Check core properties
                var coreProps = doc.CoreFilePropertiesPart;
                if (coreProps != null)
                {
                    using var propStream = coreProps.GetStream();
                    var propXml = XDocument.Load(propStream);

                    Console.WriteLine("Core Properties XML:");
                    Console.WriteLine(propXml);
                    Console.WriteLine();
                }
                else
                {
                    Console.WriteLine("No core properties found");
                }

                // Analyze structure
                var body = doc.MainDocumentPart?.Document?.Body;
                if (body != null)
                {
                    var paragraphs = body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().Count();
                    var tables = body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Table>().Count();
                    var sections = body.Descendants<DocumentFormat.OpenXml.Wordprocessing.SectionProperties>().Count();

                    Console.WriteLine("Document Structure:");
                    Console.WriteLine($"  Paragraphs: {paragraphs}");
                    Console.WriteLine($"  Tables: {tables}");
                    Console.WriteLine($"  Sections: {sections}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex}");
        }
    }
}
