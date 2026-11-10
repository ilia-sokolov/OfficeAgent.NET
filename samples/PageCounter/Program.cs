using System.Text.Json;
using OfficeAgent.Abstractions;
using OfficeAgent.Rendering;
using DocumentFormat.OpenXml.Packaging;

var arguments = args.Where(a => !a.StartsWith("--")).ToArray();
var quiet = args.Contains("--quiet", StringComparer.Ordinal);
var useOpenXml = args.Contains("--openxml", StringComparer.Ordinal);

if (arguments.Length == 0)
{
    Console.Error.WriteLine("Word Document Page Counter - Evaluation of Different Approaches");
    Console.Error.WriteLine();
    Console.Error.WriteLine("Usage: PageCounter document.docx [--openxml] [--quiet]");
    Console.Error.WriteLine();
    Console.Error.WriteLine("Options:");
    Console.Error.WriteLine("  --openxml            Use OpenXML SDK approach (no LibreOffice required)");
    Console.Error.WriteLine("  --quiet              Suppress verbose output");
    Console.Error.WriteLine();
    Console.Error.WriteLine("Default (no --openxml): Uses LibreOffice rendering for accurate page count");
    Console.Error.WriteLine("With --openxml: Uses the page count Word recorded, or a structure estimate");
    return 2;
}

int exitCode = 0;

foreach (var documentPath in arguments)
{
    var fullPath = Path.GetFullPath(documentPath);
    if (!File.Exists(fullPath))
    {
        if (!quiet) Console.Error.WriteLine($"Error: File not found: {fullPath}");
        exitCode = 1;
        continue;
    }

    if (!quiet)
        Console.WriteLine($"Processing: {Path.GetFileName(fullPath)}");

    try
    {
        if (useOpenXml)
        {
            // Approach 1: OpenXML SDK (no external dependencies)
            var (pageCount, recorded) = GetPageCountFromOpenXml(fullPath);
            Console.WriteLine($"Page count: {pageCount}");

            if (!quiet)
            {
                Console.WriteLine(recorded
                    ? "  Method: OpenXML SDK, the Pages value Word recorded in docProps/app.xml"
                    : "  Method: OpenXML SDK, estimated from the body (no Pages value recorded)");
                Console.WriteLine(recorded
                    ? "  Accuracy: Medium (Word's count when it last saved the file)"
                    : "  Accuracy: Low (one page per 50 paragraphs plus one per 3 tables)");
                Console.WriteLine($"  External dependencies: None");
            }
        }
        else
        {
            // Approach 2: LibreOffice rendering (most accurate)
            if (!await TestLibreOfficeApproachAsync(fullPath, quiet))
                exitCode = 1;
        }
    }
    catch (OperationCanceledException)
    {
        Console.Error.WriteLine($"Rendering timed out for {Path.GetFileName(fullPath)}");
        exitCode = 1;
    }
    catch (Exception ex) when (ex is FileNotFoundException or DirectoryNotFoundException or IOException)
    {
        Console.Error.WriteLine($"I/O error: {ex.Message}");
        exitCode = 1;
    }
    catch (RenderFailedException rfe)
    {
        Console.Error.WriteLine($"Render error: {rfe.FailureCode} - {rfe.Message}");
        exitCode = 1;
    }
    catch (Exception ex)
    {
        Console.Error.WriteLine($"Error: {ex.Message}");
        exitCode = 1;
    }

    if (!quiet) Console.WriteLine();
}

return exitCode;

/// <summary>
/// Approach 1: Using OpenXML SDK to read the page count Word recorded, or estimate it.
/// Requires NO external dependencies. Recorded is false when the count is an estimate.
/// </summary>
static (int Pages, bool Recorded) GetPageCountFromOpenXml(string filePath)
{
    using (var doc = WordprocessingDocument.Open(filePath, isEditable: false))
    {
        // Word records the page count in the extended properties (docProps/app.xml) when it saves
        var pages = doc.ExtendedFilePropertiesPart?.Properties?.Pages?.Text;
        if (int.TryParse(pages, out var pageCount) && pageCount > 0)
        {
            return (pageCount, true);
        }

        // Fallback: estimate from document structure
        var body = doc.MainDocumentPart?.Document?.Body;
        if (body == null) return (1, false);

        var paragraphs = body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().Count();
        var tables = body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Table>().Count();

        // Rough estimate: typical page has 40-60 paragraphs
        var estimatedPages = Math.Max(1, (paragraphs / 50) + (tables / 3));
        return (estimatedPages, false);
    }
}

/// <summary>
/// Approach 2: Using LibreOffice to render document to PDF, then rasterize to count pages.
/// Most accurate but requires LibreOffice installation. Returns false when the render failed.
/// </summary>
async Task<bool> TestLibreOfficeApproachAsync(string filePath, bool quiet)
{
    using var content = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, useAsync: true);

    var renderer = new LibreOfficeDocumentRenderer(new LibreOfficeRendererOptions
    {
        LibreOfficeExecutable = Environment.GetEnvironmentVariable("SOFFICE") ?? "soffice",
        PdfToPpmExecutable = Environment.GetEnvironmentVariable("PDFTOPPM") ?? "pdftoppm"
    });

    var fileSize = new FileInfo(filePath).Length;
    if (!quiet)
        Console.WriteLine($"  File size: {FormatBytes(fileSize)}");

    var options = new RenderOptions
    {
        FileName = Path.GetFileName(filePath),
        Dpi = 144,
        Timeout = TimeSpan.FromSeconds(120),
        MaximumInputBytes = 100L * 1024 * 1024,
        MaximumOutputBytes = 500L * 1024 * 1024,
        MaximumPages = 10000,
        MaximumWorkingSetBytes = 512L * 1024 * 1024
    };

    if (!quiet)
        Console.WriteLine($"  Rendering with 144 DPI, 120s timeout...");

    var startTime = DateTime.UtcNow;
    var result = await renderer.RenderAsync(content, options);
    var elapsedMs = (DateTime.UtcNow - startTime).TotalMilliseconds;

    if (!result.Succeeded)
    {
        Console.Error.WriteLine($"Rendering failed: {result.FailureCode}");
        Console.Error.WriteLine($"  Details: {result.Message}");
        return false;
    }

    var pageCount = result.PageCount;
    Console.WriteLine($"Page count: {pageCount}");

    if (!quiet)
    {
        Console.WriteLine($"  Method: LibreOffice rendering");
        Console.WriteLine($"  Accuracy: Very High (visual rendering)");
        Console.WriteLine($"  Rendering completed in {elapsedMs:F0} ms");
        if (result.Pages.Count > 0)
        {
            var totalImageBytes = result.Pages.Sum(p => p.Content.Length);
            Console.WriteLine($"  Total image data: {FormatBytes(totalImageBytes)}");
            var avgPageSize = totalImageBytes / result.Pages.Count;
            Console.WriteLine($"  Average page size: {FormatBytes(avgPageSize)}");
        }
    }

    return true;
}

static string FormatBytes(long bytes)
{
    string[] sizes = { "B", "KB", "MB", "GB" };
    double len = bytes;
    int order = 0;
    while (len >= 1024 && order < sizes.Length - 1)
    {
        order++;
        len = len / 1024;
    }
    return $"{len:F2} {sizes[order]}";
}
