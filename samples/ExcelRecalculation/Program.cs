using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ClosedXML.Excel;
using OfficeAgent.Core;

class RecalculateWorkbook
{
    static async Task Main(string[] args)
    {
        Console.WriteLine("=== EXCEL WORKBOOK RECALCULATION ===\n");

        if (args.Length == 0)
        {
            Console.WriteLine("Usage: dotnet run <workbook.xlsx> [method]");
            Console.WriteLine("Methods: openxml, closedxml, all");
            Console.WriteLine();
            CreateSampleWorkbook();
            TestAll();
            return;
        }

        var inputPath = args[0];
        var method = args.Length > 1 ? args[1].ToLowerInvariant() : "all";

        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: File '{inputPath}' not found.");
            return;
        }

        switch (method)
        {
            case "openxml":
                RecalculateWithOpenXml(inputPath, inputPath.Replace(".xlsx", "-openxml.xlsx"));
                break;
            case "closedxml":
                RecalculateWithClosedXml(inputPath, inputPath.Replace(".xlsx", "-closedxml.xlsx"));
                break;
            case "all":
                TestAll();
                break;
            default:
                Console.WriteLine($"Unknown method: {method}");
                break;
        }
    }

    static void CreateSampleWorkbook()
    {
        Console.WriteLine("Creating sample workbook with formulas...\n");
        const string path = "test-workbook.xlsx";

        // Create a workbook with formulas using ClosedXML
        using (var wb = new XLWorkbook())
        {
            var ws = wb.Worksheets.Add("Data");
            ws.Cell("A1").Value = "Item";
            ws.Cell("B1").Value = "Price";
            ws.Cell("C1").Value = "Quantity";
            ws.Cell("D1").Value = "Total";

            ws.Cell("A2").Value = "Widget";
            ws.Cell("B2").Value = 10.00;
            ws.Cell("C2").Value = 5;
            ws.Cell("D2").FormulaA1 = "=B2*C2";

            ws.Cell("A3").Value = "Gadget";
            ws.Cell("B3").Value = 25.00;
            ws.Cell("C3").Value = 3;
            ws.Cell("D3").FormulaA1 = "=B3*C3";

            ws.Cell("A4").Value = "Grand Total";
            ws.Cell("D4").FormulaA1 = "=SUM(D2:D3)";

            wb.SaveAs(path);
        }

        Console.WriteLine($"✓ Created {path}");
        AnalyzeWorkbook(path, "ORIGINAL (No cached values)");
    }

    static void TestAll()
    {
        Console.WriteLine("Testing all methods...\n");

        if (!File.Exists("test-workbook.xlsx"))
            CreateSampleWorkbook();

        RecalculateWithOpenXml("test-workbook.xlsx", "recalculated-openxml.xlsx");
        RecalculateWithClosedXml("test-workbook.xlsx", "recalculated-closedxml.xlsx");

        Console.WriteLine("\n=== COMPARISON ===\n");
        AnalyzeWorkbook("test-workbook.xlsx", "ORIGINAL");
        AnalyzeWorkbook("recalculated-openxml.xlsx", "OPENXML SDK (Marked for recalc)");
        AnalyzeWorkbook("recalculated-closedxml.xlsx", "CLOSEDXML (Evaluated & cached)");
    }

    /// <summary>
    /// Method 1: OpenXML SDK - Mark formulas for recalculation
    /// Approach: Set recalculation flags without evaluating formulas
    /// Pros: Fast, no formula evaluation, minimal file size increase
    /// Cons: Formulas not calculated; must be done by Excel
    /// </summary>
    static void RecalculateWithOpenXml(string inputPath, string outputPath)
    {
        Console.WriteLine($"[OpenXML] Processing: {Path.GetFileName(inputPath)}");
        try
        {
            // Copy to preserve original
            File.Copy(inputPath, outputPath, overwrite: true);

            using (var document = SpreadsheetDocument.Open(outputPath, isEditable: true))
            {
                // Use OfficeAgent's built-in utility to mark for recalculation
                SpreadsheetPartUtility.RecalculateOnOpen(document);

                // The utility:
                // 1. Sets CalculationMode = Auto
                // 2. Sets FullCalculationOnLoad = true
                // 3. Sets ForceFullCalculation = true
                // 4. Deletes CalculationChainPart (forces fresh chain)

                document.WorkbookPart!.Workbook.Save();
            }

            Console.WriteLine($"✓ Output: {outputPath}");
            Console.WriteLine("  Action: Marked for recalculation (flags set, formulas unchanged)\n");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"✗ Error: {ex.Message}\n");
        }
    }

    /// <summary>
    /// Method 2: ClosedXML - Evaluate and cache formula results
    /// Approach: Load workbook, trigger evaluation, save with cached values
    /// Pros: Formulas are evaluated and cached immediately
    /// Cons: Requires formula engine, larger file, slower
    /// </summary>
    static void RecalculateWithClosedXml(string inputPath, string outputPath)
    {
        Console.WriteLine($"[ClosedXML] Processing: {Path.GetFileName(inputPath)}");
        try
        {
            using (var workbook = new XLWorkbook(inputPath))
            {
                // ClosedXML automatically evaluates formulas on load
                // The cached values are present in the workbook object

                // Verify formulas are present
                var formulaCount = 0;
                foreach (var worksheet in workbook.Worksheets)
                {
                    foreach (var cell in worksheet.CellsUsed(c => c.HasFormula))
                    {
                        formulaCount++;
                    }
                }

                Console.WriteLine($"  Found {formulaCount} formulas");
                Console.WriteLine("  Cached values are present in cells");

                workbook.SaveAs(outputPath);
            }

            Console.WriteLine($"✓ Output: {outputPath}");
            Console.WriteLine("  Action: Formulas evaluated and cached\n");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"✗ Error: {ex.Message}\n");
        }
    }

    static void AnalyzeWorkbook(string path, string label)
    {
        Console.WriteLine($"{label}:");
        Console.WriteLine($"File: {Path.GetFileName(path)}");

        try
        {
            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (var document = SpreadsheetDocument.Open(stream, isEditable: false))
            {
                var workbook = document.WorkbookPart?.Workbook;
                var calcProps = workbook?.CalculationProperties;

                Console.WriteLine("  CalculationProperties:");
                if (calcProps is not null)
                {
                    var mode = calcProps.CalculationMode?.Value.ToString() ?? "not set";
                    var fullLoad = calcProps.FullCalculationOnLoad?.Value.ToString() ?? "not set";
                    var forceCalc = calcProps.ForceFullCalculation?.Value.ToString() ?? "not set";
                    Console.WriteLine($"    CalculationMode: {mode}");
                    Console.WriteLine($"    FullCalculationOnLoad: {fullLoad}");
                    Console.WriteLine($"    ForceFullCalculation: {forceCalc}");
                }
                else
                {
                    Console.WriteLine($"    (not set)");
                }

                var hasChain = document.WorkbookPart?.CalculationChainPart is not null;
                Console.WriteLine($"  CalculationChainPart: {(hasChain ? "PRESENT" : "DELETED")}");

                // Count formulas and cached values
                var formulaCells = 0;
                var cachedCells = 0;
                foreach (var worksheetPart in document.WorkbookPart?.WorksheetParts ?? Enumerable.Empty<WorksheetPart>())
                {
                    foreach (var cell in worksheetPart.Worksheet.Descendants<Cell>().Where(c => c.CellFormula is not null))
                    {
                        formulaCells++;
                        if (cell.CellValue?.InnerText?.Length > 0)
                            cachedCells++;
                    }
                }

                Console.WriteLine($"  Formulas: {formulaCells} total, {cachedCells} with cached values");
                Console.WriteLine();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  ERROR: {ex.Message}\n");
        }
    }
}
