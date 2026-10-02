using System.Globalization;
using ClosedXML.Excel;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Excel;

// Updates the inputs of a pricing workbook. OfficeAgent writes the new values and marks the
// workbook so Excel recalculates every formula when it opens. It does not calculate anything
// itself, so results stored in the file still reflect the old inputs until Excel opens it. When a service needs the new totals without
// Excel, step 3 evaluates the saved file in memory with a formula engine, and never saves
// through it.
//
//   dotnet run --project samples/ExcelInputUpdate
//   dotnet run --project samples/ExcelInputUpdate -- <input.xlsx> <output.xlsx>

var input = args.Length > 0 ? args[0] : Path.Combine("samples", "documents", "pricing-model.xlsx");
var output = args.Length > 1 ? args[1] : "pricing-model-updated.xlsx";

var inputs = new (string Address, string Value, string Label)[]
{
    ("B2", "15", "consulting days"),
    ("C3", "52.5", "licence unit price"),
};
string[] totals = ["D5", "D6", "D7", "D8"];

using var source = new MemoryStream(await File.ReadAllBytesAsync(input), writable: false);
var workbook = new StreamHandle(source, Path.GetFileName(input));
var client = new OfficeAgentClient(new ExcelModule());

var inspection = await client.InspectAsync(workbook);
var sheet = inspection.Worksheets.Single(w => w.Name == "Pricing");

Console.WriteLine("Before the update (results Excel last calculated and stored):");
PrintTotals(inspection);

// 1. Set the inputs. Each cell keeps its style; only its value changes.
var plan = new DocumentPlan
{
    Snapshot = inspection.Snapshot,
    Operations = inputs.Select(i => (PlanOperation)new SetCellOp
    {
        Target = new CellAnchor { SheetId = sheet.SheetId, Address = i.Address },
        Value = i.Value
    }).ToArray()
};

using var result = await client.CommitAsync(workbook, plan);
if (!result.Committed)
{
    foreach (var error in result.Report.Errors)
        Console.Error.WriteLine($"{error.Code}: {error.Message}");
    return 1;
}
await result.SaveAsync(output);
Console.WriteLine();
Console.WriteLine($"Set {string.Join(" and ", inputs.Select(i => $"{i.Label} ({i.Address}) to {i.Value}"))}; written to {output}");

// 2. Read the saved file back. The formulas are unchanged and the workbook is marked for a
//    full recalculation, but any result still stored in the file was computed from the old
//    inputs. Do not report it as the new total: Excel replaces it when the file is opened.
using var saved = new MemoryStream(await File.ReadAllBytesAsync(output), writable: false);
var after = await client.InspectAsync(new StreamHandle(saved, Path.GetFileName(output)));
Console.WriteLine();
Console.WriteLine("After the update, as stored in the file (Excel recalculates these on open):");
PrintTotals(after);

// 3. Optional: when the service needs the new numbers without Excel, evaluate the saved
//    workbook with a formula engine. This reads only; saving through another library would
//    rewrite the package OfficeAgent just edited in place.
using (var evaluated = new XLWorkbook(output))
{
    evaluated.RecalculateAllFormulas();
    var pricing = evaluated.Worksheet("Pricing");
    Console.WriteLine();
    Console.WriteLine("Evaluated in memory with ClosedXML (file not modified):");
    foreach (var address in totals)
    {
        var value = pricing.Cell(address).Value;
        var text = value.IsNumber ? value.GetNumber().ToString(CultureInfo.InvariantCulture) : value.ToString();
        Console.WriteLine($"  {address} ={pricing.Cell(address).FormulaA1,-23} = {text}");
    }
}
return 0;

void PrintTotals(InspectResult result)
{
    foreach (var address in totals)
    {
        var cell = result.Cells.SingleOrDefault(c => c.Anchor.SheetId == sheet.SheetId && c.Anchor.Address == address);
        var stored = string.IsNullOrEmpty(cell?.DisplayValue) ? "(no stored result)" : cell!.DisplayValue;
        Console.WriteLine($"  {address} ={cell?.Formula,-23} stored: {stored}");
    }
}
