# ExcelInputUpdate

Updates the inputs of a pricing workbook, and shows what OfficeAgent does and does not do
with the formulas that depend on them.

OfficeAgent edits workbooks; it has no formula engine. When `setCell` changes a value, it
writes the new value, removes the calculation chain, and marks the workbook for a full
recalculation when Excel opens it. Results already stored in other cells are not
recalculated, so until Excel opens the file they still reflect the old inputs. This sample
makes that visible, and then shows one way to get the new numbers on a server without Excel:

1. **Set the inputs.** `SetCellOp` changes `B2` (consulting days) to 15 and `C3` (licence unit
   price) to 52.5. Each cell keeps its style.
2. **Read the file back.** The totals in `D5:D8` still show the results Excel stored for the
   old inputs. Do not report them as the new totals.
3. **Optional: evaluate without Excel.** ClosedXML opens the saved workbook, recalculates in
   memory, and prints the new totals. It never saves the file: writing it with another
   library would rewrite the package OfficeAgent just edited.

It uses the direct API (`new OfficeAgentClient(new ExcelModule())` over a `StreamHandle`),
with no storage provider or dependency injection. ClosedXML is referenced only for step 3.

## Run it

From the repository root:

```powershell
dotnet run --project samples/ExcelInputUpdate
```

This reads [`samples/documents/pricing-model.xlsx`](../documents/pricing-model.xlsx) and
writes `pricing-model-updated.xlsx` to the current directory. To use other files:

```powershell
dotnet run --project samples/ExcelInputUpdate -- <input.xlsx> <output.xlsx>
```

Expected output:

```text
Before the update (results Excel last calculated and stored):
  D5 =SUM(D2:D4)              stored: 17000
  D6 =-ROUND(D5*0.075,2)      stored: -1275
  D7 =ROUND((D5+D6)*0.2,2)    stored: 3145
  D8 =D5+D6+D7                stored: 18870

Set consulting days (B2) to 15 and licence unit price (C3) to 52.5; written to pricing-model-updated.xlsx

After the update, as stored in the file (Excel recalculates these on open):
  D5 =SUM(D2:D4)              stored: 17000
  D6 =-ROUND(D5*0.075,2)      stored: -1275
  D7 =ROUND((D5+D6)*0.2,2)    stored: 3145
  D8 =D5+D6+D7                stored: 18870

Evaluated in memory with ClosedXML (file not modified):
  D5 =SUM(D2:D4)              = 20070
  D6 =-ROUND(D5*0.075,2)      = -1505.25
  D7 =ROUND((D5+D6)*0.2,2)    = 3712.95
  D8 =D5+D6+D7                = 22277.7
```

Opened in Excel, the output shows the same new totals as step 3, because Excel recalculates
the whole workbook on open.

## What to take from it

- Use OfficeAgent to change a workbook's contents safely. Use Excel, or a formula engine, to
  calculate its results. See [Excel support](../../docs/excel.md).
- A formula engine covers the functions it implements, not all of Excel's. Check yours
  against the functions your workbooks use before relying on its numbers.
- If a file must leave your service with correct stored results, and nobody opens it in Excel
  first, calculate those results with a formula engine and decide deliberately how to write
  them. OfficeAgent does not do it for you.
