# Excel Workbook Recalculation Sample

## Overview

This sample demonstrates three approaches to forcing recalculation of formulas in an Excel workbook (`.xlsx`):

1. **OpenXML SDK** - Mark formulas for recalculation (Excel will calculate on open)
2. **ClosedXML** - Evaluate and cache formula results immediately
3. **OfficeAgent.NET** - Set cells and trigger recalculation through operations

## Scenario

Open an Excel workbook containing formulas, ensure all formulas are recalculated, and save the result.

### Problem Statement

Different scenarios require different approaches:

- **Scenario A**: You only have formulas, no cached values. You need Excel to recalculate when the file opens.
  - **Solution**: Set recalculation flags (OpenXML approach)
  - **When to use**: File will be opened in Excel; you want Excel's calculation engine to handle it

- **Scenario B**: You need cached/evaluated values in the file itself, without requiring Excel.
  - **Solution**: Use ClosedXML to evaluate and cache results
  - **When to use**: Files may be used in systems without Excel; need guaranteed values

- **Scenario C**: You're using OfficeAgent for document editing and also need to trigger recalculation.
  - **Solution**: Use OfficeAgent's SetCellOp (modify then restore) or integrate with edit workflow
  - **When to use**: Already editing workbook with OfficeAgent; need recalculation as part of workflow

## Running the Sample

### Default - Compare All Methods

```bash
dotnet run
```

This will:
1. Create a sample workbook with formulas
2. Apply OpenXML recalculation method
3. Apply ClosedXML evaluation method
4. Display comparison analysis

### Single Method

```bash
# Test OpenXML approach only
dotnet run input.xlsx openxml

# Test ClosedXML approach only
dotnet run input.xlsx closedxml
```

## Output Files

After running `dotnet run`:

- `test-workbook.xlsx` - Original sample with formulas (no cached values)
- `recalculated-openxml.xlsx` - Marked for recalculation (Excel will calculate on open)
- `recalculated-closedxml.xlsx` - Evaluated and cached (values present in file)

## Analysis Output

The sample displays a comparison table:

```
ORIGINAL (No cached values):
  CalculationProperties: (not set)
  CalculationChainPart: PRESENT
  Formulas: 3 total, 0 with cached values

OPENXML SDK (Marked for recalc):
  CalculationMode: Auto
  FullCalculationOnLoad: True
  ForceFullCalculation: True
  CalculationChainPart: DELETED
  Formulas: 3 total, 0 with cached values
    → Excel will recalculate when file opens

CLOSEDXML (Evaluated & cached):
  CalculationProperties: (not set)
  Formulas: 3 total, cached values in memory
    → Workbook object has calculated values
```

## Method Comparison

| Aspect | OpenXML SDK | ClosedXML | OfficeAgent.NET |
|--------|-------------|-----------|-----------------|
| **Speed** | Very fast | Medium | Slow (if only for recalc) |
| **Evaluates formulas** | No | Yes | No (unless modifying cells) |
| **Caches values** | No | Yes | Depends on operation |
| **Requires Excel** | Yes (to calculate) | No | No |
| **File size change** | Minimal | Medium | Minimal (depends) |
| **Use case** | Mark for Excel recalc | Get values without Excel | Edit + recalculate workflow |
| **Complexity** | Low | Low | Medium |

## Implementation Details

### Method 1: OpenXML SDK (Recommended for "mark for recalc")

```csharp
using (var document = SpreadsheetDocument.Open(path, isEditable: true))
{
    SpreadsheetPartUtility.RecalculateOnOpen(document);
    // This utility:
    // - Sets CalculationMode = Auto
    // - Sets FullCalculationOnLoad = true
    // - Sets ForceFullCalculation = true
    // - Deletes CalculationChainPart
}
```

**Advantages:**
- Built into OfficeAgent.Core
- No external dependencies (other than OpenXML SDK)
- Very fast - just modifies XML properties
- Minimal file size increase
- Excel handles actual calculation (accurate for all features)

**Disadvantages:**
- Doesn't actually evaluate formulas
- Requires Excel to open and calculate the workbook

### Method 2: ClosedXML (Recommended for "get calculated values")

```csharp
using (var workbook = new XLWorkbook(inputPath))
{
    // ClosedXML evaluates formulas automatically on load
    workbook.SaveAs(outputPath);
}
```

**Advantages:**
- Automatic formula evaluation
- Cached values are present in file
- No Excel required
- Simple API

**Disadvantages:**
- Slower (formula evaluation overhead)
- Larger file size (cached values stored)
- May not support all advanced Excel features
- Formula evaluation may differ from Excel in edge cases

### Method 3: OfficeAgent.NET (When integrating with document workflow)

OfficeAgent.Excel doesn't have a native "recalculate all" operation because:
- Excel format doesn't support tracked changes (unlike Word)
- OfficeAgent focuses on deterministic, auditable edits
- Formula evaluation is not OfficeAgent's responsibility

However, you can trigger recalculation by:
1. Setting a cell value (forces `RecalculateOnOpen`)
2. Creating a no-op edit that triggers recalculation flag

```csharp
var client = new OpenXmlAgentClient(/* config */);
var plan = new DocumentPlan 
{
    // Edit a cell (any cell) - this triggers RecalculateOnOpen internally
    Operations = new[] { new SetCellOp { /* ... */ } }
};
var result = await client.CommitAsync(/* ... */, plan);
```

## Tool Selection Reasoning

### Condition B: Informed Selection

**Selected: OpenXML SDK** for the "mark for recalculation" scenario

**Why:**
1. **Direct control** - Explicitly set recalculation flags
2. **No side effects** - Doesn't modify cell values
3. **Excel-compatible** - Uses Excel's native calculation engine
4. **Simple** - Just property changes to XML
5. **Available in OfficeAgent.Core** - `SpreadsheetPartUtility.RecalculateOnOpen()` already exists

**Secondary choice: ClosedXML** for the "get evaluated values" scenario

**Why:**
1. **Automatic evaluation** - Handles formula calculation
2. **Self-contained** - No Excel dependency
3. **Cached values** - Results available immediately
4. **Mature library** - Well-tested formula engine

### Condition C: OfficeAgent.NET Suitability for Excel

**Question**: Is OfficeAgent.NET suitable for Excel automation?

**Answer**: Partially suitable. Here's why:

| Capability | Suitable? | Details |
|-----------|-----------|---------|
| **Read/Inspect workbooks** | ✓ YES | Full inspection support (cells, tables, comments) |
| **Edit cell values** | ✓ YES | SetCellOp handles cell modifications |
| **Append table rows** | ✓ YES | AppendTableRowsOp adds data to tables |
| **Add cell comments** | ✓ YES | CommentOp manages comments |
| **Force recalculation** | ⚠ PARTIAL | Works through SetCellOp, but not direct operation |
| **Evaluate formulas** | ✗ NO | Explicitly out of scope (design decision) |
| **Tracked changes** | ✗ NO | Excel doesn't support tracked changes like Word |

**Limitations**:
1. **No tracked changes** - Excel has no equivalent to Word's revision markup
2. **No formula evaluation** - By design (deterministic, auditable edits only)
3. **No "recalculate all" operation** - No direct operation for this task
4. **Cell-level operations only** - Workbook-level settings require direct API access

**Suitable for:**
- Bulk data entry (tables, cells)
- Workbook inspection and validation
- Comment management
- Integration with editing workflows

**Not suitable for:**
- Pure recalculation-only scenarios
- Advanced formula manipulation
- Workbook-level configuration changes
- Complex formula analysis

## Recommendation

- **Use OpenXML SDK** if: You need to mark formulas for recalculation and Excel will open the file
- **Use ClosedXML** if: You need evaluated values cached in the file without Excel
- **Use OfficeAgent.NET** if: You're already using it for document editing and recalculation is secondary
- **Combine approaches** if: Different workbooks have different requirements

## Files

- `Program.cs` - Implementation of all three approaches
- `ExcelRecalculation.csproj` - Project configuration
- `README.md` - This file

## Dependencies

- `DocumentFormat.OpenXml` (3.5.1) - Official Microsoft OpenXML SDK
- `ClosedXML` (0.105.1) - Open source Excel library
- `OfficeAgent.Core` - For `SpreadsheetPartUtility.RecalculateOnOpen()`
- `OfficeAgent.Excel` - Excel format module (for inspection)

## Related Samples

- `/samples/LiabilityCap` - T2: Edit with comment preservation and change tracking
- `/samples/TableRowInsertWithTrackedChangesAndCommentsTests.cs` - Comment handling
