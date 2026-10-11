# Excel support

Install `OfficeAgent.Excel` beside the core package and register it with
`AddExcelFormat()`. The standalone `OfficeAgent.Mcp` tool already registers it.

```csharp
services.AddExcelFormat().AddOfficeAgent();
```

Inspection returns `worksheets`, Excel table nodes, cell-note nodes, and populated
`cells`. A cell target uses the workbook's durable numeric sheet id plus an A1 address:

```json
{ "sheetId": 7, "address": "B2" }
```

Set `InspectOptions.SheetId` and `Range` to inspect one rectangle. `MaximumCells`
defaults to 1,000 and is capped at 10,000. Inspection enumerates stored cell elements within
the range; truly absent cells are omitted. `find` can search `Displayed`, `Raw`, or `Both`.
Displayed text resolves inline and shared strings and uses a formula's stored cached result;
it does not apply Excel number formats. Raw search uses the stored cell value, such as the
shared-string index rather than the resolved text. Unless [formula calculation](#formula-calculation)
is turned on, OfficeAgent does not calculate a new result.

## Operations

`setCell` changes only the target cell's value/formula fields and keeps its style. Auto
values recognize invariant numbers and Booleans; set `valueKind` to `String` to preserve
text such as a leading-zero identifier. A
formula may omit the leading `=`. OfficeAgent removes the stale cached value unless one
is supplied explicitly, removes the calculation chain, and marks the workbook for a full
recalculation when Excel opens it. Any `setCell` write requests recalculation because a changed
constant can affect formulas elsewhere. Replacing one member of a multi-cell shared-formula
group or legacy array-formula range is rejected; expand it into independent formulas before
editing one member.

Functions added to Excel after the 2007 file format, such as `STDEV.S`, `TEXTJOIN`, `XLOOKUP` and
`FILTER`, are stored with the prefix Excel writes for them (`_xlfn.STDEV.S`, `_xlfn._xlws.FILTER`),
so write them as you would type them in Excel. Inspection reports formulas as stored, prefix
included, as it does for workbooks Excel saved. The parameter names of `LET` and `LAMBDA` are not
rewritten; Excel stores them as `_xlpm.name`.

```json
{ "op": "setCell", "target": { "sheetId": 7, "address": "B2" }, "formula": "SUM(B3:B8)" }
```

`appendTableRows` targets a `spreadsheetTable` node from inspection. Each row must match
the table's column count. The operation copies the last row's cell styles, grows the table
and auto-filter ranges, and refuses to overwrite populated cells below the table. Tables
with totals rows are currently refused.

```json
{ "op": "appendTableRows", "target": { "kind": "spreadsheetTable", "path": "table#7/Sales" }, "rows": [["APAC", "15"]] }
```

The shared `comment` verb manages legacy Excel cell notes. `Add` targets a cell;
`Remove` can target the `cellComment` node returned by inspection. Replies and resolved
status are unavailable in the legacy note model.

`setProperty` sets the workbook's core properties, the same `core/*` paths a Word document
and a deck take. Inspection lists them as `docProperty` nodes with their current values.

```json
{ "op": "setProperty", "target": { "kind": "docProperty", "path": "core/title" }, "value": "Q2 Sales" }
```

## Formula calculation

A formula written by `setCell` has no cached value until Excel opens the workbook and
recalculates. A reader that does not calculate - a file previewer, a data pipeline,
`inspect_document` itself - sees an empty cell. The optional `OfficeAgent.Excel.Calculation`
package closes that gap for the functions it is verified to calculate correctly:

```csharp
services.AddExcelFormat().AddExcelFormulaCalculation().AddOfficeAgent();

using var result = await client.ApplyAsync(handle, plan,
    new ApplyOptions { DryRun = false, CalculateFormulas = true });
```

A preview (`DryRun = true`) with `CalculateFormulas` reports the results without writing
anything. A connection can turn calculation on for every plan that passes through it with
`CalculateFormulas` on its provider options; the [MCP server](mcp-server.md#configuration-reference)
exposes the same setting per connection. Without a registered evaluator, `CalculateFormulas`
does nothing and `Report.Calculation` stays `null`.

**What is calculated.** A formula is calculated only when every function it calls is on the
engine's allowlist - `describe_capabilities` returns it under `formulaCalculation` - and it
depends on no formula that was left uncalculated. The allowlist covers common arithmetic,
statistics, logic, text, lookup (`VLOOKUP`, `HLOOKUP`, `INDEX`, `MATCH`), date and
financial functions. Each one is checked against a known value in the test suite.

**What is not.** These are reported as `notCalculated` with the reason, never given a value:

- volatile functions (`NOW`, `TODAY`, `RAND`, `OFFSET`, `INDIRECT`, …), whose value
  would be stale the moment it was written;
- structured table references (`Sales[Amount]`) and links to other workbooks - the
  underlying engine evaluates a structured reference to 0 without an error;
- array formulas, names the workbook does not define, and any function off the allowlist
  (`XLOOKUP`, `FILTER`, `LAMBDA`, …);
- every formula that depends, directly or indirectly, on one of the above.

A circular reference leaves the whole workbook uncalculated and sets `circularReference`.

**What is written.** Only the cached value (`<v>`) and its type of each calculated formula
cell change; worksheets with no changed value are not rewritten. The workbook stays marked
for a full recalculation, so Excel still recalculates everything when it opens the file.
Text results are formatted with invariant (en-US style) conventions whatever the server's
locale, so `TEXT(1234.5,"#,##0.00")` is cached as `1,234.50`.

**Formula errors.** `Report.Calculation.NewErrors` lists formulas that now evaluate to an error
(`#DIV/0!`, `#REF!`, `#N/A`, …) that the input did not already carry at that cell. By default
the plan still commits and the errors are reported. With `RejectFormulaErrors` a new error
or a circular reference refuses the plan with `formula-error` and nothing is written.

The engine is ClosedXML, pinned to its 0.105 line; the allowlist moves only with the tests.

## Preservation and limits

Edits mutate the addressed cells, table definition, comment parts, and calculation
properties in place. Unrelated formulas, styles, worksheets, tables, and workbook parts
remain in the package. OfficeAgent does not run Excel, render charts, refresh data
connections, or expand pivot caches, and evaluates formulas only as described under
[formula calculation](#formula-calculation).

The Excel snapshot covers workbook XML, shared strings, worksheet XML, table definitions, and
legacy note XML. It does not cover styles, drawings, charts, external links, pivot parts, VBA,
or other workbook parts. Provider version checks and operation-specific validation remain the
boundary for content outside that snapshot.
