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
shared-string index rather than the resolved text. OfficeAgent does not calculate a new result.

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

## Preservation and limits

Edits mutate the addressed cells, table definition, comment parts, and calculation
properties in place. Unrelated formulas, styles, worksheets, tables, and workbook parts
remain in the package. OfficeAgent does not run Excel, evaluate formulas, render charts,
refresh data connections, or expand pivot caches.

The Excel snapshot covers workbook XML, shared strings, worksheet XML, table definitions, and
legacy note XML. It does not cover styles, drawings, charts, external links, pivot parts, VBA,
or other workbook parts. Provider version checks and operation-specific validation remain the
boundary for content outside that snapshot.
