# Report refresh

Update a Word report or a PowerPoint deck from an Excel workbook: this month's figures in
the narrative, this month's tables, and this month's charts, with the document's formatting
left as the author made it. Link each part to a workbook range once. Every refresh after that
is one call.

## How it works

The document records what to update with a **binding marker**: `xl:` followed by a workbook
range. The marker lives where Office keeps names, so it survives editing and saving in Word
and PowerPoint.

| What updates | Where the marker lives | What a refresh writes |
| --- | --- | --- |
| A number or phrase in Word text | The tag of a content control around it | The cell's displayed value |
| A PowerPoint text box, such as a KPI tile | The shape's name in the Selection Pane | The cell's displayed value, replacing the box's text |
| A Word table | The table's alt-text title | The range, cell by cell |
| A PowerPoint table | The table's name in the Selection Pane | The range, cell by cell |
| A PowerPoint chart | The chart's name in the Selection Pane | The range's data |

A range is written as Excel writes one: `Summary!B2`, `'Monthly Sales'!A1:D13`, a defined
name such as `Revenue`, or an Excel table name such as `SalesTable`.

## Link the document once

An agent links each part with the `bindData` operation, usually in one `edit_document` call:

```json
[
  { "op": "bindData", "target": { "find": "$3.1M" }, "source": "Revenue" },
  { "op": "bindData", "target": { "kind": "table", "path": "table#0" }, "source": "'Monthly Sales'!A1:C13" }
]
```

In a deck the targets are nodes from `inspect_document`:

```json
[
  { "op": "bindData", "target": { "kind": "shape", "path": "shape#256/5" }, "source": "Summary!B2" },
  { "op": "bindData", "target": { "kind": "chart", "path": "chart#257/4" }, "source": "SalesData" }
]
```

`bindData` changes nothing a reader sees: the Word text is wrapped in a content control
whose title shows the range, and the PowerPoint shape is renamed. Inspect the workbook first
and copy its sheet names, `definedNames` and table names rather than guessing them.

To link by hand instead: in PowerPoint, rename the shape, table or chart in the Selection Pane
(Alt+F10) to `xl:Summary!B2`. In Word, set a table's alt-text title (Table Properties → Alt
Text) to `xl:Sales!A1:D6`, or add a content control from the Developer tab and set its tag to
`xl:Summary!B2`. Content-control tags must be unique, so a second link to the same cell is
written `xl:Summary!B2#2`; anything after `#` is ignored.

## Refresh it each month

```json
{ "connectionId": "reports", "source": "board/october.pptx", "workbookSource": "finance/october.xlsx", "preview": true }
```

`refresh_document` reads every marker, reads the ranges, and either previews or saves. Run
it with `preview: true` first: the result lists each binding with the range it resolved to and
what it will write. Then run it again without `preview`. Use `saveMode` `NewVersion` to keep
last month's document beside the new one. The workbook can be on another connection, named by
`workbookConnectionId`.

```json
{
  "isValid": true,
  "committed": true,
  "receipt": { … },
  "bindings": [
    { "kind": "Value", "marker": "xl:Revenue", "reference": "Revenue", "range": "Summary!B2", "summary": "$4,200,000", … },
    { "kind": "Table", "marker": "xl:'Monthly Sales'!A1:C13", "range": "'Monthly Sales'!A1:C13", "summary": "13 row(s) × 3 column(s)", … },
    { "kind": "Chart", "marker": "xl:SalesData", "range": "'Monthly Sales'!A1:C13", "summary": "12 categories, 2 series", … }
  ],
  …
}
```

A refresh is all or nothing. If any binding cannot be read - a sheet was renamed, a formula
has no stored result, a chart range holds text - nothing is written, and every problem is
listed with its marker. A report showing this month's revenue beside last month's costs is
worse than one left as it was.

From C#, `OfficeAgentClient.PlanDataRefresh(document, workbook)` returns the same bindings and
a `DocumentPlan` to preview or commit like any other.

## What each binding writes

**Values** take the cell as Excel displays it: the number format is applied, so `4200000`
under `"$"#,##0` becomes `$4,200,000` and `0.125` under `0.00%` becomes `12.50%`. Separators
are the invariant culture's. In Word the content control keeps its formatting; a PowerPoint
text box keeps its first paragraph and run formatting and takes the value as its whole text,
so put the label in a box of its own.

**Tables** are filled from the range row by row, starting top-left, and include any header
row the range has. The range must have as many columns as the table. A table with fewer rows
grows by copies of its last row, and one with more loses its last rows, so a top-ten table
can become a top-twelve. Each cell keeps its paragraph and first run's formatting. Tables
with merged cells are refused.

**Charts** take the range with series in columns: the first row names the series, the first
column holds the categories, and the rest are numbers. The chart keeps its type, colours,
labels, axes and layout. A new series copies the last series' formatting, and series beyond
the range are removed. The chart's embedded data sheet is rewritten to hold exactly the new
values, so Edit Data in PowerPoint opens what the chart shows; a link to another workbook
is replaced by that embedded sheet. Empty cells become gaps. Bar, column, line, area, pie,
doughnut and radar charts are written; scatter, bubble, surface and stock charts are refused.
Categories are written as text, so a date axis shows the dates as Excel displays them.

## Limits

- Formula results come from the workbook as saved. A workbook saved by Excel stores every
  result; one written by another program may not, and its formula cells are refused with
  `not-calculated` until it is opened and saved in Excel.
- One binding reads at most 10,000 cells.
- Word charts are not written; bind the table behind them instead, or put the chart in the
  deck.
- PowerPoint text is bound a whole text box at a time. A number inside a sentence on a slide
  needs a text box of its own; in Word, any span of text can be bound.
- The plan is bound to the document's snapshot, so a document edited between reading and
  saving is refused as `stale-snapshot`; run the refresh again.
- Error codes: `no-bindings`, `range-not-found`, `not-one-cell`, `not-calculated`,
  `not-chart-data`, `duplicate-binding` and `range-too-large`, each with a message naming the
  marker.
