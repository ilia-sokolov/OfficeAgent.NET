# Document plans

A `DocumentPlan` is the wire contract between an agent and the engine. It is a JSON-serialisable list of typed operations against anchors returned by `inspect_document` / `find_in_document`.

## Shape

```json
{
  "operations": [
    {
      "op": "changeText",
      "target": { "paraId": "w14:00000002", "expect": "Acme Corp", "occurrence": 0 },
      "with":   "Globex Inc.",
      "mode":   "Tracked"
    }
  ]
}
```

- `operations` is required. Each entry is one operation.
- `op` is the verb discriminator.
- `target` is an anchor. The `$anchor` field is optional - the engine infers the anchor type from the property names (`paraId` → text span, `tag` → structural, `kind`/`path` → node).
- `contractVersion` and `snapshot` are optional. Omit unless you need explicit drift detection.

## Anchors

| Shape | Type |
| --- | --- |
| `{ "paraId": "…", "expect": "…", "occurrence": 0 }` | text span in a paragraph |
| `{ "tag": "ClientName" }` | content control or bookmark |
| `{ "kind": "<nodeKind>", "path": "<path>" }` | document-level node - `kind` is `"table"`, `"docProperty"`, `"revision"`, or `"image"`, *not* the literal string `"node"` |

Table-row, table-cell, and image paths come from `inspect_document.nodes`:

| Path | Addresses |
| --- | --- |
| `table#N` | the N-th table |
| `table#N/row#M` | row M of table N (used by `format` against a row) |
| `table#N/cell#R/C` | cell at row R, column C |
| `image#N` | the N-th inline image |

## Supported verbs

Concrete JSON shapes follow. Every verb's target works across the body, headers, footers, footnotes, and endnotes.

### `changeText`

Replace a content-verified text span. Default mode is `Tracked`.

```json
{ "op": "changeText",
  "target": { "paraId": "w14:…", "expect": "Acme Corp", "occurrence": 0 },
  "with":   "Globex Inc.",
  "mode":   "Tracked" }
```

### `format`

Unified styling. Apply any combination of properties to a paragraph or text span, a table, a table row, a table cell, or an image. Properties left unset are not changed.

```json
{ "op": "format",
  "target": { "paraId": "w14:…", "expect": "important", "occurrence": 0 },
  "highlight": "yellow", "bold": true, "color": "FF0000" }

{ "op": "format",
  "target": { "paraId": "w14:…", "expect": "" },
  "styleId": "Heading2", "alignment": "center" }

{ "op": "format",
  "target": { "kind": "table", "path": "table#0" },
  "styleId": "TableGrid", "borderStyle": "single",
  "borderSizeEighths": 4, "borderColor": "000000" }

{ "op": "format",
  "target": { "kind": "tableRow",  "path": "table#0/row#0" },
  "bold": true, "highlight": "yellow" }

{ "op": "format",
  "target": { "kind": "tableCell", "path": "table#0/cell#1/2" },
  "alignment": "right", "color": "0000FF" }

{ "op": "format",
  "target": { "kind": "image", "path": "image#0" },
  "widthPx": 320, "heightPx": 200 }
```

Properties (all optional):

| Property | Notes |
| --- | --- |
| `styleId` | named paragraph style (paragraph targets) or table style (table targets) |
| `fontFamily`, `sizeHalfPoints` | character size: 24 = 12 pt |
| `bold`, `italic`, `underline` | character toggles |
| `highlight` | `yellow`, `green`, `cyan`, `magenta`, `blue`, `red`, `dark*`, `lightGray`, `black`, `white`, `none` |
| `color` | hex RGB font colour, e.g. `FF0000` |
| `alignment` | `left`, `center`, `right`, `justify` |
| `indentLeftTwips`, `indentRightTwips`, `indentFirstLineTwips` | paragraph indent (1 inch = 1440 twips) |
| `spacingBeforeTwips`, `spacingAfterTwips` | paragraph spacing |
| `borderStyle`, `borderSizeEighths`, `borderColor` | paragraph / table / cell border (8 = 1 pt) |
| `widthPx`, `heightPx` | image and row sizing at 96 DPI |

### `copyStyles` / `clearStyles`

Copy direct formatting from one text span to another, or clear it on the target.

```json
{ "op": "copyStyles",
  "source": { "paraId": "w14:…", "expect": "" },
  "target": { "paraId": "w14:…", "expect": "" },
  "scope":  "all" }

{ "op": "clearStyles",
  "target": { "paraId": "w14:…", "expect": "important" },
  "scope":  "run" }
```

`scope` is `run`, `paragraph`, or `all`. Empty `expect` means the whole paragraph.

### `fill`

Populate a content control or bookmark by tag.

```json
{ "op": "fill",
  "target": { "tag": "ClientName" },
  "value":  "Globex" }
```

### `comment`

Attach a review comment.

```json
{ "op": "comment",
  "target": { "paraId": "w14:…", "expect": "Acme Corp", "occurrence": 0 },
  "text":   "Confirm the counterparty name.",
  "author": "Reviewer", "initials": "R" }
```

### `insert`

Insert a new paragraph or *new* table near an anchor paragraph. Use `insertTableRows` / `insertTableColumns` to extend an existing table.

```json
{ "op": "insert",
  "target":   { "paraId": "w14:…", "expect": "…" },
  "position": "After",
  "text":     "New paragraph." }

{ "op": "insert",
  "target":   { "paraId": "w14:…", "expect": "…" },
  "position": "After",
  "table":    { "headers": ["Country", "Population"],
                "rows":    [["US", "332"], ["UK", "68"]] } }
```

### `insertTableRows` / `removeTableRows`

```json
{ "op": "insertTableRows",
  "target":   { "kind": "table", "path": "table#0" },
  "rows":     [["NL", "17"], ["GR", "10"]],
  "position": "End" }

{ "op": "insertTableRows",
  "target":   { "kind": "table", "path": "table#0" },
  "rows":     [["Header", "Header"]],
  "position": "Before", "rowIndex": 0 }

{ "op": "removeTableRows",
  "target":      { "kind": "table", "path": "table#0" },
  "rowIndices":  [-1, -2] }                  // explicit indices (negative = from end)

{ "op": "removeTableRows",
  "target":      { "kind": "table", "path": "table#0" },
  "onlyIfEmpty": true }                      // safe cleanup of blank rows
```

### `insertTableColumns` / `removeTableColumns`

`columns` is column-major: one inner list per new column, one entry per row (header first).

```json
{ "op": "insertTableColumns",
  "target":   { "kind": "table", "path": "table#0" },
  "columns":  [["Capital", "Washington", "London"]],
  "position": "End" }

{ "op": "removeTableColumns",
  "target":        { "kind": "table", "path": "table#0" },
  "columnIndices": [-1] }
```

### `insertImage` / `removeImage`

Add an inline image and remove one by `image#N` (discover paths via `inspect_document.nodes`).

The image bytes come from one of two routes - **exactly one** must be set:

- `base64Bytes` - the image inline as base64.
- `imageConnectionId` + `imageDocumentId` - the opaque id of an image already registered with a provider connection (use a connection whose `AllowedExtensions` permits image extensions). The client fetches the bytes through the provider before the plan reaches the engine, so the LLM still never handles paths.

```json
// Inline base64.
{ "op": "insertImage",
  "target":      { "paraId": "w14:…", "expect": "…" },
  "base64Bytes": "iVBORw0KGgo…",
  "imageType":   "png",
  "widthPx": 200, "heightPx": 80,
  "position": "After",
  "altText":  "Company logo" }

// By opaque id resolved through a provider.
{ "op": "insertImage",
  "target":            { "paraId": "w14:…", "expect": "…" },
  "imageConnectionId": "images",
  "imageDocumentId":   "5f2c1a9b8e0d4f7a",
  "imageType":         "png",
  "widthPx": 200, "heightPx": 80,
  "position": "After",
  "altText":  "Company logo" }

// Remove by node anchor.
{ "op": "removeImage",
  "target": { "kind": "image", "path": "image#0" } }
```

### `setProperty`

Update document metadata or a selected document-level setting.

```json
{ "op": "setProperty",
  "target": { "kind": "docProperty", "path": "core/title" },
  "value":  "Service Agreement v2" }

{ "op": "setProperty",
  "target": { "kind": "field" },
  "name":   "updateOnOpen" }
```

### `revision`

Accept or reject tracked revisions.

```json
{ "op": "revision",
  "target": { "kind": "revision", "path": "all" },
  "action": "Accept" }

{ "op": "revision",
  "target": { "kind": "revision", "path": "ins#5" },
  "action": "Reject" }
```

## Validation errors

Returned by `preview_plan` and `apply_plan` in the `errors` array. Stable wire codes.

| Code | Meaning |
| --- | --- |
| `stale-snapshot` | The plan's snapshot does not match the live document. Re-inspect and rebuild. |
| `anchor-not-found` | The target anchor cannot be resolved in the live document. |
| `expect-mismatch` | The live content no longer matches the anchor's `expect`. |
| `ambiguous-anchor` | The target is not specific enough to edit safely. |
| `unsupported-operation` | No registered handler supports the verb / anchor combination. |
| `invalid-operation` | The operation is structurally invalid (e.g. empty `expect`, no formatting properties). |
| `requires-renderer` | The requested change needs a layout / calculation engine. |
| `operation-conflict` | Two operations target the same location in one plan. |

Provider boundary errors (`apply_plan` only) use a separate set of wire codes - see [document-providers.md](document-providers.md).
