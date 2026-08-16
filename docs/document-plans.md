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
- `op` is the verb discriminator. It may appear anywhere in the object - property order does not matter - and an unknown verb comes back as `invalid-json` naming the ones that exist.
- `target` is an anchor. The `$anchor` field is optional - the engine infers the anchor type from the property names (`paraId` → text span, `tag` → structural, `kind`/`path` → node).
- `contractVersion` and `snapshot` are optional. Omit unless you need explicit drift detection.
- `format` is optional and *asserts* the document's format (`"Word"`, `"PowerPoint"`); a mismatch fails the plan with `contract-mismatch`. The verb vocabulary is shared, so a plan that does not care works against either.

The anchor shapes and node paths below are the Word ones. A deck addresses slides, shapes, and speaker notes instead - see [PowerPoint support](powerpoint.md#addressing).

## Anchors

| Shape | Type |
| --- | --- |
| `{ "paraId": "…", "expect": "…", "occurrence": 0 }` | text span in a paragraph |
| `{ "tag": "ClientName" }` | content control or bookmark |
| `{ "kind": "<nodeKind>", "path": "<path>" }` | document-level node - `kind` is `"table"`, `"tableRow"`, `"tableCell"`, `"image"`, `"docProperty"`, `"revision"`, or `"field"`, *not* the literal string `"node"` |

Table-row, table-cell, and image paths come from `inspect_document.nodes`:

| Path | Addresses |
| --- | --- |
| `table#N` | the N-th table |
| `table#N/row#M` | row M of table N (used by `format` against a row) |
| `table#N/cell#R/C` | cell at row R, column C |
| `image#N` | the N-th inline image |

## Supported verbs

A text-span target resolves across the body, headers, footers, footnotes, and endnotes.

A paragraph's text is everything the reader sees, including text wrapped in a tracked
insertion (`w:ins`), a hyperlink, or a content control. So a redline pass can be revisited:
change something with `"mode": "Tracked"`, re-open, and the text you inserted is a normal
target for `find`, `changeText` and `format`. Text struck through by a tracked deletion is
not part of it, and a text box's paragraphs belong to their own body rather than to the
paragraph that carries them.

### `changeText`

Replace a content-verified text span. An operation that does not state a `mode` takes the
connection's `DefaultChangeMode`, which is `Tracked` unless the host configured otherwise -
see [document providers](document-providers.md#default-change-mode). An operation that does
state one is never overridden.

```json
{ "op": "changeText",
  "target": { "paraId": "w14:…", "expect": "Acme Corp", "occurrence": 0 },
  "with":   "Globex Inc.",
  "mode":   "Tracked" }
```

`expect` may be empty **only when the target paragraph is itself empty**, which means "write
here" — the way the first words go into a document `create_document` just made, and into a
new slide's placeholder. Against a paragraph that has text an empty `expect` is refused: it
cannot be told apart from a caller who left the field out, and the cost of guessing wrong is
a rewritten paragraph nobody named. Under `Tracked`, filling an empty paragraph is recorded
as an insertion.

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
| `indentLeftTwips`, `indentRightTwips`, `indentFirstLineTwips` | paragraph indent (1 inch = 1440 twips). A **negative** `indentFirstLineTwips` is a hanging indent — the first line set back from the rest, how a bullet hangs its dash outside the text |
| `listStyle`, `listLevel`, `listId` | makes the paragraph a real list item (Word only). `bullet`, `decimal` (1. / a. / i.), `clause` (1. / 1.1 / 1.1.1), or `none` to remove it. `listLevel` is 0–8. Paragraphs sharing a style **and** a `listId` form one running sequence; a different `listId` starts a separate one, which is how a second chapter restarts its steps at 1. Word owns the numbers, so inserting an item renumbers the rest — text that merely begins "4.2" does not |
| `pageBreakBefore` | the paragraph starts a new page (Word only). A property of the paragraph, not a character in the text, so the page goes on starting there as the text above it is edited |
| `spacingBeforeTwips`, `spacingAfterTwips` | paragraph spacing |
| `borderStyle`, `borderSizeEighths`, `borderColor` | paragraph / table / cell border (8 = 1 pt) |
| `borderEdges` | which edges the border is drawn on — a comma-separated subset of `top`, `left`, `bottom`, `right` and, on a table, `insideH`, `insideV`. Unset means every edge. `"left"` is a pull quote's rule; all four is a callout box |
| `widthPx`, `heightPx` | image and row sizing at 96 DPI |
| `xPx`, `yPx` | shape position at 96 DPI (PresentationML shape targets) |
| `fillColor`, `lineColor` | hex RGB, or `none`, on a shape or slide background (PresentationML) |
| `lineWidthPx` | shape outline width at 96 DPI (PresentationML) |
| `verticalAlignment` | `top`, `middle`, `bottom` — where text sits in a shape's box (PresentationML) |

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

`scope` is `run`, `paragraph`, or `all`. Empty `expect` means the whole paragraph. Only *direct* formatting travels or is removed - a deck's layout and master, and a Word document's style definitions, are never touched, so clearing returns the text to the look its template gives it.

### `fill`

Populate a named slot by tag. Slot tags come from `inspect_document.contentControls`, each carrying the `kind` that says what it is.

```json
{ "op": "fill",
  "target": { "tag": "ClientName" },
  "value":  "Globex" }
```

In Word a slot is a content control or bookmark (`kind: "contentControl"`). On a deck it is the **shape name** a template author sets, which PowerPoint shows in its Selection Pane (`kind: "shapeName"`); names repeat across slides, so a tag matching more than one shape is refused as `ambiguous-anchor` rather than filling an arbitrary one, and is qualified as `slide256/ClientName`.

**`fill` populates a slot; it does not create one.** In Word that means the document must already carry the content control — no verb adds one, so `fill` applies to templates the host supplies rather than to documents `create_document` produced. A deck has no such gap: every named shape is a slot, and `insertShape` makes one.

### `comment`

Attach a review comment.

```json
{ "op": "comment",
  "target": { "paraId": "w14:…", "expect": "Acme Corp", "occurrence": 0 },
  "text":   "Confirm the counterparty name.",
  "author": "Reviewer", "initials": "R" }
```

### `insert`

Insert a new paragraph near an anchor paragraph. Use `insertTable` to add a *new* table, or `insertTableRows` / `insertTableColumns` to extend an existing one.

```json
{ "op": "insert",
  "target":   { "paraId": "w14:…", "expect": "…" },
  "position": "After",
  "text":     "New paragraph." }
```

On a deck this adds a bullet or line beside an existing one, inheriting the neighbour's bullet and run styling. `level` (0-8) sets the bullet depth there and is refused in Word, where numbering comes from the paragraph style; `styleId` is the reverse. Because a slide paragraph id is positional, a plan that inserts and then addresses the same text body at an equal or higher index is refused — see [anchor stability](powerpoint.md#anchor-stability).

### `insertTable` / `removeTable`

Insert a whole new table near an anchor paragraph, or remove an entire table addressed by its `table#N` path.

```json
{ "op": "insertTable",
  "target":   { "paraId": "w14:…", "expect": "…" },
  "position": "After",
  "table":    { "headers": ["Country", "Population"],
                "rows":    [["US", "332"], ["UK", "68"]] } }

{ "op": "removeTable",
  "target":   { "kind": "table", "path": "table#0" } }
```

`removeTable` deletes the table and all of its rows.

### `insertTableRows` / `removeTableRows`

Remove rows by explicit `rowIndices` (negative counts from the end), or set `onlyIfEmpty` to drop only blank rows.

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
  "rowIndices":  [-1, -2] }

{ "op": "removeTableRows",
  "target":      { "kind": "table", "path": "table#0" },
  "onlyIfEmpty": true }
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
- `imageConnectionId` + `imageDocumentId` - the opaque id of an image already registered with a provider connection (use a connection whose `AllowedExtensions` permits image extensions). The client fetches the bytes through the provider before the plan reaches the engine.

Inline base64:

```json
{ "op": "insertImage",
  "target":      { "paraId": "w14:…", "expect": "…" },
  "base64Bytes": "iVBORw0KGgo…",
  "imageType":   "png",
  "widthPx": 200, "heightPx": 80,
  "position": "After",
  "altText":  "Company logo" }
```

By opaque id resolved through a provider:

```json
{ "op": "insertImage",
  "target":            { "paraId": "w14:…", "expect": "…" },
  "imageConnectionId": "images",
  "imageDocumentId":   "5f2c1a9b8e0d4f7a",
  "imageType":         "png",
  "widthPx": 200, "heightPx": 80,
  "position": "After",
  "altText":  "Company logo" }
```

Remove by node anchor:

```json
{ "op": "removeImage",
  "target": { "kind": "image", "path": "image#0" } }
```

### `backgroundImage`

An image *behind* the content, as opposed to `insertImage`'s picture *in* it. On a deck a
slide target paints that slide and no target paints every slide; in Word there is no target
and the image repeats on every page.

```json
{ "op": "backgroundImage",
  "base64Bytes": "iVBORw0KGgo…", "imageType": "png", "opacity": 0.2 }

{ "op": "backgroundImage",
  "target": { "kind": "slide", "path": "slide#256" },
  "imageConnectionId": "images", "imageDocumentId": "<id from a prior add>",
  "opacity": 0.15 }
```

`opacity` runs 0 to 1 and defaults to 1. Set it for any photograph that has text over it:
at full strength almost any image destroys the contrast the text needs, and 0.1–0.3 is the
usable range. It is written as `a:alphaModFix` on the blip — the same attribute PowerPoint's
own Transparency slider reads.

Supplying neither `base64Bytes` nor `imageDocumentId` takes an existing background away. A
flat colour is [`format`](#format) with `fillColor`, not this.

Word has no usable page-background element — its own is a lump of VML that does not print by
default — so the image is anchored page-sized and behind the text in the section's header,
which is what Word's designed templates do. It therefore lands in the first-page and
even-page headers too when the section uses them, and `insertImage` remains the way to put a
picture in the text flow.

### `headerFooter`

Running heads and page numbers. Word only in this section; the deck's own footer settings are
in [powerpoint.md](powerpoint.md).

```json
{ "op": "headerFooter",
  "header": "Northwind Traders — Q2 Board Review",
  "footer": "Confidential",
  "showPageNumber": true,
  "alignment": "edges",
  "differentFirstPage": true }

{ "op": "headerFooter", "scope": "firstPage", "header": "", "footer": "" }
```

| Property | Notes |
| --- | --- |
| `header`, `footer` | the text. An empty string clears it |
| `showPageNumber` | a `PAGE` field, not a number, so it stays right as the document grows |
| `differentFirstPage` | gives page one its own header and footer — how a cover keeps the running head off it |
| `scope` | which pages this writes: `default` (the fallback), `firstPage`, `evenPage` |
| `alignment` | `left` (default), `center`, `right`, or `edges` — text left and page number right on one line |

The deck-only settings (`showSlideNumber`, `showFooter`, `showDateTime`, `dateTime`) are
**refused** here rather than ignored.

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

`updateOnOpen` sets `w:updateFields`, which makes Word ask *"This document contains fields that may refer to other files. Do you want to update the fields in this document?"* every time the document is opened. On a document with **no fields** that prompt buys nothing and reads to the user as a damage warning, so it is refused with `invalid-operation` rather than armed. Set it once the document actually contains a field.

### `headerFooter` / `insertMedia`

PowerPoint only; see [PowerPoint support](powerpoint.md#footer-slide-number-and-date) and [embedded media](powerpoint.md#embedded-video-and-audio).

```json
{ "op": "headerFooter",
  "footer": "Confidential", "showSlideNumber": true, "showDateTime": true }

{ "op": "insertMedia",
  "target":      { "kind": "slide", "path": "slide#257" },
  "kind":        "Video",
  "mediaType":   "mp4",
  "base64Bytes": "AAAAIGZ0eXBpc29t…",
  "widthPx": 480, "heightPx": 270 }
```

`headerFooter` with no target applies to every slide. A slide has no header - that is a notes and handout concept - so one is refused rather than written where nothing renders it.

### `transition` / `animate`

PowerPoint only; see [transitions and animations](powerpoint.md#transitions-and-animations).

```json
{ "op": "transition", "effect": "push", "direction": "up", "durationMs": 700 }

{ "op": "animate",
  "target":     { "kind": "shape", "path": "shape#257/2" },
  "effect":     "fade",
  "kind":       "Entrance",
  "trigger":    "OnClick",
  "durationMs": 600 }
```

An untargeted `transition` applies to every slide. `trigger` is `OnClick`, `WithPrevious` or `AfterPrevious`, and effects play in the order the operations are sent. `effect: "none"` removes a transition or a shape's animations. Fly-in, zoom, grow and motion paths are refused - they need interpolated properties rather than a filter.

### `section`

PowerPoint only. Named slide groups; see [PowerPoint support](powerpoint.md#sections).

```json
{ "op": "section", "action": "Add", "name": "Financials",
  "target": { "kind": "slide", "path": "slide#257" } }

{ "op": "section", "action": "Rename", "name": "FY27 Financials",
  "target": { "kind": "section", "path": "section#{GUID}" } }

{ "op": "section", "action": "Remove",
  "target": { "kind": "section", "path": "section#{GUID}" } }
```

Removing a section keeps its slides - they join the section before it.

### `insertShape` / `removeShape`

PowerPoint only. `insertShape` adds a free-standing text box to a slide; `removeShape` deletes any shape by its node path — text box, table frame, or picture. Removing a layout *placeholder* is refused, because the layout would re-offer it as an empty prompt and the slide would look unchanged with its content gone.

```json
{ "op": "insertShape",
  "target":  { "kind": "slide", "path": "slide#257" },
  "text":    [ "Draft - not for circulation" ],
  "xPx": 40, "yPx": 620, "widthPx": 420, "heightPx": 50 }

{ "op": "removeShape", "target": { "kind": "shape", "path": "shape#257/4" } }
```

Moving and resizing go through `format` on the same shape node — see [PowerPoint support](powerpoint.md#supported-operations).

### `insertSlide` / `removeSlide` / `moveSlide` / `duplicateSlide`

PowerPoint only; a Word document reports them as `unsupported-operation`. Several `insertSlide` operations in one plan author a whole deck — see [PowerPoint support](powerpoint.md#generating-a-deck) for the layouts and what a slide inherits from them.

```json
{ "op": "insertSlide",
  "slide": { "layout": "titleAndContent",
             "title":  "FY27 Priorities",
             "body":   [ "Finish the migration", "Rebuild the pipeline" ],
             "notes":  "Do not commit to a date." } }

{ "op": "moveSlide",
  "target":     { "kind": "slide", "path": "slide#259" },
  "position":   "After",
  "relativeTo": "slide#256" }

{ "op": "duplicateSlide", "target": { "kind": "slide", "path": "slide#259" } }

{ "op": "removeSlide", "target": { "kind": "slide", "path": "slide#258" } }
```

`position` is `Start`, `End`, `Before`, or `After`. `insertSlide` defaults to `End` and takes its reference slide from `target`; `moveSlide` and `duplicateSlide` target the slide they act on and name the reference in `relativeTo`. `duplicateSlide` defaults to landing immediately after the original.

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
| `stale-snapshot` | The plan's format-specific text-host snapshot does not match live content. Re-inspect and rebuild. |
| `anchor-not-found` | The target anchor cannot be resolved in the live document. |
| `expect-mismatch` | The live content no longer matches the anchor's `expect`. |
| `ambiguous-anchor` | The target is not specific enough to edit safely. |
| `unsupported-operation` | No registered handler supports the verb / anchor combination. |
| `invalid-operation` | The operation is structurally invalid (e.g. empty `expect` against a paragraph that has text, no formatting properties). |
| `requires-renderer` | The requested change needs a layout / calculation engine. |
| `operation-conflict` | Two operations target the same location in one plan. |
| `contract-mismatch` | The plan's contract version does not match the engine, or the plan asserted a `format` the document is not. (The version check is informational pre-1.0; the format assertion is not.) |
| `invalid-json` | The plan is not valid JSON, or an operation names no known verb. The message lists the verbs that exist. |
| `invalid-argument` | A tool argument outside the plan is wrong - an unrecognised `saveMode`, for example. |

Provider boundary errors (`apply_plan` only) use a separate set of wire codes - see [document-providers.md](document-providers.md).
