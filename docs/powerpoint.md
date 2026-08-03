# PowerPoint support

`OfficeAgent.PowerPoint` is a format module, registered exactly like the Word one.
A host that registers both serves either format from one client: the engine
routes each document to the module that can handle it, so nothing in the calling
code switches on file type.

```csharp
services
    .AddWordFormat()
    .AddPowerPointFormat()
    .AddFileSystemDocumentProvider("workspace", "/srv/officeagent/workspace", o =>
        o.AllowedExtensions = new[] { ".docx", ".pptx" })
    .AddOfficeAgent();
```

The connection's `AllowedExtensions` is what admits `.pptx`; the default is
`.docx` only.

## Addressing

A deck has no text flow, so anchors are scoped to the slide and the shape rather
than to a document-wide position.

| Anchor | Form | Notes |
| --- | --- | --- |
| Paragraph | `slide{slideId}/shape{shapeId}/p{n}` | `slide{slideId}/notes/shape{shapeId}/p{n}` for a speaker-notes body; `slide{slideId}/shape{shapeId}/r{row}c{col}/p{n}` inside a table cell |
| Slide node | `slide#{slideId}` | The target for inserting a table, image, or comment |
| Table node | `table#{slideId}/{shapeId}` | |
| Image node | `image#{slideId}/{shapeId}` | |
| Comment node | `comment#{slideId}/{commentId}` | |

`slideId` is the id from `p:sldIdLst`, which survives reordering — unlike a slide
number. Node paths are keyed by shape id rather than by ordinal for the same
reason: adding a table to an earlier slide must not silently retarget a path that
already exists.

The trailing `p{n}` is positional **within its own text body**. That is safe for
every verb below because none of them add or remove paragraphs inside a body. A
paragraph-inserting verb would need a durable id scheme first — see
[Stabilize](#anchor-stability) below.

## Supported operations

| Verb | Target | Notes |
| --- | --- | --- |
| `changeText` | paragraph | Works across slides, table cells, and notes. An empty `expect` writes into an empty paragraph — the route to filling a new deck's placeholder, since a slide has no paragraph-inserting verb. `mode: "Tracked"` is **refused** — PresentationML has no redline vocabulary |
| `format` | paragraph / image node | bold, italic, underline, `sizeHalfPoints`, `fontFamily`, `color`, `highlight`, `alignment`; `widthPx`/`heightPx` resizes an image. An empty `expect` styles the whole paragraph. Word-only measures (`styleId`, indents, spacing, borders) are **refused**, not ignored |
| `insertTable` | slide node | Placed below existing content |
| `removeTable` | table node | Removes the frame, not just the `a:tbl` |
| `insertTableRows` / `removeTableRows` | table node | `Start`/`End`/`Before`/`After`; negative indices count from the end |
| `insertTableColumns` / `removeTableColumns` | table node | Grid and rows stay in step |
| `insertImage` / `removeImage` | slide / image node | `base64Bytes` or a provider-backed `imageDocumentId`; `altText` is carried through |
| `comment` (add) | slide node | |
| `comment` (resolve) | comment node | `"action": "Resolve"` |

Verbs the module does not implement — `insert`, `fill`, `setProperty`,
`revision`, `copyStyles`, `clearStyles` — are reported per-operation as
`unsupported-operation`, and nothing in the plan is applied.

## Comments

Comments use the Office 2021 model (`p188:cm` in a `PowerPointCommentPart`),
which is what current PowerPoint writes and the only one carrying a status that
can be resolved. Legacy `p:cm` comments have no resolved state, so they are not
surfaced as resolvable.

Resolving keeps the comment and any replies and changes only its status, so the
review trail survives. Resolving an already-resolved comment is refused rather
than silently repeated.

Repeated comments by the same author share one author entry; duplicating it would
make PowerPoint list the same person once per comment.

## Creating a deck

The module implements `IBlankDocumentFactory` for `.pptx`, so `create_document`
and `OfficeAgentClient.CreateAsync` work for decks — the extension of the
requested name selects the format:

```csharp
var deck = await client.CreateAsync("workspace", "review.pptx");
```

A new deck is one slide with an empty title placeholder, addressable as
`slide256/shape2/p0`, over the slide master, layout, and theme a presentation
needs in order to open at all. That anchor is the PowerPoint counterpart of a
blank Word document's `auto-0000`.

Fill it with `changeText` and an empty `expect`:

```jsonc
[ { "op": "changeText",
    "target": { "paraId": "slide256/shape2/p0", "expect": "" },
    "with": "Quarterly Review", "mode": "Direct" } ]
```

The empty `expect` is still content-verified — it asserts the paragraph *is*
blank, so if the deck drifted and something is already there the operation fails
with `expect-mismatch` rather than overwriting it. Where a blank Word document
takes an `insert`, a deck takes this, because slides have no paragraph-inserting
verb.

## Plans do not need a format

`DocumentPlan.Format` defaults to `Unspecified`, so the plan shape the tools
document — a bare list of operations — applies to whichever format the document
turns out to be. Set `format` only to *assert* a format; a mismatch then fails
the plan with `contract-mismatch`.

## What is preserved

- **Character formatting.** Text replacement isolates the matched span through
  the shared text engine, so runs the span does not fully cover keep their
  formatting. Text that runs across several `a:r` elements is replaced as one
  edit.
- **Shape properties and layout.** Nothing rewrites a shape's transform, style,
  or placeholder relationship; edits touch the text body, the table, or the
  shape tree only.
- **Slide layouts and masters.** Untouched. New shapes inherit from the layout
  the slide already points at.

## Anchor stability

`Stabilize` is deliberately a no-op returning an empty alias map. Word assigns
`w14:paraId` there so that an operation which shifts paragraph offsets cannot
redirect a later operation's target. DrawingML has no equivalent identifier to
mint, and none of the verbs above add or remove paragraphs inside a text body, so
offsets cannot shift mid-plan.

**If a paragraph-inserting verb is added to this module, that reasoning stops
holding** and a durable id scheme is required first — an empty alias map would
silently let one operation retarget another.

## One plan, one operation per target

The shared conflict rule keys an operation on its verb, its anchor, and (for positional
verbs) its slot — not on its payload. Two operations of the same verb on the same target
are therefore refused with `operation-conflict` even when they carry different content:

```jsonc
// Refused: both append to the same table, so both key the same.
[ { "op": "insertTableRows", "target": { "kind": "table", "path": "table#256/3" }, "rows": [["A","1"]] },
  { "op": "insertTableRows", "target": { "kind": "table", "path": "table#256/3" }, "rows": [["B","2"]] } ]
```

Merge them into one operation (`"rows": [["A","1"],["B","2"]]`), give them distinct
positions, or send separate plans.

This is shared engine behaviour rather than something specific to decks, and it is partly
protective: index-positioned operations are validated against the *pre-apply* document, so
a second operation's `rowIndex` would otherwise be read against a table the first one had
already resized.

## Known gaps

- Slide-level verbs (add, delete, reorder, duplicate slides) are not implemented.
- `format` does not reach shape fills, outlines, or table styles; it covers run and paragraph formatting plus image size.
- Legacy `p:cm` comments are neither read nor written.
- Charts and SmartArt are not addressable.
