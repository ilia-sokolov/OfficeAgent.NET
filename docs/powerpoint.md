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
| Slide node | `slide#{slideId}` | The target for inserting a table, image, text box, or comment |
| Shape node | `shape#{slideId}/{shapeId}` | Any shape by identity — what moving, resizing and deleting address |
| Table node | `table#{slideId}/{shapeId}` | |
| Image node | `image#{slideId}/{shapeId}` | |
| Comment node | `comment#{slideId}/{commentId}` | |

A picture appears twice — as `image#…` and as `shape#…` — because the verbs care
about different things: one about the picture, the other about the box it sits in.

`slideId` is the id from `p:sldIdLst`, which survives reordering — unlike a slide
number. Node paths are keyed by shape id rather than by ordinal for the same
reason: adding a table to an earlier slide must not silently retarget a path that
already exists.

The trailing `p{n}` is positional **within its own text body**, and `insert` adds
paragraphs — so inserting renumbers every later paragraph in that same body. The
module refuses a plan that would then address the body positionally rather than
guessing which line was meant; see [Anchor stability](#anchor-stability).

## Supported operations

| Verb | Target | Notes |
| --- | --- | --- |
| `changeText` | paragraph | Works across slides, table cells, and notes. An empty `expect` writes into an empty paragraph — the route to filling a new deck's placeholder. `mode: "Tracked"` is **refused** — PresentationML has no redline vocabulary |
| `format` | paragraph / image node / shape node | bold, italic, underline, `sizeHalfPoints`, `fontFamily`, `color`, `highlight`, `alignment`; `widthPx`/`heightPx` resizes an image; on a **shape** node `xPx`/`yPx`/`widthPx`/`heightPx` move and resize anything - text box, table frame, picture - and nothing else is accepted there. An empty `expect` styles the whole paragraph. Word-only measures (`styleId`, indents, spacing, borders) are **refused**, not ignored |
| `insertTable` | slide node | Placed below existing content |
| `removeTable` | table node | Removes the frame, not just the `a:tbl` |
| `insertTableRows` / `removeTableRows` | table node | `Start`/`End`/`Before`/`After`; negative indices count from the end |
| `insertTableColumns` / `removeTableColumns` | table node | Grid and rows stay in step |
| `insertImage` / `removeImage` | slide / image node | `base64Bytes` or a provider-backed `imageDocumentId`; `altText` is carried through |
| `comment` (add) | slide node | |
| `comment` (resolve) | comment node | `"action": "Resolve"` |
| `insert` | paragraph | Adds a bullet or line beside an existing one. Inherits the neighbour's bullet and run styling; `level` (0-8) sets the depth. `styleId` is refused - a deck has no style table |
| `insertShape` | slide node | A free-standing text box with its own `xPx`/`yPx`/`widthPx`/`heightPx` |
| `removeShape` | shape node | Any shape - text box, table frame, picture. Removing a *placeholder* is refused |
| `insertSlide` | none, or slide node | Adds a slide from one of the deck's layouts. `position` defaults to `End`; `Before`/`After` take a slide target |
| `removeSlide` | slide node | Takes the slide's notes with it. Refused when it would empty the deck |
| `moveSlide` | slide node | `position` + `relativeTo` |
| `duplicateSlide` | slide node | Copy gets its own slide id, shape ids, and notes; lands after the original by default |

Verbs the module does not implement — `fill`, `setProperty`, `revision`,
`copyStyles`, `clearStyles` — are reported per-operation as
`unsupported-operation`, and nothing in the plan is applied. The four slide verbs
run the other way: a Word document reports *them* as unsupported.

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
`slide256/shape2/p0`, over the slide master, theme, and the five layouts
[Generating a deck](#generating-a-deck) names. That anchor is the PowerPoint
counterpart of a blank Word document's `auto-0000`.

Fill it with `changeText` and an empty `expect`:

```jsonc
[ { "op": "changeText",
    "target": { "paraId": "slide256/shape2/p0", "expect": "" },
    "with": "Quarterly Review", "mode": "Direct" } ]
```

The empty `expect` is still content-verified — it asserts the paragraph *is*
blank, so if the deck drifted and something is already there the operation fails
with `expect-mismatch` rather than overwriting it. Use it to fill a placeholder
that is already there; use `insert` to add a line beside one that has text.

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
- **Shape properties and layout.** Nothing rewrites a shape's style or
  placeholder relationship. A shape's transform changes only when a `format`
  explicitly moves or resizes it — and on a placeholder that means materialising
  a transform it previously inherited, which pins it against later layout
  changes. Everything else touches the text body, the table, or the shape tree
  only.
- **Slide layouts and masters.** Untouched. New shapes inherit from the layout
  the slide already points at.

## Anchor stability

`Stabilize` is deliberately a no-op returning an empty alias map. Word assigns
`w14:paraId` there so that an operation which shifts paragraph offsets cannot
redirect a later operation's target. DrawingML has no equivalent identifier to
mint — `a:p` admits neither an id attribute nor an extension list — so the alias
map cannot be built, and `insert` genuinely does shift offsets.

The shift is therefore handled the other way round. The module implements
`IPlanValidatingModule`, and refuses any plan that both inserts a paragraph and
addresses that **same text body** at an equal or higher index:

```text
operation-conflict: 'slide257/shape3/p2' is addressed in the same plan that
inserts a paragraph at 'slide257/shape3/p1'. … Apply the insert, re-inspect,
then send the rest as a second plan.
```

Earlier paragraphs in that body, other shapes, and other slides are unaffected
and stay addressable in the same plan. Content verification would already catch
most of this — the renumbered line rarely carries the expected text — but not an
anchor with an empty `expect`, and not two lines that read alike. Splitting the
plan costs one round trip; guessing costs a wrong edit.

The slide verbs need none of that, because a paragraph id names its slide by the
durable id from `p:sldIdLst`, not by the slide's position:

- `insertSlide` and `duplicateSlide` mint an id above every existing one, so they
  cannot collide with an anchor already issued.
- `moveSlide` rewrites only the order of `p:sldIdLst`. Every id, and therefore
  every anchor, is untouched — which is the whole reason anchors key on the id.
- `removeSlide` invalidates anchors into the slide it removes, and only those.
  An operation later in the same plan that targets the removed slide fails with
  `anchor-not-found` at apply time and the whole plan is rolled back.

**If a paragraph-inserting verb is added to this module, that reasoning stops
holding** and a durable id scheme is required first — an empty alias map would
silently let one operation retarget another.

## Generating a deck

Several `insertSlide` operations in one plan author a deck end to end, so
`create_document` plus an initial plan produces a finished presentation in a
single call:

```jsonc
[ { "op": "changeText",
    "target": { "paraId": "slide256/shape2/p0", "expect": "" },
    "with": "FY27 Operating Plan", "mode": "Direct" },
  { "op": "insertSlide",
    "slide": { "layout": "titleAndContent",
               "title": "FY27 Priorities",
               "body": [ "Finish the billing migration by Q2",
                         "Rebuild the APAC pipeline" ],
               "notes": "Do not commit to a date beyond Q2." } },
  { "op": "insertSlide",
    "slide": { "layout": "sectionHeader", "title": "Financials" } } ]
```

| Layout | Placeholders |
| --- | --- |
| `title` | centred title + subtitle — the opening slide |
| `titleAndContent` | title + bulleted body. The default when `body` is supplied |
| `sectionHeader` | title + short standfirst |
| `titleOnly` | title. The default when only a `title` is supplied |
| `blank` | none. The default when neither is supplied |

`layout` is resolved against the layouts **the deck itself defines**, matched by
PresentationML layout type. A deck created here ships all five; a deck built from
a corporate template is matched against that template's layouts, so its styling
wins. A layout the deck does not define falls back to the deck's first rather
than failing — a slide in the wrong layout is recoverable, a refused edit on
someone's template is just an obstacle.

An inserted slide carries no geometry of its own: its shapes name the layout's
placeholders and leave `p:spPr` empty, so position, size, font, and bullet
styling are inherited. Restyling the layout later restyles the slide, which would
not happen had the coordinates been baked in at creation time.

**A slide added by a plan cannot be edited by a later operation in that same
plan** — its slide id does not exist until the plan is applied, and `find`
targets resolve against the pre-edit document. Set its text through
`insertSlide`'s own `title`/`body`/`notes`, or apply, re-inspect, and edit.

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

- `format` does not reach shape fills, outlines, or table styles; it covers run and paragraph formatting plus image size.
- A slide's content is set when it is inserted; there is no verb that adds a shape to an existing slide beyond a table or an image.
- Legacy `p:cm` comments are neither read nor written.
- Charts and SmartArt are not addressable.
