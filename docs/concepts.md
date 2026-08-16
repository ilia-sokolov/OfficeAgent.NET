# Concepts

OfficeAgent.NET is a translation layer between AI agents and OOXML. An agent
expresses intent as a typed document plan; a format module validates that plan
and performs the Open XML changes. The agent works with engine-issued addresses
instead of editing package XML.

## Document providers

Documents live behind `IDocumentProvider` connections. Filesystem and SharePoint
providers ship with the project; another store can implement the same
`RegisterAsync`, `OpenReadAsync`, `SaveAsync`, and `RemoveAsync` contract.

A provider stores references, not document bytes. Registering an existing
document mints an opaque `documentId`; later calls address it by
`(connectionId, documentId)`. The default in-process tools expose only those ids
and never credentials. Opt-in registration and composite tools also accept a
path, SharePoint URL, or `driveId/itemId` from the model. The MCP server enables
registration by default because an MCP client has no other staging channel.
Those source values can enter model context, but the provider still enforces its
configured root or identity boundary.

## Inspect

`InspectAsync` returns a format-specific structured model. Common fields include
the format, snapshot etag, paragraph text and ids, style catalog, and
addressable nodes.

- Word inspection includes outline headings, content controls, bookmarks,
  tables, images, document properties, and revisions.
- PowerPoint inspection uses slide/shape-scoped paragraph ids and nodes for
  slides, shapes, tables, images, comments, media, and sections.

See [Document plans](document-plans.md) and
[PowerPoint support](powerpoint.md) for exact target paths.

## Anchors

An anchor is an engine-issued address. Callers reuse anchors from
`InspectAsync` or `FindAsync`; they do not invent them.

| Anchor | Targets | Example |
| --- | --- | --- |
| `TextSpanAnchor` | Expected text in a paragraph | `{ paraId, expect, occurrence }` |
| `StructuralAnchor` | A named slot, such as a content control, bookmark, or deck shape name | `{ tag, kind: "contentControl" }` |
| `NodeAnchor` | A format node such as a table, image, slide, shape, property, or revision | `{ kind: "table", path: "table#0" }` |
| `StyleAnchor` | A named style | `{ styleId: "Heading1" }` |

Text and node anchors carry expected content. At apply time the engine checks
the live document again. Drift produces `expect-mismatch` instead of an edit at
the wrong location.

## Snapshots and drift detection

`Inspect` returns a `SnapshotToken` etag over the format's text hosts. For Word
it hashes body, header, footer, footnote, and endnote XML. For PowerPoint it
hashes slide and speaker-notes XML. It does not cover properties, comments,
sections, media/image bytes, masters, or layouts. Put the token in
`DocumentPlan.Snapshot` to detect drift in that covered content; a mismatch
produces `stale-snapshot`. Omitting it relies only on per-anchor checks.

## Document plans

A `DocumentPlan` is a JSON-serializable list of typed operations:

```jsonc
{
  "operations": [
    {
      "op": "changeText",
      "target": {
        "paraId": "w14:00000002",
        "expect": "Acme Corp",
        "occurrence": 0
      },
      "with": "Globex Inc.",
      "mode": "Tracked"
    }
  ]
}
```

The vocabulary is shared, but each format implements only operations it can
express. An unsupported operation produces `unsupported-operation`, and no part
of the plan is written. Some verbs, including `headerFooter`, have
format-specific semantics; slide lifecycle, shape, section, media, transition,
and animation operations are deck-only. Use the maintained
[operation reference](document-plans.md#supported-verbs) instead of relying
on verb counts.

`DocumentPlan.Format` defaults to `Unspecified`. Set it only to assert a format;
a mismatch produces `contract-mismatch`. An operation without `mode` uses the
connection's default change mode—`Tracked` unless configured otherwise. A deck
rejects tracked changes because PresentationML has no redline model.

## Preview and commit

- `PreviewAsync` validates the complete plan and reports proposed changes and
  errors without writing.
- `CommitAsync` revalidates against live state, applies the complete plan in
  memory, and saves through the provider. A plan is never partially committed.

`Replace` is the default save mode and uses optimistic concurrency. `NewVersion`
preserves the source, writes a versioned sibling, and returns a new id.
`NewDocument` also returns a new id. `NewName` is optional; when omitted, the
provider derives a versioned sibling name just as it does for `NewVersion`.

## Capabilities

Every proposed change declares the engine capability it needs:

- `Deterministic`: pure Open XML completes the edit.
- `DeferredToWordOnOpen`: Word refreshes a displayed value when the file opens.
- `NeedsRenderer`: layout or calculation is required. The engine refuses these
  changes rather than guessing.

## Transactions and anchor stabilization

Before a Word commit, the module gives paragraphs without a durable
`w14:paraId` one and maps positional `auto-NNNN` ids to the new ids. Re-inspect
after a commit: saved bytes and ids may have changed, so a previous snapshot or
`auto-NNNN` address is not a safe follow-up target.

PresentationML has no equivalent paragraph id. The PowerPoint module instead
refuses a plan that inserts a paragraph and then positionally addresses the same
text body at or after that insertion. Apply the insert, re-inspect, and send a
second plan. Slide and shape node paths use durable OOXML ids and survive slide
reordering. See [PowerPoint anchor stability](powerpoint.md#anchor-stability).
