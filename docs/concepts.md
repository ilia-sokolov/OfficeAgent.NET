# Concepts

OfficeAgent.NET is a translation layer between AI agents and OOXML: the agent expresses intent as a typed plan, and the library turns it into valid Open XML manipulations. It keeps document understanding separate from document mutation - an agent receives stable addresses (*anchors*) and a typed *plan* shape; it never writes Open XML.

## Document providers

A document lives behind an `IDocumentProvider`. Hosts register one or more provider *connections* - a filesystem provider ships in the box; SharePoint, Google Drive, MCP resources, or anything else implement the same `RegisterAsync` / `OpenReadAsync` / `SaveAsync` / `RemoveAsync` interface.

A provider is a **registry of references**: it persists only where each document lives (a path, URL, drive id, …), never the bytes themselves. Registering a document with a connection mints an **opaque, provider-assigned `documentId`**, and every later call addresses the document by `(connectionId, documentId)`. The provider routes reads and saves back to the referenced location. The agent cannot register documents, construct an id, escape the connection, or see a filesystem path.

## Inspect

`InspectAsync` returns a structured caller-facing model:

- format and snapshot etag;
- outline nodes derived from heading styles;
- paragraphs with stable ids, style ids, and their logical text (including which paragraphs live inside a table cell);
- structural anchors for content controls and bookmarks (by tag);
- the style catalog;
- **nodes** - non-paragraph addressable objects (tables, images, document properties, tracked revisions) with their stable path strings.

This map is what the agent reads before producing a plan.

## Anchors

An anchor is an engine-issued address. The caller (or the LLM) does not invent anchors - it reuses the ones returned by `InspectAsync` / `FindAsync`.

| Anchor | Targets | Example |
| --- | --- | --- |
| `TextSpanAnchor` | expected text in a paragraph | `{ paraId, expect, occurrence }` |
| `StructuralAnchor` | content control or bookmark | `{ tag, kind: "contentControl" }` |
| `NodeAnchor` | table, image, document property, revision | `{ kind: "table", path: "table#0" }` |
| `StyleAnchor` | a named style | `{ styleId: "Heading1" }` |

Text and node anchors carry expected content. At apply time the engine re-verifies the live document against the anchor; if the content has drifted, the operation fails safely with `expect-mismatch` rather than editing the wrong run.

## Snapshots and drift detection

`Inspect` returns a `SnapshotToken` etag covering every editable host (body, headers, footers, footnotes, endnotes). Plans may carry that snapshot in `DocumentPlan.Snapshot`; if the live document has drifted since inspection, validation fails with `stale-snapshot`. Leaving the snapshot unset opts out of document-level drift and relies on per-anchor verification.

## Document plan

A `DocumentPlan` is a typed list of operations against anchors. The plan and every operation type are JSON-serialisable so an LLM can produce them.

```jsonc
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

The Word module ships 15 verbs covering text, runs, tables (rows / columns), images, styles, comments, properties, and revisions. See [document-plans.md](document-plans.md).

## Preview / commit

- `PreviewAsync` validates the whole plan against the current document and returns a *change report* - proposed before/after, declared capability per change, and any validation errors. Nothing is written.
- `CommitAsync` applies the plan atomically and saves through the provider. Every operation is re-verified against live state immediately before it runs; if any step fails, no partial result is written.

The default save mode is `NewVersion`: the source document is preserved and a fresh `documentId` is minted for the result. `NewDocument` mints a fresh id with a caller-supplied display name. `Replace` overwrites the source after an optimistic-concurrency check on the source version.

## Capabilities

Every proposed change declares what kind of engine support it needs:

- **`Deterministic`** - the engine completes the edit with pure Open XML.
- **`DeferredToWordOnOpen`** - the XML can be written, but Word refreshes the displayed value when the file is opened (e.g. `setProperty` for `updateOnOpen`).
- **`NeedsRenderer`** - a layout or calculation engine is required (pagination, field recalculation, table-of-contents rendering). The engine rejects these instead of guessing.

## Transactions and anchor stabilisation

Before commit, the format module stabilises anchor ids: Word assigns a durable `w14:paraId` to any paragraph that lacks one and maps the positional `auto-NNNN` ids returned by inspection to the stable ids. Each operation is then re-verified against live state immediately before it is applied, so an operation that inserts a paragraph cannot redirect a later operation's anchor.

Because stabilisation rewrites the saved bytes, the snapshot returned by `Inspect` will not match the *saved* document - re-inspect before authoring a follow-up plan.
