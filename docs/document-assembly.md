# Word document assembly

OfficeAgent 0.8 can assemble ordered, whole `.docx` documents into one new editable Word
package. It preserves supported source content and formatting, starts each imported document
in a section on the next page, and leaves the sources untouched. This workflow does not
reconcile independently edited versions, select page ranges, or normalize a corporate theme.

## Preview and commit

```csharp
var preview = await client.PreviewMergeAsync(new DocumentMergeRequest
{
    Sources = new[] { proposalReference, scopeReference, appendixReference },
    Options = new DocumentMergeOptions { Title = "Proposal packet", Author = "Northwind" }
});
if (preview.IsValid)
{
    var result = await client.CommitMergeAsync(
        preview.Plan!, "output-connection", "proposal-packet.docx");
}
```

`Sources` contains provider references in output order. A reference can omit `Provider` when
its `ConnectionId` selects a unique configured provider. `ItemId` is the opaque document ID.
Preview reads every source, builds and schema-validates a candidate in memory, and returns
source counts, decisions, diagnostics, and a serializable plan. It saves nothing.

Commit re-reads every source and checks its exact SHA-256 hash before rebuilding and validating
the output. Drift returns `stale-merge-source`, with no output. Changed options, input order,
or plan fields require another preview. The plan hash is an integrity checksum, not a signature
or authorization token. A host that requires approval must bind that approval to the plan hash
and destination separately. Every commit revalidates compatibility even for a supplied plan.

Provider commit always calls the destination's new-document creation capability. It cannot
overwrite an input or an existing output. Provider and cancellation exceptions propagate:
storage may have accepted a write before a response was lost. Follow the provider's recovery
procedure and do not automatically retry an uncertain creation.

For callers holding bytes, use `client.PreviewMerge(orderedBytes, options)` followed by
`client.CommitMerge(preview.Plan!, freshlyReadBytes)`. Successful in-memory results contain
`Content`; provider results contain a document reference and never return document bytes.
Compose with `WordModule`, or register `AddWordFormat()` with `AddOfficeAgent()`.

## Compatibility and preservation

| Content | Behavior |
| --- | --- |
| Paragraphs, headings, direct formatting | Copy XML and remap explicit and default styles; import all style definitions and dependencies |
| Tables and merged cells | Preserve table geometry, cell merges, styles, and content |
| Bullets and multilevel numbering | Import abstract definitions and list instances with unique IDs; source lists remain independent |
| Images | Preserve embedded PNG, JPEG, GIF, BMP, and TIFF parts and drawing placement; remap drawing IDs and relationships |
| Hyperlinks and bookmarks | Remap bookmark IDs/names and local targets; preserve HTTP, HTTPS, and mailto links without fetching them |
| Plain-text content controls | Remap control IDs; preserve tags unless they collide, then report deterministic replacements |
| Sections and headers/footers | Preserve supported geometry and section properties; explicitly resolve missing/inherited references within each source, starting with blank references |
| Page fields | Preserve PAGE, NUMPAGES, and SECTIONPAGES instructions without switches; Word or a renderer updates the displayed values |
| Properties | Keep the first source's package properties, with optional title/author overrides |

Formatting preservation has a strict boundary: sources must have equivalent document defaults,
themes, font tables, and layout-affecting settings. Differences fail as
`merge-incompatible-formatting`; the importer does not guess how to resolve them. View/zoom,
proofing state, revision-session IDs, and odd/even-header activation do not require equality.
If any source uses distinct even-page headers, sources without them explicitly reuse their
default header/footer on even pages. Style IDs and imported style names are namespaced.
Missing default styles are represented by explicit neutral styles to prevent formatting
from the first document leaking into later content.

The boundary after each non-final source uses `nextPage`, so the next source starts on a new
page. Internal section types are preserved. A source ending in non-paragraph content may require a minimal extra paragraph
to carry the boundary; preview reports it. Pagination can change, including total page fields.
Assembly does not promise pixel-identical rendering or calculate fields.

Inputs must use the standard Word part layout for shared parts. Schema-invalid packages,
existing revisions, comments, footnotes/endnotes, charts, SmartArt, VML/vector drawings,
macros, signatures, embedded objects, custom XML bindings, unsupported fields/markup,
linked images, and unknown package dependencies block the entire assembly. No content is
silently omitted. Diagnostics identify the source index and offending part. The initial
assessment stops at the first blocking issue; preview again after correcting it.

## Receipts, limits, and authorization

`DocumentMergeReceipt` records the ordered input references/hashes, effective plan hash,
exact output hash, timestamp, output reference, and host-authenticated actor. Output Author
is display metadata and cannot set the authenticated actor. No source text or tokens are logged
by the assembly workflow. Store receipts using the host's audit storage policy.

Set `OfficeAgentClient.MergeLimits` when constructing the client. These are host settings;
requests and plans cannot raise them. Defaults are 20 sources, 64 MiB compressed per source,
128 MiB compressed across sources, 256 MiB expanded across inputs, 4,000 package entries,
and 128 MiB compressed output. Expanded output and output entry counts are checked too.
Cancellation is checked during input reading, import, validation boundaries, and output writing.
Invalid requests and compressed-input limits raise argument/invalid-data exceptions; package
compatibility and expanded/output-limit failures return blocking workflow diagnostics.

MCP and Agent Framework expose:

- `preview_document_merge(requestJson)`: requires read access to every source before any
  provider is opened. JSON example: `{"sources":[{"connectionId":"docs","documentId":"a"},{"connectionId":"docs","documentId":"b"}],"options":{"title":"Packet"}}`.
- `merge_documents(planJson, destinationConnectionId, outputName)`: exposed only with
  `AllowCreation`; requires read access to all sources and create access to the destination.
  Pass the complete `plan` object from preview. Denials use `connection-forbidden`.

## Sample and acceptance corpus

Run the [DocumentAssembly sample](../samples/DocumentAssembly/) to generate a fictional
proposal packet and receipt. Its optional `--render` flag uses the isolated renderer to
produce PNG pages; renderers are not required for structural preview or assembly.

`DocumentMergeTests` covers style collisions, independent numbering, merged cells, source
order, portrait/landscape sections, header isolation, bookmarks, hyperlinks, tag collisions,
unsupported content, source drift, plan serialization, limits, provider creation, DI, and
tool behavior. Connection policy tests exercise denials before provider lookup. Acceptance
must distinguish package/schema integrity, preserved content, and visual layout: inspect
rendered sample pages and representative user documents before expanding compatibility claims.
