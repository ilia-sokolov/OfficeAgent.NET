# Changelog

Notable changes per release. The body of each version section is also the text used for the
corresponding GitHub release.

## 1.0.0 — Unreleased

### Breaking changes and migration

- **Every tool response uses camelCase property names.** Before 1.0, casing depended on how a
  payload was built: `describe_capabilities` was PascalCase, and `inspect_document` mixed both.
  Requests are unaffected, since they have always been read case-insensitively. Migration: read
  camelCase, or match names case-insensitively. A host that deserialises responses into the
  library's own types should use `JsonSerializerDefaults.Web`. See
  [upgrading to 1.0](docs/upgrading-to-1.0.md).

- **Engine seams are marked `[Experimental("OFFICEAGENT001")]` and are outside the 1.x
  promise.** Twenty-five public types that exist to build format modules and operation
  handlers, including `IFormatModule`, `IOperationHandler` and `IDocumentService`, now produce
  error `OFFICEAGENT001` where your code names them. Hosting the engine never needs them.
  Migration: if you build a custom module, suppress `OFFICEAGENT001` and expect these types to
  change in minor releases. See [compatibility](docs/compatibility.md#engine-extensibility).

- **`SpreadsheetPartUtility` is internal.** It was public only so the Excel and PowerPoint
  assemblies could share it, it was hidden from IntelliSense, and its documentation called it
  infrastructure. Migration: none supported. It was never part of the documented surface.

### Added

- A [compatibility policy](docs/compatibility.md) covering versioning, unknown fields, nulls,
  defaults, deprecation and what is excluded from the 1.x promise.
- Public catalogues for every stable code: `ToolErrorCodes`, `TemplateDiagnosticCodes`,
  `ComparisonDiagnosticCodes`, `AssemblyDiagnosticCodes` and `RenderFailureCodes`, beside
  `ValidationErrorCodes`. Values are unchanged from 0.9.
- A [wire contract](docs/wire-contract.md) baseline, gated in CI beside the C# API reference.
  It records plan operations, anchors, the input schemas of all 22 tools, configuration keys, the
  default of every member a caller constructs, and responses in three layers: schemas derived
  from the 7 typed results, every anchor summary and receipt state, and 69 named response states
  including failures, cancellation, uncertain writes and partial batches. The gate also checks
  that the netstandard2.0 build exposes exactly the net8.0 surface.
- A tested [sandboxed renderer reference](deploy/renderer/README.md): a pinned LibreOffice and
  Poppler image and a host-side `IDocumentRenderer` that runs each render in a fresh container
  with no network, a read-only root, no capabilities, and memory, process and scratch limits.
  Each boundary is proved by an automated test in a dedicated CI job.
- `ToolErrorCodes.Cancelled`. The `cancelled` code every tool already emitted was the only one
  missing from the catalogues.

### Fixed

- Documents OfficeAgent creates can be merged. The blank Word document's heading styles carry
  an outline level, as do styles written by `defineStyle` and `format`, and merge refused
  `w:outlineLvl` as unsupported content, so two documents `create_document` had just produced
  failed `preview_document_merge` with `unsupported-merge-content`.
- The ingestion-limits guide no longer shows an internal constructor that no host could call.
- Removing an unknown id from a memory or session connection fails with `not-found`, as every
  other provider already did, instead of reporting `removed: true`.
- The renderer refuses a file that is not the macro-free OOXML package its extension declares,
  with `renderer-failed`, before LibreOffice starts. LibreOffice chooses an import filter by
  content, so random bytes named `.docx` were rendered as a text document, and an OpenDocument
  file under a `.docx` name reached LibreOffice's ODF import.
- `describe_capabilities` reports `renderingAvailable: true` when an `IDocumentRenderer` is
  registered. It was always `false`, contradicting its documentation.

## 0.9.0 — 2026-09-20

### Breaking changes and migration

- **Edit plans are now checked against a wire contract.** A plan carrying an unknown
  `contractVersion` is rejected with the stable code `contract-mismatch` before any provider
  is touched. Omitting the field keeps the legacy `0.2` behaviour, and the only value accepted
  explicitly is `"0.2"`. Plans with unknown properties, unknown operations or unknown enum
  values are rejected with `invalid-json` rather than being silently ignored. A v0.8 plan that
  omitted the field, or set it to `0.2`, still applies with identical change semantics.
  Migration: send no `contractVersion`, or exactly `"0.2"`, and stop sending fields the schema
  does not define. See [document plans](docs/document-plans.md).

- **Host resource ceilings are now enforced on every ingested package.** Seven limits bound
  compressed size, expanded size, entry count, XML characters, nesting depth, expansion ratio
  and part count. A package over a limit is refused with `input-too-large`, and a damaged or
  hostile package with `malformed-package`, both before any content is parsed. These are host
  ceilings: a request may narrow them and can never raise them. Documents that worked in v0.8
  are unaffected unless they exceed a ceiling, in which case they are now refused rather than
  processed. Migration: if you handle very large documents, set the limits you need on the
  host rather than relying on the defaults. See [ingestion limits](docs/ingestion-limits.md).

- A Word edit that would overlap an existing pending revision is now refused with
  `revision-overlap` instead of producing a redline that could not be rejected back to the
  original. This fixes silent content corruption; an edit that v0.8 accepted may now be
  refused, which is the point.

### Added

- Template preflight: `discover_template` reports a template's scalar slots, image and native
  chart slots and repeating rows; `preview_template_batch` validates a whole batch without
  writing and returns a token binding that review to the exact template bytes and batch; and
  `populate_template_batch` accepts that token so a commit whose inputs changed after review is
  refused with `stale-batch-preview`. Available on the direct API, the Agent Framework tools
  and the MCP server.
- Typed template values: a binding can place an image or set the data of a native chart the
  template already contains, bounded by per-batch media limits.
- Capability discovery: `describe_capabilities` reports the accepted plan contract, every
  registered format with its verbs and change modes, the host's ceilings, and the connections
  available with the capabilities held on each.
- Word comparison now covers table cell text where the table geometry is unchanged, and no
  longer reports a difference when the same words are merely split differently across equally
  formatted runs. Findings in covered areas survive a refusal elsewhere, so a blocked
  comparison still reports what it saw.
- Reproducible benchmarks over the tracked acceptance corpus, for both the direct API and an
  installed MCP server, with published raw runs and a documented threshold policy. See
  [benchmarks](docs/benchmarks.md).
- Open XML SDK interoperability guidance for applications that use both this library and the
  SDK directly. See [SDK interoperability](docs/sdk-interoperability.md).
- A generated [C# API reference](docs/csharp-api.md) covering the public surface of every
  library package, plus a package-backed `quickedit-sample.zip` release asset that runs outside
  the repository.

### Fixed

- Concurrent replacement saves through the in-memory session provider now perform the version
  check and mutation atomically. One of two writers using the same version succeeds and the other
  receives a version conflict instead of both being told they committed.
- Session import now validates the package against every configured ingestion ceiling before
  retaining it, and inline base64 admission uses the host's configured compressed-size ceiling.
- Word comparison now attributes table-cell changes to the table coverage area and does not report
  an unchanged table as exceeding the difference limit when body changes exactly fill that limit.
- Concurrent saves through the filesystem provider no longer lose an update. The version check
  and the publish are now one step per document, so two writers cannot both be told they
  succeeded while one write is discarded.
- A package that declares a main part but omits it is refused as `malformed-package` at the
  boundary instead of escaping later as an unhandled exception.
- Comparison findings are no longer discarded when an unsupported area blocks the plan.
- Release SBOM validation now reconciles declared first-party package dependencies and rejects
  a missing component or root dependency edge. Package, sample, and container publication now
  share the validated release workflow, and every third-party workflow action is pinned to an
  immutable commit.
- The QuickEdit sample and primary tracked-edit examples now set an explicit displayed revision
  author. Quickstart project creation works with current SDKs without requiring an installed
  .NET 8 template pack.

- Made a failed render fail closed. `RenderResult.Pages` and the new `RenderResult.PageCount`
  now throw `RenderFailedException`, carrying the stable failure code, when rendering did not
  succeed. Previously a caller that did not check `Succeeded` read the empty page list of a
  failed render as a zero-page document, so a page-count gate admitted every document on a host
  where LibreOffice was absent or a limit was hit. Use `Succeeded`, `EnsureSucceeded()`, or
  `TryGetPages` to branch on the failure without an exception. `RenderResult` is now built
  through the validated `RenderResult.Success` and `RenderResult.Failure` factories instead of
  object initializers, so a result cannot hold a success state and a failure code at once, and
  the failure contract is documented in [Visual rendering](docs/rendering.md#reading-the-result).

## 0.8.0 — 2026-09-12

- Added ordered Word document assembly with source formatting compatibility checks, native
  part/identifier import, isolated section headers, hash-bound preview/commit, multi-source
  receipts, host resource limits, MCP/Agent Framework tools, and a proposal-packet sample.

- Aligned the public MCP Registry entry with the current release and added a scheduled
  distribution check across GitHub, NuGet, and the registry.
- Made `QuickEdit` runnable against the bundled fictional contract, with a non-zero failure
  when its exact source text is absent.
- Added security, support, compatibility, and release-assurance policies plus a fail-closed
  NuGet vulnerability gate with expiring documented exceptions.
- Added plan-level Word revision author/timestamp metadata and structured apply receipts
  containing plan/input/output SHA-256 hashes, outcomes, provider output references, and a
  host-supplied authenticated actor kept separate from the displayed revision author.
- Added native editable PowerPoint charts with embedded workbooks for clustered-column, bar,
  line, and pie charts, including OfficeAgent-owned chart updates.
- Added the `OfficeAgent.Excel` module for bounded workbook inspection, raw or displayed-value
  search, typed cell and formula writes, table row appends, and cell-note management. Formula
  and value writes request recalculation on open; OfficeAgent does not evaluate formulas.
- Added an injectable, principal-aware connection access policy across MCP registration, read,
  create, edit, and removal operations, plus an authenticated two-user hosted gateway reference
  whose connection discovery and document access are isolated per caller.
- Added an optional renderer boundary and out-of-process LibreOffice/Poppler implementation with
  input, time, working-set, page-count, and output-size limits. Structural preview remains usable
  without a renderer.
- Added bounded template-batch population for tagged Word and PowerPoint values, including
  repeating Word table rows, independent outputs, per-item diagnostics, and audit receipts.
- Added read-only Word body-paragraph comparison with exact input hashes, explicit coverage
  diagnostics, and a snapshot-bound plan that produces native tracked insertions and deletions.
- Exposed both workflows through the direct .NET API, Agent Framework, and MCP, with runnable
  quote-generation and document-comparison samples.

### Fixed before release

- Made Word comparison fail closed when image bytes, table geometry, direct paragraph formatting,
  or other unsupported package content changes alongside supported paragraph text.
- Remapped drawing, picture, bookmark, hyperlink-anchor, and content-control identifiers when
  repeating Word rows, and rejected rows containing comment or note references that cannot be
  cloned safely.
- Distinguished Excel cell targets by sheet and normalized A1 address, included shared strings in
  snapshot drift detection, rejected single-cell edits inside shared-formula groups and legacy
  array-formula ranges, kept legacy note shape IDs unique after deletion, and requested
  recalculation after constant writes.
- Bounded redirected-output draining by the renderer timeout, including when a descendant keeps an
  inherited pipe open after the tracked launcher exits.
- Propagated the hosted gateway's authenticated principal into audit receipts and aligned the
  package, tool-schema, workflow, snapshot, rendering, support, and security documentation.

## 0.7.0 — 2026-09-09

Word review became a first-class workflow, and the server no longer needs storage to run.

### Word: broad tracked-change coverage

Tracked changes previously reached only `changeText`. Every other verb wrote straight into
the document, so under the default review policy a reviewer saw the text replacement marked
up and the new clause silently present. Now `insert`, `fill`, `format`, supported table and
image operations, `insertBreak` and `note` honour `mode`, and Word tracks them when it is
omitted. Image resizing remains direct because WordprocessingML has no revision for drawing
extent changes.

The markup is the real thing rather than a decoration: an inserted paragraph marks its
paragraph mark as well as its runs, so rejecting removes the paragraph instead of emptying
it; a removed row stays in place struck through until someone accepts it; `format` records
the properties it replaced, so rejecting restores exactly what was there.

### Word: comment threads, notes, page geometry

- **Comments** are readable and answerable, not write-only. `inspect_document` lists each
  one with its author, text, resolved state and reply parentage; `comment` gained `Reply`,
  `Resolve` and `Remove`.
- **Footnotes and endnotes** through a new `note` verb — add, update, remove — with Word
  owning the numbering.
- **Page setup and breaks** through `pageSetup` and `insertBreak`: paper size, orientation,
  margins, and page, column or section breaks.
- **Revision review** covers every marker WordprocessingML defines, not only run
  wrappers — paragraph marks, rows, cells, moves and the formatting revisions — each with
  its author and date, addressable individually, by `author:<name>`, or as `all`.

### Running without storage

Two ways, and they are not interchangeable:

- **A session connection** (`EphemeralConnectionId`) keeps documents in the server process.
  The agent addresses them by opaque id exactly as it would documents on disk;
  `import_document_content` and `export_document_content` move bytes across the boundary.
- **Inline content** (`AllowInlineContent`) adds `create_document_content`,
  `inspect_document_content` and `edit_document_content`, which carry the document as base64
  in both directions and hold nothing.

Prefer the session connection for anything beyond a single call. Chaining inline edits
requires the model to reproduce the document exactly to make the second call. In an
exploratory run, the session route completed and the inline route failed after the model
altered the base64; the run did not retain enough metadata to serve as a benchmark.

### Fixed

- Saving intermittently failed on Windows with *"Unable to remove the file to be replaced"*
  when a scanner or indexer briefly held the destination. The atomic write now retries, and
  a failure that outlives the retry is reported as a provider IO error naming the file
  rather than as an unexplained internal error. The document read and the version read had
  the same gap.
- An `insertImage` operation carrying an `imageDocumentId` lost its change mode when the
  bytes were resolved, so it could arrive as a redline the caller had not asked for.
- Content that is not a readable package is now reported as such, with the likely cause,
  instead of an unexplained internal error.
- Regular-expression searches now have a two-second match bound. A pattern that exceeds it
  returns the stable `regex-timeout` tool error instead of occupying the server indefinitely.

### Added for adoption

- A [`word-document-review` Agent Skill](skills/word-document-review/SKILL.md) teaching the
  review loop, storage choice and error recovery, packaged as a GitHub release asset with
  complete Claude Code and Codex installation instructions.
- A [fictional sample contract](samples/documents/services-agreement.docx) with a clause to
  edit, a table, an open comment and a pending redline, plus three reproducible review
  workflows covered by engine tests.
- A complete [ContractReview agent](samples/ContractReview/) that screens deterministically,
  delegates judgement to Microsoft Agent Framework, and applies one snapshot-bound plan.
- A `--config` / `OFFICEAGENT_CONFIG` JSON configuration file, so connections need not be
  written as indexed environment variables.
- A `defineStyle` verb.
- Documentation CI for internal links, release-version alignment, skill metadata, the MCP
  Registry schema, and a published configuration example loaded by the production binder.

### Behaviour changes to be aware of

- Under the default `Tracked` mode, `removeTableRows`, `removeTable`, `removeTableColumns`
  and `removeImage` now **mark** rather than delete. Pass `"mode": "Direct"` for the previous
  behaviour. Callers that assumed removal happened immediately should check.
- With no connection configured the server previously refused to start. It now starts on an
  in-memory session connection with creation enabled - equivalent to `EphemeralConnectionId=session`
  plus `AllowCreation=true` - and says so on stderr. Anything the host configured is left
  alone; the fallback applies only when nothing at all was set. The MCP Registry entry marks
  the filesystem variables optional to match.
- The registration tools (`register_document`, `open_document`, `edit_document`) are offered
  only when a filesystem or SharePoint connection exists. They take a path or URL, which a
  session connection does not have, so on a session-only server they could only fail.

## 0.6.0

Backgrounds, Word headers and footers, list numbering. See the
[release history](https://github.com/ilia-sokolov/OfficeAgent.NET/releases).
