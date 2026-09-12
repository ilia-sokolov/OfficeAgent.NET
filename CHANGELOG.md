# Changelog

Notable changes per release. The body of each version section is also the text used for the
corresponding GitHub release.

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
