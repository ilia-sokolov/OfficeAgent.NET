## 0.8.0 — 2026-09-10

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
