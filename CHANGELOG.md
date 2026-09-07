# Changelog

Notable changes per release. The body of each version section is also the text used for the
corresponding GitHub release.

## 0.7.0 — unreleased

Word review became a first-class workflow, and the server no longer needs storage to run.

### Word: every edit can be a redline

Tracked changes previously reached only `changeText`. Every other verb wrote straight into
the document, so under the default review policy a reviewer saw the text replacement marked
up and the new clause silently present. Now `insert`, `fill`, `format`, the table and image
verbs, `insertBreak` and `note` all honour `mode`, and Word tracks them when it is omitted.

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
requires the model to reproduce the document exactly to make the second call, and it does
not: measured on the same three-step edit, by handle 40 seconds with nothing failing, inline
722 seconds and failure at step 2.

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

### Added for adoption

- A [`word-document-review` Agent Skill](skills/word-document-review/SKILL.md) teaching the
  review loop, storage choice and error recovery.
- A [fictional sample contract](samples/documents/services-agreement.docx) with a clause to
  edit, a table, an open comment and a pending redline.
- A `--config` / `OFFICEAGENT_CONFIG` JSON configuration file, so connections need not be
  written as indexed environment variables.
- A `defineStyle` verb.

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
