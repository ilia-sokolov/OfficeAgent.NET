---
name: word-document-review
description: Review and edit existing Word documents (.docx) through OfficeAgent with tracked changes, comments, revision decisions, table updates, footnotes, or page-layout changes. Use when a user asks to revise, redline, mark up, or resolve review state in an existing Word file. Do not use for read-only extraction, summarization, or creating a new document.
---

# Reviewing an existing Word document

OfficeAgent edits the real `.docx` package. A clause you replace keeps its numbering,
its comments, and its style; a redline you write is one Word shows in the review pane and
a human can accept or reject. That is the whole point of using it instead of extracting
text, rewriting it, and handing back a new file: **the document that comes out is the one
that went in, minus the edit you made.**

Working that way means never guessing where something is. Every edit names a location the
server gave you and the text expected there, so an edit that would land in the wrong place
fails instead.

## First: how is the document reaching the server?

Look at what `list_connections` returns, or at the tools you have. This decides everything
that follows, and getting it wrong is the most common way a session falls apart.

| What you see | What it means | Use |
| --- | --- | --- |
| A connection with `provider: "filesystem"` or `"sharepoint"` | The document is in storage the host controls | Address it by `(connectionId, documentId)`. Edits save in place. |
| A connection with `provider: "session"` | The server holds documents in memory for this session | Same `(connectionId, documentId)` loop. `import_document_content` to put a document in, `export_document_content` once at the end. |
| Only `*_content` tools, no connections | Nothing is stored anywhere | `create_document_content` / `edit_document_content`, which carry the document as base64 both ways. |

### Prefer a handle to the bytes

If both a session connection and the `*_content` tools are available, use the session
connection for anything beyond a single edit.

Chaining `edit_document_content` calls means passing the whole document back as base64 to
make the second edit — which requires reproducing several thousand characters exactly. That
does not survive contact with reality: a single wrong character makes the content stop
being a readable package, and every call after it fails on a copy you cannot repair,
because your copy *is* the damaged one. Retrying sends the same damaged string again.

Measured on the same three-step edit: by handle, 6 calls and 40 seconds, none failed. Inline,
722 seconds and failure at step 2.

So: `import_document_content` once, edit by id as many times as you need,
`export_document_content` once at the end. Reach for `edit_document_content` only when the
whole job is a single self-contained call.

If you find yourself about to paste a long base64 string into a tool call, stop — that is
the moment to import the document and switch to ids.

## The loop

**inspect → find → preview → apply → export**

### 1. Inspect

`inspect_document` returns the outline, the paragraphs with their ids and `location`
(`body`, `header`, `footer`, `footnote`, `endnote`), the content controls, the styles, a
`snapshot` etag, and `nodes` — tables, images, revisions, comments, notes, document
properties.

Read the nodes before proposing anything. On a document under review they carry the state
that decides what the user actually wants:

- `kind: "revision"` — pending tracked changes, each with author and date, under paths like
  `ins#7`, `del#7`, `markIns#7` (a paragraph split), `rowIns#7`, `runFormat#7`.
- `kind: "comment"` — each comment with its author, its text, whether it is resolved, and
  which comment it replies to.
- `kind: "table"` / `"note"` / `"image"` — addressable structures.

On a large document use `fidelity: "outline"` or `"structure"` first and only pull
`"content"` for the part you need. Page with `paragraphOffset` / `paragraphLimit`.

**Read the open comments before editing a document under review.** They usually say what
the edit should be, and answering one is often more useful than guessing at the text.

### 2. Find

`find_in_document` returns content-verified anchors — `paraId`, `expect`, `occurrence`, and
`location`. Use it rather than composing an anchor yourself. `location` matters: identical
text in a header and in the body are different places, and only the anchor tells them apart.

Alternatively, in `edit_document` and `edit_document_content`, a target may name text
directly as `{ "find": "Acme Corp" }` and the server resolves it. That saves a round trip.
If the text matches more than once the call fails and lists each candidate — re-issue with
`{ "find": "Acme Corp", "match": 2 }` (zero-based), or use more surrounding text. Never
guess a match index; use one the error listed.

### 3. Preview

`preview_plan` runs the whole plan and reports what would change, writing nothing. Use it
when the plan is speculative, touches many places, or was built from text you inferred
rather than looked up. Skip it for a single edit you are confident of — it costs a round
trip.

### 4. Apply

`apply_plan` commits and saves. **Applying is all-or-nothing**: if any operation fails,
nothing is written. That is what makes a multi-operation plan safe — you never end up with
half a change.

Put related edits in one plan rather than one call per edit. It is faster, and it means the
document is never in a half-edited state.

Copy the `snapshot` from inspect into the plan as
`"snapshot": { "eTag": "<snapshot>" }`. It detects the document changing under you between
inspect and apply.

### 5. Export

Only for a session connection: `export_document_content` once, at the end, so the host can
save or deliver the result. Do not export between edits — each export spends the whole
document in context for nothing. Documents in a session connection are gone when the server
stops, so say so if the user might expect the file to be saved somewhere.

For a filesystem or SharePoint connection there is nothing to export: `apply_plan` already
wrote the file where it lives.

## Make review edits explicitly tracked

An omitted `mode` inherits the MCP connection's `DefaultChangeMode`. That connection
default may be `Direct`, especially when one connection handles both Word documents and
PowerPoint decks. For a reviewable Word edit, pass `"mode": "Tracked"` on every
operation that supports `mode`; do not rely on the connection default. This applies to
text changes, insertions, deletions, formatting, table and image edits, breaks, and notes.
Pass `"mode": "Direct"` only when the user has asked to change the document outright.

A deck (`.pptx`) has no revision vocabulary at all and refuses `"mode": "Tracked"`. Use
`Direct` there and tell the user that deck edits cannot be redlined; add a comment if the
change needs flagging.

Reviewing an existing redline:

```json
{ "op": "revision", "target": { "kind": "revision", "path": "all" }, "action": "Accept" }
{ "op": "revision", "target": { "kind": "revision", "path": "author:Jane Doe" }, "action": "Reject" }
```

`all` takes every revision; `author:<name>` takes one person's, which is how a document gets
cleared one reviewer at a time. A named path like `ins#7` addresses exactly one.

Comments are a conversation, not a write-only log:

```json
{ "op": "comment", "target": { "kind": "comment", "path": "comment#1" }, "action": "Reply", "text": "Confirmed with legal." }
{ "op": "comment", "target": { "kind": "comment", "path": "comment#1" }, "action": "Resolve" }
```

## When something fails

The server answers with a stable code. Each one has a specific recovery — retrying the same
call is almost never it.

| Code | What happened | Do this |
| --- | --- | --- |
| `expect-mismatch` | The text at that anchor is not what the plan expected — the document moved on | Re-inspect or re-find that operation's target, then re-issue. Do not force it. |
| `stale-snapshot` | The document changed since you inspected it | Re-inspect and rebuild the plan against the new snapshot. |
| `ambiguous-anchor` | The text matches several places | Re-issue with the `match` index the error listed, or more surrounding text. |
| `anchor-not-found` | Nothing matches | Check the wording against `inspect_document`. Re-sending the same text will fail identically. |
| `invalid-operation` | The operation is structurally wrong — an empty `expect` against a paragraph with text, a `format` with no properties, a deck asked for `Tracked` | Read the message; it names the specific problem. |
| `unsupported-operation` | The verb does not exist for this format — `revision`, `setProperty`, `pageSetup`, `insertBreak`, `note` on a deck | Do not translate it into something else silently. Tell the user this format cannot do it. |
| `requires-renderer` | Needs pagination or field recalculation | Out of scope. Explain rather than approximating. |
| `operation-conflict` | Two operations target one location in a plan | Merge them, or apply in separate plans. |
| `version-conflict` | The stored document changed since you opened it | Re-open and rebuild. Someone else may be editing it. |
| `invalid-argument` on `contentBase64` | The content did not arrive intact | Your copy is damaged; re-sending it fails the same way. Import the document by handle instead. |
| `io-error` | The file could not be read or written | Usually open in Word. Tell the user to close it. |
| `not-found` / `access-denied` | The id does not resolve, or is outside the connection | Do not retry with a path — ids are provider-assigned. |

A plan that fails wrote nothing. Say what failed and why rather than reporting partial
success.

## What this cannot do

Be straight with the user about these rather than approximating:

- **No rendering.** Page counts, "does this fit on one page", table-of-contents *text*, and
  field values are outside the engine. It edits the document; Word lays it out.
- **No Excel.** Word and PowerPoint only.
- **Decks cannot be redlined.** PresentationML has no revision model.
- **A note's number is Word's.** Never type a superscript number into the text — add a real
  footnote and Word renumbers the rest.
- **List numbers are Word's too.** Use `format` with `listStyle`; typing "1." into the text
  produces "1. 1." once Word draws its own.

## A worked example

The user says: *"Open contract.docx, change the payment terms from 30 to 45 days, and flag
the liability cap for legal."*

1. `list_connections` → `documents` (filesystem).
2. `open_document("documents", "contract.docx")` — registers and inspects in one call.
3. Read the paragraphs and the `nodes`. Note any existing comments and revisions.
4. `find_in_document("thirty days")` → one hit, `location: "body"`.
5. One plan, applied once:

```json
{ "snapshot": { "eTag": "<from inspect>" },
  "operations": [
    { "op": "changeText",
      "target": { "paraId": "w14:...", "expect": "thirty days", "occurrence": 0 },
      "with": "forty-five days",
      "mode": "Tracked" },
    { "op": "comment",
      "target": { "paraId": "w14:...", "expect": "liability cap" },
      "text": "Legal to confirm the cap before signature.",
      "author": "Reviewer", "initials": "RV" }
  ] }
```

6. Report: the term is now a tracked change awaiting review, and a comment is on the
   liability cap. The file was saved in place.

The edit uses `mode: "Tracked"` explicitly, so it remains reviewable even when the
connection default is `Direct`. Both changes are in one plan, so the document is never
left half-changed.
