# Editing a document that already has tracked changes

A document arriving from review usually carries pending revisions from someone else.
This page defines what OfficeAgent shows you of such a document, which further edits
it will accept, and which it refuses before touching anything.

The short rule: OfficeAgent never resolves somebody else's revision by implication. An
edit either lands cleanly beside the existing redline, or it is refused and you decide
what to do about the earlier changes.

## What inspect and find show

Inspection and search present the document as it would read **if every pending change
were accepted**.

| Markup | In the text view | Why |
| --- | --- | --- |
| `w:ins` inserted run | Shown | The text a reader sees with changes displayed |
| `w:moveTo` moved-in run | Shown | Same content, recorded as the destination of a move |
| `w:del` deleted run | Hidden | Struck-through text is not part of the document's text |
| `w:moveFrom` moved-out run | Hidden | The source half of a move, struck through like a deletion |
| `w:rPrChange`, `w:pPrChange` and the other `*PrChange` markers | Text unaffected | Formatting history, not content |
| `w:ins`/`w:del` on a paragraph mark | Text unaffected | Records a split or merge, not words |

So `find_in_document` cannot match text that is pending deletion, and a paragraph's
`text` in `inspect_document` already contains other authors' pending insertions. Anchors
address that same view, which is what makes the overlap rules below necessary.

Every pending revision is separately discoverable. `inspect_document` returns a
`revision` node per marker, each with a path such as `ins#7`, `del#3`, `markDel#4`,
`rowIns#2` or `runFormat#1`, and a summary naming the author and date. Paragraph-mark
revisions (`markIns`, `markDel`) are distinct from run revisions (`ins`, `del`): the
first records a paragraph boundary appearing or disappearing, the second records words.

## Which successive edits are allowed

| Case | Result |
| --- | --- |
| Edit clear of every pending revision | Allowed. Earlier authors' revision identity, author and date are untouched |
| Edit wholly inside one pending insertion | Allowed. Produces the nested `w:ins`/`w:del` Word itself writes when someone edits text a pending insertion added |
| Edit wholly inside one pending deletion | Not reachable: deleted text is not in the text view, so no anchor addresses it |
| Edit spanning a pending deletion or move-from | **Refused** with `revision-overlap` |
| Edit reaching out of a pending insertion into unmarked text | **Refused** with `revision-overlap` |
| Edit spanning two different revision containers | **Refused** with `revision-overlap` |
| Formatting change over text with a pending `w:rPrChange` | Allowed, but see the formatting-history limit below |
| Structural verbs (rows, columns, tables) over pending row revisions | Allowed. Marked with their own `rowIns`/`rowDel`/`cellIns`/`cellDel` markers |
| Accept or reject addressed by id, `author:<name>`, or `all` | Allowed. Resolves exactly what is addressed and nothing else |

The refusal is mode-independent. A `direct` edit across a pending deletion is refused
for the same reason a `tracked` one is: it would leave the earlier author's deletion
stranded in a paragraph that no longer surrounds it, which resolves their revision by
implication.

### Why spanning is refused

The text view hides deleted text, so two runs that read as adjacent can have another
author's deleted words physically between them:

```text
Original           Alpha beta gamma delta.
Author A deletes   Alpha [del: beta ]gamma delta.
Text view          Alpha gamma delta.
```

An edit to `Alpha gamma` looks contiguous but is not. A tracked replacement writes its
`w:del` and `w:ins` beside the first run it covers, which puts the replacement *before*
A's earlier deletion. Rejecting everything then restores A's words in the wrong place:

```text
Reject all         Alpha gammabeta  delta.     ← not the original
```

That document cannot be recovered by any sequence of accepts and rejects, so the edit is
refused before it is applied rather than after. `preview_plan` reports it and
`apply_plan` writes nothing.

### Recovering from `revision-overlap`

1. Read the `revision` nodes from `inspect_document` to see whose changes are pending.
2. Resolve them with a `revision` operation: accept or reject by id, by `author:<name>`,
   or `all`.
3. Re-inspect. Anchors and snapshots taken before the resolution are stale.
4. Reissue the edit.

Alternatively, target a narrower span that lies wholly inside or wholly outside the
pending revision.

### The limit of formatting history

WordprocessingML allows one `w:rPrChange` per run, so formatting revisions do not stack.
When a second author restyles text that already carries a pending formatting revision,
the formatting is applied but the existing revision is kept rather than replaced.

That is deliberate. The stored revision records what the run looked like before anyone
touched it, so rejecting restores the genuine original instead of an intermediate state
that never existed in the document. The cost is attribution: the later author's
formatting change is present in the document and carries no separate author or date of
its own.

Content revisions are unaffected. Each `w:ins` and `w:del` keeps its own author and date,
and a second author editing words inside a first author's pending insertion produces two
separately attributed revisions.

## What accept and reject guarantee

- Addressed revisions are resolved and no others. Resolving `author:Dana Reyes` leaves
  every other author's revisions pending, with their author and date intact.
- Order does not matter for independent revisions: accepting A then B gives the same
  document as accepting B then A.
- Rejecting restores the original content for the addressed revisions only.
- Unrelated comments survive both operations.
- Accepting a deleted footnote or endnote reference removes the orphaned note, matching
  what Word does.
- The result is a schema-valid package.

Structural validity is not the same as Word agreeing with you about layout or
tracked-change semantics. Open a representative multi-author result in the Office
application your reviewers use before adopting a workflow. See
[Word preservation evidence](word-preservation-evidence.md) for what is measured per
operation.

## What is out of scope

- No three-way reconciliation between two independently edited copies.
- No automatic cleanup or rewriting of earlier review history.
- Document comparison still refuses a document that already carries revisions; that
  restriction is unchanged here.
- The displayed revision author is document metadata, not an authenticated actor. See
  [concepts](concepts.md) for the audit identity that is.
