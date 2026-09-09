# Sample document

`services-agreement.docx` is a fictional contract for trying OfficeAgent against something
that behaves like a real document. Northwind Traders and Contoso Consulting do not exist.

This sample exercises the Word review subset of OfficeAgent. It is one focused example,
not a complete capability demo; the project also creates and edits Word documents and
PowerPoint decks through MCP, Agent Framework tools, and the .NET API.

Copy it into the directory your connection is rooted at, then follow
[the Word edit](../../README.md#try-a-word-edit).

## What is in it

It is deliberately not a blank page — it is a document already part-way through a review,
which is the state most documents are in when an agent meets them:

| | |
| --- | --- |
| Structure | A title, six numbered clauses under `Heading2`, and a three-row milestones table |
| Text to edit | Clause 3 says payment is due within **thirty days** |
| A pending redline | Clause 3's interest rate is a tracked change from `2%` to `3%`, not yet accepted |
| An open comment | Priya Raman has asked on clause 4 whether the twelve-month liability cap is acceptable |

That mix is the point. An agent that only replaces text will miss the redline and the
comment; the tools that surface them are `inspect_document`'s `nodes`, and handling them
well is what the
[word-document-review skill](../../skills/word-document-review/SKILL.md) teaches.

## Expected result of the first edit

Asking *"In services-agreement.docx, change the payment terms from thirty days to
forty-five days"* should produce:

- **Clause 3 reads "within forty-five days of receipt"**, as a tracked change — `thirty days`
  struck through, `forty-five days` underlined beside it. Open the file in Word and both
  appear in the review pane, attributable and reversible.
- **Nothing else moved.** The heading styles, the table, the existing 2%→3% redline and
  Priya's comment are all still there, untouched. This is the part worth checking: a
  rewrite-and-replace approach loses the comment and flattens the pending revision.
- **The file was saved in place**, guarded by a version check. There is no second file
  unless you asked for `saveMode: "NewVersion"`.

If the agent reports `expect-mismatch`, the document already says forty-five days — it has
been edited before. Re-inspect and it will tell you what the clause currently reads.

## Reproducible review workflows

Start each workflow from a fresh copy of `services-agreement.docx`. This keeps the expected
state identical and makes failures comparable. Install the
[Word review skill](../../docs/skill-installation.md), connect the containing directory as
`documents`, and record the client, model, OfficeAgent version, and skill revision.

### 1. Change a clause as a tracked revision

Prompt:

> In `services-agreement.docx`, change the payment terms from thirty days to forty-five
> days. Preserve all existing comments and revisions and leave the change for review.

Expected result:

- Clause 3 contains an insertion for `forty-five days` and a deletion for `thirty days`.
- Priya's open comment and the existing 2%→3% revision remain unchanged.
- The milestones table still has its original rows and formatting.

### 2. Reply to and resolve the existing comment

Prompt:

> In `services-agreement.docx`, find Priya Raman's open liability-cap comment. Reply that
> the twelve-month cap is approved, then resolve the thread. Do not alter document text or
> accept any tracked changes.

Expected result:

- Priya's comment remains in the document, has one reply, and is marked resolved.
- Clause text is unchanged and the existing 2%→3% revision is still pending.
- No new paragraph, table row, or text revision appears.

### 3. Add a review milestone

Prompt:

> In `services-agreement.docx`, add a tracked row to the milestones table for a September
> review, due 2026-09-15, fee 4,000. Preserve its existing style, comments, and revisions.

Expected result:

- The milestones table has one additional row with the requested three values.
- The row is reviewable as an insertion rather than silently written directly.
- Priya's comment, the payment clause, and the existing interest-rate revision are unchanged.

Open every result in Word and inspect the Review pane. An agent report alone is not proof
that the package contains the expected revisions and comments.

## How this file was made

It was authored by OfficeAgent itself, through a plan of `changeText`, `insert`,
`insertTable`, `comment` and tracked `changeText` operations — the same verbs documented in
[Document plans](../../docs/document-plans.md). Nothing about it needed Word.
