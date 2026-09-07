# Sample document

`services-agreement.docx` is a fictional contract for trying OfficeAgent against something
that behaves like a real document. Northwind Traders and Contoso Consulting do not exist.

Copy it into the directory your connection is rooted at, then follow
[the first edit](../../README.md#your-first-edit).

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

## Things to try next

- *"What is still open in this document?"* — the agent should find Priya's comment and the
  pending interest-rate change, rather than reporting the document as clean.
- *"Reply to Priya saying the twelve-month cap is fine, and resolve it."*
- *"Accept all tracked changes."* — clears both the 2%→3% change and your own edit.
- *"Add a row to the milestones table for a September review, due 2026-09-15, fee 4,000."*

## How this file was made

It was authored by OfficeAgent itself, through a plan of `changeText`, `insert`,
`insertTable`, `comment` and tracked `changeText` operations — the same verbs documented in
[Document plans](../../docs/document-plans.md). Nothing about it needed Word.
