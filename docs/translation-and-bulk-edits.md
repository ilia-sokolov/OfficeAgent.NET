# Translation and bulk edits

Two jobs that come up constantly with Office files, and where an edit that loses formatting
or misses an occurrence is the usual failure:

- [Translate a document](#translate-a-document) and keep its styles, lists, tables, headers,
  and the bold or coloured words inside a sentence.
- [Edit many documents at once](#edit-many-documents-at-once): rename a company, update a
  disclaimer, or change a term across a folder, with one result and receipt per file.

Both work through the MCP server and `OfficeAgentTools`. Translation also works from C#
through `ChangeTextOp.WithMarked`.

## Translate a document

Translate paragraph by paragraph. Each paragraph keeps its id, style, numbering, table cell
and location, so headings stay headings, lists keep their numbers and footers stay footers.
Only the words change.

The difficulty is inside a paragraph. A sentence such as "Payment is due within **30 days**
of invoice." holds three runs, and a translation usually moves the bold words. `inspect_document`
marks them:

```json
{ "paraId": "w14:…", "text": "Payment is due within 30 days of invoice.",
  "markedText": "Payment is due within <1>30 days</1> of invoice." }
```

The model translates `markedText` and keeps each tag around the words that should carry that
formatting, then sends it as `withMarked`:

```json
{ "op": "changeText",
  "target": { "paraId": "w14:…", "expect": "Payment is due within 30 days of invoice." },
  "withMarked": "<1>Binnen 30 dagen</1> na factuurdatum moet worden betaald.",
  "mode": "Direct" }
```

The bold run now holds "Binnen 30 dagen" at the start of the sentence. Untagged text takes the
paragraph's usual formatting. A paragraph with one formatting throughout has no `markedText`
and is translated with a plain `with`. The full rules are in
[document plans](document-plans.md#rewrite-a-paragraph-and-keep-its-formatting).

Put many paragraphs in one plan. In Word, choose `Direct` for a clean translated copy or
leave the connection's default, `Tracked`, to get a redline a reviewer can accept. Save with
`saveMode` `NewVersion` to keep the original beside the translation.

In a workbook, `setCell` each text cell. Cell formatting belongs to the cell, so it stays,
and formulas are left alone.

### What a translation does not change

- The language tag on each run stays the source language, so Word's spelling checker may
  flag the translated text until the language is set in Word.
- A Word paragraph whose text crosses a hyperlink, content control or pending tracked change
  has no `markedText` and refuses `withMarked`. Translate the parts on either side of the
  link or control with `with`, or accept or reject the pending revisions first.
- Comments, chart labels and SmartArt are not paragraphs, so a translation leaves them as
  they are. Text boxes in a Word document are paragraphs and are translated with the rest.
- Length is not checked: a translation that runs longer can push a slide's text out of its
  box or a table onto another page. Render the result when layout matters; see
  [visual rendering](rendering.md).

## Edit many documents at once

`edit_documents` applies one plan to up to 100 documents named by source, each exactly as
`edit_document` takes it. The agent lists the folder first, with its own file tools.

```json
{
  "connectionId": "documents",
  "sources": ["contracts/msa-2024.docx", "contracts/nda-acme.docx", "decks/q3-review.pptx"],
  "planJson": "[{ \"op\": \"changeText\", \"target\": { \"find\": \"Acme Corp\", \"match\": \"all\", \"caseSensitive\": true, \"wholeWord\": true }, \"with\": \"Globex Inc.\" }]",
  "preview": true
}
```

`"match": "all"` repeats the operation for every occurrence, including occurrences in
headers, footers and notes. Text that appears nowhere in a document is skipped there. `find`
ignores case and matches inside words unless the target sets `caseSensitive` or `wholeWord`;
a rename usually wants both, so "ACME CORP" in a capitalised heading and "Acme Corporation"
are left alone. `"match": "all"` works in `edit_document` and `edit_document_content` too.

Run it with `preview: true` first. Every document is opened, matched and validated, and
nothing is written. Then run it again without `preview`. The result lists every document:

```json
{
  "isValid": false,
  "preview": false,
  "summary": { "total": 3, "committed": 1, "previewed": 0, "unchanged": 1, "failed": 1, "notAttempted": 0 },
  "items": [
    { "source": "contracts/msa-2024.docx", "status": "committed", "committed": true, "receipt": { … }, "changes": [ … ], … },
    { "source": "contracts/nda-acme.docx", "status": "failed", "committed": false, "errors": [ … ], … },
    { "source": "decks/q3-review.pptx", "status": "unchanged", "committed": false, "errors": [ … ], … }
  ]
}
```

Each item is `source` and `status` followed by exactly what `edit_document` returns for that
document, receipt included.

| Status | Meaning |
| --- | --- |
| `committed` | The document was saved. Its receipt hashes the plan as applied there and the exact input and output bytes. |
| `previewed` | `preview: true` validated it; nothing was written. |
| `unchanged` | No `"match": "all"` target matched, so there was nothing to do and the document was not saved. |
| `failed` | The document was not changed. `errors` says why; `writeOutcome` says what storage holds, as for `edit_document`. |
| `notAttempted` | The call was cancelled before this document was reached. |

Each document is edited on its own and atomically. One that fails leaves the others saved, so
re-run only the failed ones once their problem is fixed; the ones already done no longer match.
A `writeOutcome` of `unknown` or `writtenNotRegistered` needs reconciling before a retry, as
described in [storage outcomes and recovery](recovery.md).

### Limits

- At most 100 documents per call, each named once. Split a larger folder across calls.
- The plan cannot carry a `snapshot`, because a snapshot describes one document. Every
  operation is still verified against each document's live content. A server started with
  `RequirePlanSnapshot` refuses `edit_documents` for that reason.
- `edit_documents` is published with the other source-addressed tools, under
  `AllowRegistration`, and needs the same capabilities on the connection as `edit_document`.
- A logo or picture differs per document, so replacing one is still a per-document edit:
  inspect each document for its `image#N` node.
