# TableRowInsertion

Adds a row to a Word table as a tracked change, and checks that the rest of the document came
through: the existing comment, the table, and a package the Open XML validator accepts.

The sample builds its own small contract, so it needs no input file: a "Contract Line Items"
heading with a reviewer's comment on it, a three-column table of line items with a header row
and two data rows, and a closing paragraph. Then it:

1. inspects the document and counts its paragraphs and comments;
2. commits one `InsertTableRowsOp` that targets the first table (`table#0`), adds
   `New Contract Item | $50,000 | High Priority` at the end, and sets `Mode = ChangeMode.Tracked`
   with `ContractEditor` as the revision author;
3. validates the saved package against Office 2019 with `OpenXmlValidator`;
4. inspects the result again and confirms the comment is still there; and
5. reads the last row's `w:trPr` and prints its `w:ins` mark.

The new row copies the table's last row, and each value takes the formatting of that cell's
first run. It uses the direct API (`new OfficeAgentClient(new WordModule())` over a
`StreamHandle`), with no storage provider or dependency injection.

## Run it

From the repository root:

```powershell
dotnet run --project samples/TableRowInsertion
```

The result is written to `contract_with_inserted_row.docx` in the system temporary folder, and
the sample prints the full path. Expected output:

```text
Creating sample contract document with table and comment...
  - Created document: 2493 bytes
  - Found 11 paragraphs
  - Found 1 comments

Inserting new row with tracked changes...
  - Document updated: 2635 bytes

Validating document...
  - Document is valid

Inspecting result...
  - Paragraphs: 14
  - Comments preserved: 1
  - Last row marked as inserted (id=1, author=ContractEditor)
```

The sample exits with 1, and writes nothing, if the commit is refused or the validator reports
an error.

Open the file in Word and choose **Review > Reviewing Pane**: the table's new last row is
shown as an insertion by ContractEditor, which you can accept or reject, and the comment on the
heading is unchanged.

## What to take from it

- A row insertion is a tracked change like a text edit: the row carries `w:ins` in its row
  properties, so rejecting it removes the whole row. See
  [`insertTableRows`](../../docs/document-plans.md#inserttablerows--removetablerows).
- A table is addressed by its node path, `table#0` for the first table in the body, which
  inspection lists.
