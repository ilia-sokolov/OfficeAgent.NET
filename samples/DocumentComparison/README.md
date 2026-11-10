# DocumentComparison

This sample compares two Word files without writing either input. When all detected changes
are covered, it previews the returned snapshot-bound plan and writes a third document whose
body paragraph changes are native Word revisions.

```powershell
dotnet run --project samples/DocumentComparison -- original.docx revised.docx redline.docx
```

Or run the self-contained fixture:

```powershell
dotnet run --project samples/DocumentComparison -- --demo redline.docx
```

The comparison covers body-paragraph text, paragraph insertions and removals, and the text of
table cells whose table kept its shape. Runs that Word split differently but format the same
are not a difference. An input that already carries tracked revisions, and any other change,
such as table rows or columns, cell or paragraph formatting, styles, headers, footers, notes,
images or package metadata, return diagnostics and no plan. The coverage gate is deliberately strict and can reject metadata-only changes that
do not affect visible output. See
[what the comparison understands](../../docs/document-workflows.md#what-the-comparison-understands).
