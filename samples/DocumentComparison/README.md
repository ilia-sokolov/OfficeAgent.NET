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

Version 0.8 comparison covers free body-paragraph text and paragraph insertions/removals.
Existing revisions or changes in tables, images, headers, footers, notes, styles, fields,
properties, or package metadata return diagnostics and no plan.
