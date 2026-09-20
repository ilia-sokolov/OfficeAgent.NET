# QuickEdit package sample

This standalone sample applies one deterministic tracked change using published
OfficeAgent.NET packages. It has no repository-relative project references.

The archive includes a fictional input document. From the extracted
`quickedit-sample` directory, run:

```bash
dotnet restore
dotnet run -- services-agreement.docx quickedit-output.docx
```

The default operation changes `within thirty days of receipt` to
`within forty-five days of receipt`. Open `quickedit-output.docx` in Word and
verify that the replacement is attributed to `QuickEdit`, the existing revision
and comment remain present, and the document opens without a repair prompt.

Pass two additional arguments to edit another document with an exact text match:

```bash
dotnet run -- input.docx output.docx "Acme Corp" "Globex Inc."
```

The sample exits with a nonzero status and does not write an output when the
exact source text is absent, preview fails, or the commit is rejected.
