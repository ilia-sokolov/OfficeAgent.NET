# QuickEdit

This sample applies one deterministic tracked change through the direct .NET API. It uses no
MCP client and no language model.

From the repository root, run it against the bundled fictional contract:

```bash
dotnet run --project samples/QuickEdit -- \
  samples/documents/services-agreement.docx quickedit-output.docx
```

PowerShell:

```powershell
dotnet run --project samples/QuickEdit -- `
  samples/documents/services-agreement.docx quickedit-output.docx
```

The default operation changes `within thirty days of receipt` to
`within forty-five days of receipt`. Open `quickedit-output.docx` in Word and verify that:

- clause 3 contains the replacement as a tracked change;
- Word displays `QuickEdit` as the author of that replacement;
- the existing interest-rate revision and Priya Raman's comment remain present;
- the table and heading structure are unchanged.

Pass two additional arguments to edit another document with an exact text match:

```bash
dotnet run --project samples/QuickEdit -- input.docx output.docx "Acme Corp" "Globex Inc."
```

The sample exits with a non-zero status and does not write an output when the exact source
text is absent, preview fails, or the commit is rejected.
