# TemplateBatch

This offline sample creates a small quote template, binds one tagged scalar and a repeating
Word table row, and writes two independent `.docx` outputs. The source template remains
unchanged and every output has its own apply receipt.

```powershell
dotnet run --project samples/TemplateBatch -- generated-quotes
```

Open the two files in `generated-quotes`. The sample uses direct edits because these are
generated documents; set `TemplateBinding.Mode` to `Tracked` when a human must review the
population as revisions.
