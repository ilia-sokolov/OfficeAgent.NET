# DocumentAssembly

Assemble a fictional proposal, statement of work, and landscape appendix into one editable
Word package. The sample uses a read-only preview, re-reads the inputs, validates their hashes,
and writes a new output plus its multi-source receipt. Existing files are never overwritten.

```powershell
dotnet run --project samples/DocumentAssembly -- --demo ./assembly-output
```

Use a fresh output directory on subsequent runs. To assemble your own compatible files:

```powershell
dotnet run --project samples/DocumentAssembly -- packet.docx proposal.docx scope.docx appendix.docx
```

Add `--render` to run the optional isolated LibreOffice/Poppler renderer and write page PNGs
next to the output. Set `SOFFICE` and `PDFTOPPM` to executable paths if they are not on PATH.
Review all pages for source formatting, section boundaries, header isolation, and table layout.
Rendering is an additional acceptance check; the core assembly workflow requires no renderer.

See [Word document assembly](../../docs/document-assembly.md) for compatibility restrictions,
provider and MCP usage, audit semantics, and the acceptance corpus.
