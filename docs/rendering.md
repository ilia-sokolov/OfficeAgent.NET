# Visual rendering

OfficeAgent's normal preview is structural: it validates a plan and reports the changes it would
make without writing the package. It remains available with no renderer installed. Applications
that also need page images can add the optional `OfficeAgent.Rendering` package and call the
`IDocumentRenderer` boundary from `OfficeAgent.Abstractions`.

```csharp
IDocumentRenderer renderer = new LibreOfficeDocumentRenderer();
await using var input = File.OpenRead("proposal.docx");

var result = await renderer.RenderAsync(input, new RenderOptions
{
    FileName = "proposal.docx",
    Timeout = TimeSpan.FromSeconds(45),
    MaximumInputBytes = 25 * 1024 * 1024,
    MaximumWorkingSetBytes = 512L * 1024 * 1024,
    MaximumPages = 100,
    MaximumOutputBytes = 50L * 1024 * 1024
});
```

`LibreOfficeDocumentRenderer` accepts `.docx`, `.pptx`, and `.xlsx`. It runs LibreOffice in a
new temporary profile to produce PDF, then runs Poppler's `pdftoppm` to produce ordered PNG page
images. Configure `LibreOfficeExecutable` and `PdfToPpmExecutable` when those commands are not on
`PATH`.

The implementation copies a bounded input into a unique temporary directory, redirects and
discards process output, samples each renderer process's working set, and watches elapsed time,
page count, and output bytes. A limit violation kills the launched process tree and returns a
stable failure code with no partial pages. Cancellation also kills the process tree. Temporary
input and output are deleted after every result.

The working-set check is a process-level guard. For documents from untrusted callers, run the
host in an OS container or job boundary that applies hard aggregate CPU, memory, filesystem, and
process limits to LibreOffice and its descendants. Keep renderer workers away from service
credentials and network access. Do not log renderer stdout, stderr, input bytes, or rendered
pages; they may contain document content or sensitive paths.

The renderer reports `renderer-unavailable`, `renderer-failed`, `renderer-output-missing`,
`input-limit-exceeded`, `render-timeout`, `memory-limit-exceeded`, `page-limit-exceeded`,
`output-limit-exceeded`, or `invalid-render-options`. The first implementation returns page
images. Visual overflow detection requires a chosen renderer and a representative document
corpus and is not inferred from these images yet.
