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
drains process output through fixed-size buffers without retaining it, samples each renderer
process's working set, and watches elapsed time, page count, and output bytes. The single timeout
covers conversion, rasterization, and redirected-output draining. A limit violation stops the
tracked renderer process tree and returns a stable failure code with no partial pages.
Cancellation also stops the tracked process tree. Temporary input and output are deleted after
the result when the operating system releases them.

The working-set check is a process-level guard. For documents from untrusted callers, run the
host in an OS container or job boundary that applies hard aggregate CPU, memory, filesystem, and
process limits to LibreOffice and all descendants. The in-process monitor cannot guarantee that
an already-detached descendant is terminated or included in the working-set sample. Keep renderer
workers away from service credentials and network access. Do not log renderer stdout, stderr,
input bytes, or rendered pages; they may contain document content or sensitive paths.

The renderer reports `renderer-unavailable`, `renderer-failed`, `renderer-output-missing`,
`input-limit-exceeded`, `render-timeout`, `memory-limit-exceeded`, `page-limit-exceeded`,
`output-limit-exceeded`, or `invalid-render-options`. The first implementation returns page
images. Visual overflow detection requires a chosen renderer and a representative document
corpus and is not inferred from these images yet.

## Reading the result

`RenderAsync` does not throw when a renderer is missing or a limit is hit. It returns a
`RenderResult` that is either a success carrying pages or a failure carrying a code — the two
states cannot be mixed, and a failed result carries no pages.

Because a page count drives decisions such as "refuse anything over four pages", reading pages
from a failed render is a mistake that must not read as an empty document. `Pages` and
`PageCount` therefore throw `RenderFailedException`, carrying the failure code, unless rendering
succeeded:

```csharp
var result = await renderer.RenderAsync(input, options);

// Fails closed: on a container with no LibreOffice installed this throws
// RenderFailedException("renderer-unavailable"), it does not report zero pages.
if (result.PageCount > 4) return Reject("The contract exceeds four pages.");
```

To branch on the failure instead of catching it, check `Succeeded`, call `EnsureSucceeded()`, or
use `TryGetPages`:

```csharp
if (!result.TryGetPages(out var pages))
{
    logger.LogError("Rendering failed: {Code}", result.FailureCode);
    return Reject("The page count could not be established.");   // never admit on failure
}
```

A page count is the pagination of whichever renderer produced it. LibreOffice and Word can
disagree on documents containing tables, unembedded fonts, unresolved fields, or pending
revisions, so pin the renderer version and the container's font set, and report the count
alongside the renderer that produced it rather than as an absolute property of the file.
Resolve revisions before measuring: an unaccepted insertion changes text length and therefore
pagination.
