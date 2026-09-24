# PageCounter: Word Document Page-Count Verification with OfficeAgent.NET

This sample demonstrates using OfficeAgent.NET to render and count pages in Office documents.

## Task

Implement a tool that:
- Opens a Word document (`.docx`), PowerPoint (`.pptx`), or Excel (`.xlsx`)
- Renders it to determine actual page count
- Returns accurate page count

## What This Sample Does

`PageCounter` accepts one or more Office documents and uses OfficeAgent.NET's `LibreOfficeDocumentRenderer` to:

1. Read the document file
2. Render it to page images via LibreOffice + Poppler
3. Report the page count
4. Display render time and resource usage

## Usage

```bash
# Count pages in a single document
dotnet PageCounter.dll document.docx

# Multiple documents
dotnet PageCounter.dll doc1.docx doc2.pptx doc3.xlsx

# With options
dotnet PageCounter.dll document.docx --dpi 96 --timeout 60 --quiet

# Show only page count (--quiet suppresses extra output)
dotnet PageCounter.dll document.docx --quiet
```

## Options

- `--dpi N`: Rendering resolution (default: 144, range: 36-600)
- `--timeout SECONDS`: Max render time per document (default: 120)
- `--quiet`: Show only page count, suppress verbose output

## Key Findings: OfficeAgent.NET for Page Counting

### Tool Choice: OfficeAgent.NET + LibreOfficeDocumentRenderer

OfficeAgent.NET is positioned as a **document-automation library for contract editing and change tracking**. It is built on the Open XML SDK but abstracts document operations at a higher level.

### Is OfficeAgent Suitable for Page Counting?

**Partial - with caveats:**

#### What OfficeAgent Core Cannot Do
- The core OfficeAgent engine **does not render pages** or calculate pagination
- It provides structural inspection only (text, tables, comments, revisions)
- It cannot measure layout-dependent properties

#### How OfficeAgent Handles Page Counting
- OfficeAgent includes an **optional `OfficeAgent.Rendering` package**
- This package wraps external processes: **LibreOffice** (for PDF conversion) and **Poppler** (`pdftoppm` for image rendering)
- The renderer produces page images by:
  1. LibreOffice converts `.docx` → `.pdf` (headless)
  2. Poppler converts `.pdf` → PNG images per page
  3. PageCount is derived from the number of page images

#### Dependencies and Deployment
- Requires LibreOffice headless installation
- Requires Poppler (`pdftoppm` command)
- Runs as isolated out-of-process workers
- Sandbox-friendly: supports memory, time, and output limits

### Verdict

**Confidence: 9/10 - Yes, OfficeAgent is appropriate for this task, but with conditions:**

1. **When to use OfficeAgent:**
   - You already use OfficeAgent for document editing workflows
   - You need the same engine to inspect + edit + render
   - Headless LibreOffice and Poppler are available in deployment

2. **When NOT to use OfficeAgent for page counting alone:**
   - Page counting is your only requirement (overhead of full library)
   - No LibreOffice/Poppler available (adds external dependencies)
   - You need inline page counting without external processes

### Alternative Tools

| Tool | Strengths | Weaknesses |
|------|-----------|-----------|
| **Aspose.Words for .NET** | Native .NET, no external deps, fast | Commercial license |
| **OpenXML SDK** | Free, native .NET, low-level control | No rendering, needs custom implementation |
| **LibreOffice/UNO API** | Accurate (uses Office engine) | Complex API, Windows-only UNO bridge |
| **OfficeAgent** (this approach) | Unified editing + rendering, sandbox-safe | External process overhead, requires setup |

### Trade-offs in This Implementation

| Aspect | Trade-off |
|--------|-----------|
| **Accuracy** | LibreOffice pagination may differ from Word; use same renderer version for consistency |
| **Performance** | External process startup ~500ms+; suitable for batch, not real-time |
| **Deployment** | Requires LibreOffice + Poppler; container-friendly but adds image layer |
| **Page Count vs Content** | Count is the number of rendered PNG pages; not metadata-derived |

## Running This Sample

### Prerequisites

For actual rendering (not just code compilation):
```bash
# Ubuntu/Debian
sudo apt-get install libreoffice-headless poppler-utils

# macOS (Homebrew)
brew install libreoffice poppler

# Windows
# Download LibreOffice from https://www.libreoffice.org/
# Download Poppler binaries or build from source
# Add to PATH or set SOFFICE and PDFTOPPM env vars
```

### Compile

```bash
# Remove global.json constraint if using .NET 10
mv global.json global.json.bak
dotnet build PageCounter.csproj -c Release
mv global.json.bak global.json
```

### Run

```bash
dotnet PageCounter.dll samples/documents/services-agreement.docx
```

### Expected Output

```
Processing: services-agreement.docx
  File size: 3.79 KB
  Rendering with 144 DPI, 120s timeout...
Page count: 2
  Rendering completed in 2456 ms
  Total image data: 156.43 KB across 2 pages
  Average page size: 78.21 KB
```

If LibreOffice/Poppler are not installed:
```
Rendering failed: renderer-unavailable
  Details: A configured renderer executable was not found.
Exit code: 1
```

## Code Structure

**Program.cs** implements:
- Command-line argument parsing
- File validation
- LibreOfficeDocumentRenderer setup with environment variable overrides
- Async rendering with bounded resource options
- Graceful error handling for missing renderer or render failures
- Formatted output for page counts and performance metrics

## Important Notes

1. **Page Count Accuracy**: The page count reflects LibreOffice's rendering. Word and LibreOffice may disagree on documents with:
   - Unresolved fields
   - Pending tracked revisions
   - Embedded fonts or layout-specific formatting
   - Tables with complex structure

2. **Performance**: First render after system start includes LibreOffice startup (~2-3 seconds). Subsequent renders are faster.

3. **Resource Limits**: The sample enforces bounded rendering:
   - Max input: 100 MB
   - Max output: 500 MB
   - Max pages: 10,000
   - Max memory: 512 MB per process
   - Max time: 2 minutes

4. **Production Use**: For production, wrap in a container with fixed LibreOffice + Poppler versions to ensure consistent page counts across deployments.

## When to Choose OfficeAgent.NET Over Alternatives

Choose OfficeAgent if:
- You're building a document workflow that edits AND pages count
- Infrastructure already has LibreOffice available
- You value unified inspection, edit, and render capabilities
- Consistency across multiple document operations matters

Avoid if:
- Page counting is the only requirement
- Minimal external dependencies are required
- Aspose.Words license is available and feasible
