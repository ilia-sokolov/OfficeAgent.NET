# Choose OfficeAgent.NET

OfficeAgent.NET is a document-automation library for applications that need a
structured way to inspect and edit OOXML files. It is built on the Open XML SDK,
and its direct .NET API does not require MCP, an LLM, or an agent framework.

> [!IMPORTANT]
> This guide is versioned for OfficeAgent.NET `0.9.0`. Its executable recipe pins
> `0.9.0` and was verified against the release-candidate packages on 2026-09-16.
> Before the release appears on NuGet, use the locally packed candidate feed described
> in the release runbook. Keep every OfficeAgent package on the same version.

The supported-job summaries and recipe describe the source and package contract in
the `v0.9.0` tag. For another release, use the documentation from that release's tag.

## Use OfficeAgent.NET when

- Your .NET application needs typed inspect, find, preview, and commit operations
  over Word `.docx`, PowerPoint `.pptx`, or Excel `.xlsx` packages.
- You want callers to target engine-issued anchors instead of constructing OOXML
  paths or offsets themselves.
- You need a complete plan to validate before an atomic edit, with stable errors
  such as `stale-snapshot`, `expect-mismatch`, and `unsupported-operation`.
- You need supported Word review operations, including comments, revisions, and
  tracked forms of supported edits.
- You want the same engine behind a direct .NET application and, optionally, MCP
  or `Microsoft.Extensions.AI` tools.

The complete operation shapes and format-specific behavior are in
[Document plans](document-plans.md), [PowerPoint support](powerpoint.md), and
[Excel support](excel.md). The public
[Word preservation matrix](word-preservation-evidence.md) shows exact
operation-specific package evidence and its native Office limits.

## Do not use it as

- A general editor for `.docm`, `.xlsm`, `.pptm`, legacy Office formats, PDF, or
  arbitrary OOXML features outside the documented operation set.
- A Word or PowerPoint layout and pagination engine.
- An Excel formula calculation or data-refresh engine. Formula operations write
  formulas and request recalculation when Excel next opens the workbook.
- A substitute for native Office opening, rendered visual review, or human review
  of consequential output.
- An authentication or authorization layer. The host owns identities, credentials,
  allowed storage roots, and access policy.

Core previews are structural. The optional rendering package runs external
LibreOffice and Poppler processes, but it does not provide Office's native layout
engine or Excel calculation. See [Scope and limitations](../README.md#scope-and-limitations)
and [Support and compatibility](../SUPPORT.md) before production use.

## Choose the integration level

| Need | Choose | Packages |
| --- | --- | --- |
| Direct Word automation in a .NET application | OfficeAgent direct API | `OfficeAgent.Core` and `OfficeAgent.Word` |
| Direct PowerPoint automation | OfficeAgent direct API | `OfficeAgent.Core` and `OfficeAgent.PowerPoint` |
| Direct Excel automation | OfficeAgent direct API | `OfficeAgent.Core` and `OfficeAgent.Excel` |
| `Microsoft.Extensions.AI` or Microsoft Agent Framework tools | Optional adapter over the same client | `OfficeAgent.AgentFramework` plus the required format modules |
| A standalone MCP process | Optional server over the same engine | `OfficeAgent.Mcp` |
| Low-level package-part or schema work outside OfficeAgent operations | Open XML SDK directly | `DocumentFormat.OpenXml` |
| Both, in one application | [SDK interoperability recipe](sdk-interoperability.md) | OfficeAgent packages plus `DocumentFormat.OpenXml` |

Add a storage provider only when the hosting model needs it. `OfficeAgent.Core`
includes the filesystem and in-memory provider surfaces. `OfficeAgent.SharePoint`
adds Microsoft Graph storage. A caller that already owns the bytes can use a
`StreamHandle` without registering storage.

## OfficeAgent.NET or the Open XML SDK directly

OfficeAgent.NET uses `DocumentFormat.OpenXml` internally. The choice is therefore
about abstraction level, not competing file formats.

Choose OfficeAgent.NET when its operation set matches the job and you want its
anchors, previews, plan validation, atomic apply behavior, and format-specific
handlers. Choose the Open XML SDK directly when you need precise control of parts
and schema elements, an operation OfficeAgent does not implement, or a custom
package transformation whose invariants you will own and test.

This comparison is based on the following primary documentation, checked on
2026-09-14:

- Microsoft's [Open XML SDK overview](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
  describes strongly typed classes for manipulating packages and schema elements.
- Microsoft's [Open XML SDK design considerations](https://learn.microsoft.com/office/open-xml/open-xml-sdk-design-considerations)
  state that the SDK is not an abstraction over the file formats, does not
  guarantee validity, and does not provide Word layout or Excel recalculation.

There is no popularity or product-ranking claim here. Evaluate the actual
operation and deployment requirements of your application.

You can also combine the two. The tested copy, edit, validate, reinspect workflow and
the exact guarantees that apply at each stage are in
[Open XML SDK interoperability](sdk-interoperability.md). Anchors, snapshots, approved
previews, and receipts do not carry across an SDK edit, and there is no supported path
that replaces a registered document with externally edited bytes through OfficeAgent.

## Verified direct Word edit

This self-contained trial creates a fictional input file, finds a company name,
previews a tracked replacement, commits it, and checks both the document semantics
and its Office 2019 Open XML schema. It then previews an Excel-only operation
against the Word file to demonstrate the fail-closed unsupported result.

Create the project and install the exact packages:

```bash
dotnet new console -n OfficeAgentSelectionTrial
cd OfficeAgentSelectionTrial
dotnet add package OfficeAgent.Core --version 0.9.0
dotnet add package OfficeAgent.Word --version 0.9.0
```

Replace `Program.cs` with:

```csharp
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Word;

const string oldName = "Northwind Labs";
const string newName = "Contoso Research";
var inputPath = Path.GetFullPath("proposal.docx");
var outputPath = Path.GetFullPath("proposal-edited.docx");

using (var document = WordprocessingDocument.Create(
           inputPath, WordprocessingDocumentType.Document))
{
    var main = document.AddMainDocumentPart();
    main.Document = new Document(
        new Body(new Paragraph(new Run(new Text($"Prepared for {oldName}.")))));
}

var inputBytes = await File.ReadAllBytesAsync(inputPath);
using var input = new MemoryStream(inputBytes, writable: false);
var handle = new StreamHandle(input, Path.GetFileName(inputPath));
var client = new OfficeAgentClient(new WordModule());

var inspection = await client.InspectAsync(handle);
var hit = (await client.FindAsync(handle, new FindQuery(oldName))).Single();
var plan = new DocumentPlan
{
    Snapshot = inspection.Snapshot,
    Revision = new RevisionMetadata { Author = "OfficeAgent Selection Trial" },
    Operations = new PlanOperation[]
    {
        new ChangeTextOp
        {
            Target = hit.Anchor,
            With = newName,
            Mode = ChangeMode.Tracked
        }
    }
};

var preview = await client.PreviewAsync(handle, plan);
Console.WriteLine($"preview-valid={preview.IsValid}");
Console.WriteLine($"preview-change={preview.Changes.Single().Before}->{preview.Changes.Single().After}");

using var result = await client.CommitAsync(handle, plan);
if (!result.Committed)
    throw new InvalidOperationException(string.Join("; ", result.Report.Errors.Select(e => $"{e.Code}: {e.Message}")));

await result.SaveAsync(outputPath);
Console.WriteLine($"committed={result.Committed}");

using (var edited = WordprocessingDocument.Open(outputPath, isEditable: false))
{
    var body = edited.MainDocumentPart?.Document?.Body
        ?? throw new InvalidOperationException("The edited document has no main body.");
    Console.WriteLine($"output-text={string.Concat(body.Descendants<Text>().Select(t => t.Text))}");
    Console.WriteLine($"deleted-text={string.Concat(body.Descendants<DeletedText>().Select(t => t.Text))}");
    var schemaErrors = new OpenXmlValidator(FileFormatVersions.Office2019)
        .Validate(edited)
        .ToList();
    Console.WriteLine($"schema-errors={schemaErrors.Count}");
    foreach (var error in schemaErrors)
        Console.WriteLine($"schema-error={error.Description}");
}

var unsupportedPlan = new DocumentPlan
{
    Snapshot = inspection.Snapshot,
    Operations = new PlanOperation[]
    {
        new SetCellOp
        {
            Target = new CellAnchor { SheetId = 1, Address = "A1" },
            Value = "Not applicable to Word"
        }
    }
};

var unsupported = await client.PreviewAsync(handle, unsupportedPlan);
Console.WriteLine($"unsupported-valid={unsupported.IsValid}");
Console.WriteLine($"unsupported-code={unsupported.Errors.Single().Code}");
```

`Revision.Author` is the name Word displays for the tracked change. It is separate
from any authenticated actor recorded by the host. See
[revision identity and audit receipts](operations.md#revision-identity-and-audit-receipts).

Run it:

```bash
dotnet run --configuration Release
```

The expected result is:

```text
preview-valid=True
preview-change=Northwind Labs->Contoso Research
committed=True
output-text=Prepared for Contoso Research.
deleted-text=Northwind Labs
schema-errors=0
unsupported-valid=False
unsupported-code=unsupported-operation
```

`proposal-edited.docx` is the output. The new company name is normal document
text, while the old name remains in Word deletion markup for review. A zero schema
error count is structural validation, not proof of native layout or visual quality.
Open the result in the Office application used by your reviewers before adopting
the workflow for production documents.

## Version and support checks

Before upgrading, read the [version policy and support terms](../SUPPORT.md) and
the [changelog](../CHANGELOG.md). OfficeAgent is pre-1.0, so pin versions, keep
all OfficeAgent packages on one version, and test representative documents.

The package version used by the application is the package identity. A repository
branch name, unreleased documentation, or local source build does not make that
version available from NuGet.
