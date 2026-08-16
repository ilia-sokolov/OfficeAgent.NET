# Getting started

In this tutorial, you build a console application that registers a Word
document, finds text, previews a tracked change, commits it, and writes an edited
copy you can open in Microsoft Word.

Estimated time: 10 minutes.

## Prerequisites

- [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0) or later.
- A Word `.docx` file containing the text `Acme Corp`.

Create a console project and install the engine, Word module, and concrete
dependency-injection container:

```bash
dotnet new console --framework net8.0 -n OfficeAgentQuickstart
cd OfficeAgentQuickstart
dotnet add package OfficeAgent.Core
dotnet add package OfficeAgent.Word
dotnet add package Microsoft.Extensions.DependencyInjection
```

Copy your input file to the project directory as `contract.docx`.

## 1. Configure the client

Add these namespaces to `Program.cs`:

```csharp
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Word;
```

Create a provider root using a path that works on Windows, macOS, and Linux,
then register the Word module, filesystem provider, and engine:

```csharp
var storageRoot = Path.GetFullPath("officeagent-workspace");
Directory.CreateDirectory(storageRoot);

using var services = new ServiceCollection()
    .AddWordFormat()
    .AddFileSystemDocumentProvider("contracts", storageRoot)
    .AddOfficeAgent()
    .BuildServiceProvider();

var client = services.GetRequiredService<OfficeAgentClient>();
```

`OfficeAgentClient` and the built-in providers are safe for concurrent use.
Register one client per application host.

## 2. Register the input

The filesystem provider stores an id-to-path registration, not another copy of
the document. Stage the input under its configured root and register that path:

```csharp
var stagedPath = Path.Combine(storageRoot, "contract.docx");
File.Copy("contract.docx", stagedPath, overwrite: true);

var document = await client.RegisterAsync("contracts", stagedPath);
Console.WriteLine($"Registered documentId: {document.ItemId}");
```

Later calls use `(connectionId, documentId)`. The model or calling component
does not need the storage path after registration.

## 3. Inspect and find an anchor

Inspect returns the document structure and a snapshot used for drift detection.
Find returns content-verified anchors that can safely target text across OOXML
runs.

```csharp
var inspection = await client.InspectAsync("contracts", document.ItemId);

foreach (var paragraph in inspection.Paragraphs.Take(5))
    Console.WriteLine($"{paragraph.ParaId}: {paragraph.Text}");

var hits = await client.FindAsync(
    "contracts", document.ItemId, new FindQuery("Acme Corp"));

if (hits.Count == 0)
    throw new InvalidOperationException(
        "The tutorial document must contain the text 'Acme Corp'.");

var anchor = hits[0].Anchor;
```

If the anchored content or Word text-host XML changes before commit, the anchor
or snapshot check fails instead of editing a different occurrence. Properties,
embedded bytes, and other parts outside snapshot coverage rely on their
operation-specific validation and the provider's optimistic version check.

## 4. Preview the plan

Build a typed plan and preview it. Preview validates the complete plan and does
not write to storage.

```csharp
var plan = new DocumentPlan
{
    Snapshot = inspection.Snapshot,
    Operations = new PlanOperation[]
    {
        new ChangeTextOp
        {
            Target = anchor,
            With = "Globex Inc.",
            Mode = ChangeMode.Tracked
        }
    }
};

var preview = await client.PreviewAsync("contracts", document.ItemId, plan);
if (!preview.IsValid)
{
    foreach (var error in preview.Errors)
        Console.Error.WriteLine($"{error.Code}: {error.Message}");
    return;
}

foreach (var change in preview.Changes)
    Console.WriteLine($"{change.Verb}: {change.Before} -> {change.After}");
```

## 5. Commit and save the output

Commit applies the plan atomically. `SaveMode.Replace` is the default, so the
provider updates the staged document after an optimistic version check.

```csharp
var result = await client.CommitAsync("contracts", document.ItemId, plan);
if (!result.Committed || result.Document is null)
{
    foreach (var error in result.Report.Errors)
        Console.Error.WriteLine($"{error.Code}: {error.Message}");
    return;
}

var outputPath = Path.GetFullPath("contract-edited.docx");
using (var saved = await client.OpenReadAsync(result.Document))
using (var output = File.Create(outputPath))
    await saved.Stream.CopyToAsync(output);

Console.WriteLine($"Edited document: {outputPath}");
```

Open `contract-edited.docx` in Word. The replacement appears as a tracked
deletion and insertion.

To preserve the staged source, pass
`new SaveDocumentOptions { Mode = SaveMode.NewVersion }` to `CommitAsync`. The
result then receives a new document id and a sibling name such as
`contract.v2.docx`. The cleanup in the next section shows the default `Replace`
path. If you try `NewVersion`, unregister both ids when they differ and keep or
copy `stagedPath` before deleting any host-owned files.

## 6. Unregister when finished

Unregistering removes the provider's id-to-path entry. It does not delete the
underlying document.

```csharp
await client.RemoveAsync("contracts", result.Document.ItemId);
if (result.Document.ItemId != document.ItemId)
    await client.RemoveAsync("contracts", document.ItemId);

// Optional host-owned cleanup for Replace, or when a preserved source is no longer needed:
// File.Delete(stagedPath);
```

Under `NewVersion`, the two ids refer to different files, so both registrations
are removed. Delete `stagedPath` only when you no longer need the preserved
source.

In-process registration tools are omitted by default. Set
`OfficeAgentToolsOptions.AllowRegistration = true` only when an agent should be
able to register sources and remove registrations itself. `remove_document`
still never deletes document content.

## Create a document instead

`CreateAsync` does not require an input file. The requested extension selects
the registered blank-document factory:

```csharp
var created = await client.CreateAsync("contracts", "brief.docx");
Console.WriteLine(created.Document?.ItemId);
```

A new Word document starts with one empty paragraph at `auto-0000`; an optional
initial plan can target it. For `.pptx`, install `OfficeAgent.PowerPoint`, call
`AddPowerPointFormat()`, allow `.pptx` on the connection, and see
[Creating a deck](powerpoint.md#creating-a-deck).

## Next steps

- [Concepts](concepts.md) — anchors, snapshots, plans, capabilities, and transactions.
- [Document plans](document-plans.md) — every supported operation and JSON shape.
- [Document providers](document-providers.md) — save modes and SharePoint.
- [Agent integration](agent-integration.md) — expose bounded agent tools.
- [MCP server](mcp-server.md) — use the same workflow without a .NET integration.
- [Troubleshooting](troubleshooting.md) — common setup and runtime failures.
