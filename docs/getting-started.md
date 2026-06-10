# Getting started

This guide registers an existing Word document with a provider connection, inspects it, finds text, previews a tracked change, and reads the saved revision back - all by the document's opaque, provider-assigned id.

## Prerequisites

- .NET 8 SDK (or a later 8.x).
- A Word document (`.docx`).

Install the packages in your project:

```bash
dotnet add package OfficeAgent.Core
dotnet add package OfficeAgent.Word
```

## 1. Compose the client

Register the Word format module, a filesystem provider connection, and the engine:

```csharp
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Word;

var services = new ServiceCollection()
    .AddWordFormat()
    .AddFileSystemDocumentProvider("contracts", "/srv/officeagent/contracts")
    .AddOfficeAgent()
    .BuildServiceProvider();

var client = services.GetRequiredService<OfficeAgentClient>();
```

`OfficeAgentClient` is safe for concurrent use; share one per host.

## 2. Register the document with storage

The provider is a registry of references - it persists the path you hand it, not the bytes. Drop the file under the connection's configured root and register it:

```csharp
File.Copy("contract.docx", "/srv/officeagent/contracts/contract.docx", overwrite: true);
var doc = await client.RegisterAsync("contracts", "/srv/officeagent/contracts/contract.docx");
Console.WriteLine($"Registered as {doc.ItemId}");
```

From here on the document is addressed by `(connectionId, documentId)`. The provider routes reads and saves back to the registered path; the host stays in control of the file's lifecycle.

## 3. Inspect

```csharp
var inspect = await client.InspectAsync("contracts", doc.ItemId);

foreach (var paragraph in inspect.Paragraphs)
    Console.WriteLine($"{paragraph.ParaId}: {paragraph.Text}");
```

Returns paragraph ids, the outline, styles, content controls, nodes (tables / images / document properties / revisions), and a snapshot etag used for drift detection.

## 4. Find text

```csharp
var hits   = await client.FindAsync("contracts", doc.ItemId, new FindQuery("Acme Corp"));
var anchor = hits.First().Anchor;
```

Find results are *content-verified anchors*: a text-span anchor stores the paragraph id, the expected text, and the occurrence number. If the document changes underneath you before commit, the operation fails safely rather than editing the wrong run.

## 5. Preview

```csharp
var plan = new DocumentPlan
{
    Snapshot   = inspect.Snapshot,             // opt in to drift detection
    Operations = new PlanOperation[]
    {
        new ChangeTextOp { Target = anchor, With = "Globex Inc.", Mode = ChangeMode.Tracked }
    }
};

var report = await client.PreviewAsync("contracts", doc.ItemId, plan);

foreach (var change in report.Changes)
    Console.WriteLine($"{change.Verb}: {change.Before} -> {change.After}");
```

Preview validates the whole plan and reports the proposed before/after without writing.

## 6. Commit

```csharp
var result = await client.CommitAsync("contracts", doc.ItemId, plan);

if (!result.Committed)
{
    foreach (var error in result.Report.Errors)
        Console.WriteLine($"{error.Code}: {error.Message}");
    return;
}

// SaveMode.NewVersion (default) mints a fresh id for the result; the source is preserved.
Console.WriteLine($"Saved revision {result.Document.ItemId}");

using var saved = await client.OpenReadAsync(result.Document);
// saved.Stream contains the edited document. The host can now send this stream
// through its HTTP response or chat-platform attachment API.
```

Commit is all-or-nothing: if validation or application fails, no partial edit is written.

In agent-driven workflows, deliver results by id, not by bytes - see [Agent integration](agent-integration.md).

## 7. Unregister a document

`RemoveAsync` drops the id from the provider's registry; the underlying file the provider only referenced is left untouched, so the host owns cleanup.

```csharp
await client.RemoveAsync("contracts", result.Document.ItemId);
File.Delete("/srv/officeagent/contracts/contract.v2.docx");   // host deletes the file
```

`OfficeAgentTools` deliberately does not expose registration or deletion to the agent - the host pre-registers documents and removes them through application code.

## Next steps

- [Concepts](concepts.md) - anchors, plans, snapshots, capabilities, transactions.
- [Document plans](document-plans.md) - the JSON shape of every supported verb.
- [Document providers](document-providers.md) - registration, save modes, optimistic concurrency.
- [Agent integration](agent-integration.md) - exposing OfficeAgent as Microsoft Agent Framework tools.
