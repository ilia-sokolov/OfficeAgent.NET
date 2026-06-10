# Document providers

A document provider is a **registry of references**: it persists where each
document lives (a filesystem path, a URL, a drive id, …), not the bytes
themselves. Hosts register an existing document with a connection to receive an
opaque id; later reads and saves route back to the referenced location. The
agent only sees the opaque id, never a path or URL. Provider contracts and
implementations live in the `OfficeAgent.Core.DocumentProviders` namespace.

Every reference contains a provider name, a host-configured connection id, an **opaque, provider-assigned document id**, and an optional version. Ids are tokens, not addresses - a caller never constructs one. The id is minted on `RegisterAsync` and echoed back on every subsequent open and save.

## Register a connection

```csharp
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Word;

services
    .AddWordFormat()
    .AddFileSystemDocumentProvider(
        connectionId: "contracts",
        rootPath:     "/srv/officeagent/contracts")
    .AddOfficeAgent();
```

The filesystem provider only accepts source paths that lie under its configured root. It persists the id → path mapping in `{root}/.officeagent/index.json` and rejects path-like ids, traversal escapes, symlinks/reparse points, disallowed extensions, oversized files, and stale versions.

Cloud or document-management integrations implement `IDocumentProvider` and register one singleton per connection:

```csharp
services.AddSingleton<IDocumentProvider, SharePointDocumentProvider>();
```

Credentials live in host configuration, managed identity, delegated identity, or a secret store - **never** on `DocumentReference`. A cloud provider's `RegisterAsync` accepts a provider-specific locator (a sharing URL, a drive id, etc.) just as the filesystem provider accepts a path.

## Register, inspect, commit

```csharp
var client = services.GetRequiredService<OfficeAgentClient>();

// 1. Register the existing file with a connection. The provider mints the opaque id;
//    only the path is persisted, not the bytes.
var doc = await client.RegisterAsync("contracts", "/srv/officeagent/contracts/acme.docx");

// 2. Address by (connectionId, documentId) from here on.
var inspect = await client.InspectAsync("contracts", doc.ItemId);
var result  = await client.CommitAsync(
                "contracts", doc.ItemId, plan,
                new SaveDocumentOptions { Mode = SaveMode.NewVersion });

// 3. Read the saved revision back out of storage.
using var saved = await client.OpenReadAsync(result.Document);
```

A stored id can be reconstituted into a reference if you only have the id:

```csharp
var reference = DocumentReference.ForFileSystem("contracts", doc.ItemId);
```

## Remove (unregister) a document

`OfficeAgentClient.RemoveAsync` unregisters a document, by canonical reference or by `(connectionId, documentId)`:

```csharp
await client.RemoveAsync("contracts", doc.ItemId);
// or: await client.RemoveAsync(result.Document);
```

The provider drops the id → path mapping; the underlying file the provider only referenced is left untouched, so the host owns the file's lifecycle. Removing an unknown id throws with `ProviderErrorCode.NotFound`. `OfficeAgentTools` deliberately does not expose deletion to the agent - the host removes documents through application code.

## Save modes

| Mode | Behaviour |
| --- | --- |
| `NewVersion` | Default. Writes the result as `{base}.v{n}.{ext}` beside the source file, mints a fresh opaque id, and registers the new path. The source is preserved. Use `NewName` to override the destination's file name. |
| `NewDocument` | Same as `NewVersion`, named for clarity when the intent is "create a sibling document". A caller-chosen `DestinationItemId` is rejected - ids are provider-assigned. |
| `Replace` | Atomically overwrites the registered file in place (same id, same path) after an optimistic-concurrency check against `ExpectedVersion` (or the source reference's `Version`). |

Every save returns a canonical `DocumentReference` for the result, including its new id and version. A version mismatch raises `DocumentVersionConflictException` instead of overwriting newer content.

## Error contract

Failures at the provider boundary throw `DocumentProviderException` carrying a strongly typed `ProviderErrorCode` and `Provider` / `ConnectionId` / `ItemId`. A single `catch (DocumentProviderException ex)` covers every provider failure.

| Code | Cause |
| --- | --- |
| `NotFound` | Unknown / removed document id. |
| `AccessDenied` | Path-like id, traversal attempt, symlink, or wrong connection. |
| `ContentTooLarge` | Document exceeds the connection's configured size cap. |
| `ExtensionNotAllowed` | Display-name extension is not in the allow-list. |
| `VersionConflict` | Optimistic-concurrency check failed (`DocumentVersionConflictException`). |
| `InvalidArgument` | Structurally bad reference or save options. |
| `ConfigurationError` | Registry has zero or more than one provider for the requested `(provider, connectionId)`. |
| `IO` | Underlying storage I/O error. |

These flatten cleanly into the agent-tool surface: `OfficeAgent.AgentFramework` returns each one as a stable wire code (`not-found`, `access-denied`, `content-too-large`, `extension-not-allowed`, `version-conflict`, `invalid-argument`, `configuration-error`, `io-error`) so the LLM can act on it.
