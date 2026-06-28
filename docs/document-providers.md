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

Connection ids are host-chosen and must be unique across providers: the `(connectionId, documentId)` overloads resolve the provider from the connection id alone, whatever its type.

## The SharePoint provider

`OfficeAgent.SharePoint` ships an `IDocumentProvider` over the Microsoft Graph drive-item REST endpoints - one connection across the tenant's drives, no Graph SDK dependency:

```csharp
using OfficeAgent.SharePoint;

services.AddSingleton(new HttpClient());
services.AddSingleton<IAccessTokenProvider>(sp => new AppOnlyAccessTokenProvider(
    new AppOnlyOptions
    {
        TenantId     = configuration["Graph:TenantId"]!,
        ClientId     = configuration["Graph:ClientId"]!,
        ClientSecret = configuration["Graph:ClientSecret"]!   // secret store / env, not source
    },
    sp.GetRequiredService<HttpClient>()));
services.AddSharePointDocumentProvider(
    connectionId: "legal",
    o => o.AllowedExtensions = new[] { ".docx" });
```

The same boundary rules apply, over Microsoft Graph:

- **A document is registered by a URL or a `driveId/itemId` pair.** A registration source is either the document's SharePoint/OneDrive URL (`https://contoso.sharepoint.com/:w:/s/…`) - resolved through Graph's `/shares` endpoint - or a `driveId/itemId` pair (`b!9a3f…/01ABCDEF`). The connection is **not** pinned to a single drive: any drive the configured identity (or, under On-Behalf-Of, the signed-in user) can reach is registrable, bounded only by that identity's Graph permissions. Disallowed extensions and oversized files are still rejected.
- **Versions are ETags.** Opens and saves check the drive item's ETag; `Replace` sends `If-Match`, and a mismatch raises `DocumentVersionConflictException` instead of overwriting newer content.
- **Save modes match the filesystem provider**: `NewVersion`/`NewDocument` upload a versioned sibling (`contract.v2.docx`, …) with `conflictBehavior=fail` and mint a fresh id; `Replace` overwrites in place.
- **Tokens stay out of references and tool results.** Authentication is an `IAccessTokenProvider` - two are included (see below). Content downloads use Graph's pre-authenticated URL, so the bearer token never travels to the storage endpoint.
- **Registrations are pluggable.** `InMemoryRegistrationStore` is the default; `JsonFileRegistrationStore` survives restarts; implement `ISharePointRegistrationStore` over shared storage for multi-instance hosts.

### Choosing an authentication mode

| Provider | Identity used | When |
| --- | --- | --- |
| `OnBehalfOfAccessTokenProvider` | **The signed-in user** | Hosted multi-user agents (Copilot Studio, M365 Copilot). The provider exchanges the caller's inbound token for a Graph token via the OAuth2 On-Behalf-Of flow, so SharePoint permissions are enforced per user. |
| `AppOnlyAccessTokenProvider` | A shared app identity | Unattended / daemon hosts where no user is present. Scope the app registration narrowly (`Sites.Selected`). |

To plug in your own credential flow (Azure.Identity, MSAL, managed identity, or a token you already hold), implement `IAccessTokenProvider` and register it directly.

**On-Behalf-Of** is the option to reach for when many users share one hosted agent. It needs the inbound user token's audience to be your middle-tier API (the `ClientId` you configure), the user's prior consent to that API, and the API holding the matching delegated Graph permission. The host must capture each caller's token into `GraphUserContext` for the request - the OfficeAgent MCP HTTP server does this automatically from the inbound `Authorization` header. Wire it in code with:

```csharp
services.AddSingleton(new HttpClient());
services.AddSharePointOnBehalfOfAuthentication(o =>
{
    o.TenantId     = configuration["Graph:TenantId"]!;
    o.ClientId     = configuration["Graph:ClientId"]!;     // the middle-tier API app
    o.ClientSecret = configuration["Graph:ClientSecret"]!; // secret store / env, not source
});
services.AddSharePointDocumentProvider("legal");
```

On a stdio host there is no inbound user token, so On-Behalf-Of is a hosted-HTTP concern; a call with no captured token fails with a clear error rather than silently falling back to an app identity.

Other cloud or document-management integrations implement `IDocumentProvider` and register one singleton per connection. Credentials live in host configuration, managed identity, delegated identity, or a secret store - **never** on `DocumentReference`.

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

The provider drops the id → path mapping; the underlying file the provider only referenced is left untouched, so the host owns the file's lifecycle. Removing an unknown id throws with `ProviderErrorCode.NotFound`. `OfficeAgentTools` never exposes content deletion to the agent; hosts may opt in to the `register_document` / `remove_document` tools (see [agent integration](agent-integration.md)), which manage registrations only.

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
