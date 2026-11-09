# Document providers

A document provider is a **registry of references**: it persists where each
document lives (a filesystem path, a URL, a drive id, …), not the bytes
themselves. Hosts register an existing document with a connection to receive an
opaque id; later reads and saves route back to the referenced location. The
default edit tools use only the opaque id. Opt-in registration and composite
tools accept a source path or URL from the caller, so that value can enter model
context. Provider contracts and
implementations live in the `OfficeAgent.Core.DocumentProviders` namespace.

Every reference contains a provider name, a host-configured connection id, an **opaque, provider-assigned document id**, and an optional version. Ids are tokens, not addresses. `RegisterAsync` mints an id. `Replace` preserves it; `NewVersion`, `NewDocument`, and creation mint an id for the new item.

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

The filesystem provider only accepts source paths that lie under its configured root. It persists the id → path mapping in `{root}/.officeagent/index.json` and rejects path-like ids, traversal escapes, symlinks/reparse points, disallowed extensions, oversized files, and stale versions. New documents created through the connection are written directly under that root.

Treat that root as a host-managed trust boundary. The provider checks every path
component for links before registration, open, and save, but pathname validation
and I/O are not an atomic filesystem transaction. Use filesystem ACLs so
untrusted principals cannot create, rename, or replace directory entries beneath
the root while OfficeAgent runs. A user may have write access to an existing
file without receiving directory create/delete/rename rights. Stage uploads
through the trusted host instead of giving an untrusted uploader direct directory
control.

Connection ids are host-chosen and must be unique across providers: the `(connectionId, documentId)` overloads resolve the provider from the connection id alone, whatever its type.

## The SharePoint provider

```bash
dotnet add package OfficeAgent.Core
dotnet add package OfficeAgent.SharePoint
dotnet add package Microsoft.Extensions.DependencyInjection
```

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
    o =>
    {
        o.AllowedExtensions = new[] { ".docx" };
        o.CreationDriveId = configuration["Graph:CreationDriveId"]!;
        o.CreationFolderItemId = configuration["Graph:CreationFolderItemId"]!;
    });
```

The same boundary rules apply, over Microsoft Graph:

- **A document is registered by a URL or a `driveId/itemId` pair.** A registration source is either the document's SharePoint/OneDrive URL (`https://contoso.sharepoint.com/:w:/s/…`) - resolved through Graph's `/shares` endpoint - or a `driveId/itemId` pair (`b!9a3f…/01ABCDEF`). The connection is **not** pinned to a single drive: any drive the configured identity (or, under On-Behalf-Of, the signed-in user) can reach is registrable, bounded only by that identity's Graph permissions. Disallowed extensions and oversized files are still rejected.
- **Versions are ETags.** Opens and saves check the drive item's ETag; `Replace` sends `If-Match`, and a mismatch raises `DocumentVersionConflictException` instead of overwriting newer content.
- **Save modes match the filesystem provider**: `NewVersion`/`NewDocument` upload a versioned sibling (`contract.v2.docx`, …) with `conflictBehavior=fail` and mint a fresh id; `Replace` overwrites in place.
- **Tokens stay out of references and tool results.** Authentication is an `IAccessTokenProvider` - two are included (see below). Content downloads use Graph's pre-authenticated URL, so the bearer token never travels to the storage endpoint.
- **Registrations are pluggable.** `InMemoryRegistrationStore` is the default; `JsonFileRegistrationStore` survives restarts; implement `ISharePointRegistrationStore` over shared storage for multi-instance hosts.
- **Creation has one explicit destination.** Set both `CreationDriveId` and `CreationFolderItemId` to allow new documents in that folder. Existing-document registration remains cross-drive. Uploads use `conflictBehavior=fail`, so creation never overwrites an existing SharePoint item.
- **Large files are uploaded in chunks.** A save of `UploadSessionThresholdBytes` or more (default 4 MiB) goes through a Graph upload session in chunks of `UploadSessionChunkBytes` (default 10 MiB, a multiple of 320 KiB and less than 60 MiB); a smaller one is a single request, as before. A session keeps the same conditions: `Replace` sends `If-Match`, and new names use `conflictBehavior=fail`. A chunk that fails for a dropped connection, a 5xx or a 429 is retried up to three times from the position Graph reports, honouring `Retry-After`. The bearer token is never sent to the session's pre-authenticated upload URL.
- **What a failed chunked save means.** Graph stores the file only when it accepts the last chunk. A failure before that, including an expired session, is `WriteRejected`: nothing was stored, and the provider cancels the session. A failure on the last chunk, or a lost reply to it, is `OutcomeUnknown`, as for a single-request upload; a name conflict or a failed `If-Match` on the last chunk is a refusal and reported as such.
- **The size limit is `MaximumBytes`.** Graph accepts files up to 250 GB in SharePoint and OneDrive, so the connection's own `MaximumBytes` is the limit that applies. Content that is not seekable is read into memory before it is uploaded, and the engine holds a document in memory while it edits it, so size `MaximumBytes` to the host's memory.

### Choosing an authentication mode

| Provider | Identity used | When |
| --- | --- | --- |
| `OnBehalfOfAccessTokenProvider` | **The signed-in user** | Hosted multi-user agents. Effective access is the intersection of consented delegated Graph scopes and the user's SharePoint access. |
| `AppOnlyAccessTokenProvider` | A shared app identity | Unattended hosts. Prefer `Sites.Selected`, then assign the app to each intended site with a sufficient role such as `write`. |

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

On a stdio host there is no inbound user token, so On-Behalf-Of is a hosted-HTTP concern; a call with no captured token fails with a clear error rather than silently falling back to an app identity. The inbound token must target the middle-tier API, and the edge must validate it and forward it unchanged. The bundled provider does not complete Conditional Access claims-challenge round trips.

`Sites.Selected` admin consent alone grants no site access. An administrator must
create an application permission assignment on every intended site and grant
the required role. For sovereign clouds, configure the Graph base URL, login
authority, OBO scope, and app-only scope as one consistent set.

## The iManage Work provider

```bash
dotnet add package OfficeAgent.Core
dotnet add package OfficeAgent.IManage
dotnet add package Microsoft.Extensions.DependencyInjection
```

`OfficeAgent.IManage` ships an `IDocumentProvider` over the iManage Work REST API v2, with no
iManage SDK. It suits a law firm or legal department whose contracts live in iManage:

```csharp
using OfficeAgent.IManage;

services.AddIManageOAuthAuthentication(o =>
{
    o.ServerUrl = "https://cloudimanage.com";
    o.ClientId = configuration["IManage:ClientId"]!;
    o.ClientSecret = configuration["IManage:ClientSecret"]!; // secret store / env, not source
    o.Username = configuration["IManage:Username"]!;         // a dedicated service account
    o.Password = configuration["IManage:Password"]!;
});
services.AddIManageDocumentProvider("matters", o =>
{
    o.ServerUrl = "https://cloudimanage.com";
    o.CustomerId = configuration["IManage:CustomerId"]!;
    o.AllowedExtensions = new[] { ".docx" };
    o.CreationFolderId = "ACTIVE!773"; // optional: where new documents are filed
});
```

The same boundary rules apply, over iManage:

- **A document is registered by its id or its link.** A source is an iManage document id,
  `LIBRARY!number.version` such as `ACTIVE!12345.1`, or an iManage Work link on the configured
  server that carries one. A registration addresses the document, not one version: opens and
  saves always use its latest version. A source naming an older version is refused, so nobody
  edits a version they did not mean to.
- **Limits come from the profile.** The size and extension in the document's iManage profile
  are checked before any content is downloaded.
- **No save overwrites.** iManage keeps every version. `Replace` adds a new version to the same
  document and keeps the id. `NewVersion` adds a new version too and mints another id.
  `NewDocument` files a new document, named `contract.v2.docx` or `NewName`, in the creation
  folder. A version keeps its document's name, so `NewName` works only with `NewDocument`.
- **Every save checks the version that was opened.** The version is the latest version number,
  its last edit time and its size. A save first compares them, then locks the document, reads
  the latest version again under the lock, uploads, and releases the lock. A new version from
  someone else, or an edit of the latest version in place, raises
  `DocumentVersionConflictException`. Set `LockDuringSave = false` only for a server that does
  not accept document locks; the version is still compared before the upload.
- **A checked-out document is refused, not written.** If someone has the document checked out
  or locked, the save fails with `WriteRejected` and says to have it checked in or to save as a
  new document instead. A refused upload is `WriteRejected`. A server error during an upload is
  `OutcomeUnknown`, because iManage may have stored it.
- **Creation has one explicit destination.** Set `CreationFolderId` to a folder or workspace id
  such as `ACTIVE!773`. iManage gives every new document its own number, so creation can never
  overwrite. A name already in that folder is refused with `AlreadyExists`. Set
  `CreationDocumentClass` if the library requires a class on new documents.
- **Listing reads the creation folder and the templates folder.** Registration has no folder
  scope, so `ListAsync` lists below `CreationFolderId`, or below `TemplatesFolder` (a folder id
  such as `ACTIVE!780`) for a templates listing. It reads
  `GET /libraries/{library}/folders/{id}/children` with `limit` and `offset`, 200 children per
  request. `Folder` is a path of folder names below that folder. Each entry's source is the
  listed version's id, `LIBRARY!number.version`, which registration accepts, so a listing shows
  the agent document ids that edits otherwise keep opaque. Search folders are not descended
  into, because their results can come from anywhere in the library. Without a creation folder
  a listing reports `ConfigurationError`. The iManage API reference is behind a customer login,
  so the `limit` and `offset` parameters are unverified against a live server; the provider
  treats a short page as the last one, and the tests run it against a fake server.
- **Tokens stay out of references and tool results.** Authentication is an
  `IIManageAccessTokenProvider`. iManage sends the token in an `X-Auth-Token` header, which
  `HttpClient` does not strip on a redirect, so the provider uses its own `HttpClient` that does
  not follow redirects. If you construct the provider yourself, pass
  `IManageDocumentProvider.CreateHttpClient()`. A download redirected to another host is
  fetched without the token.
- **Registrations are pluggable.** `InMemoryIManageRegistrationStore` is the default;
  `JsonFileIManageRegistrationStore` survives restarts; implement `IIManageRegistrationStore`
  over shared storage for multi-instance hosts.

In the MCP server, configure an iManage connection with `OfficeAgent__IManageConnections__N__*`:

```bash
export OfficeAgent__IManageConnections__0__ConnectionId=matters
export OfficeAgent__IManageConnections__0__ServerUrl=https://cloudimanage.com
export OfficeAgent__IManageConnections__0__CustomerId=<customer id>
export OfficeAgent__IManageConnections__0__ClientId=<client id>
export OfficeAgent__IManageConnections__0__ClientSecret=<from your secret store>
export OfficeAgent__IManageConnections__0__Username=svc-officeagent
export OfficeAgent__IManageConnections__0__Password=<from your secret store>
export OfficeAgent__IManageConnections__0__CreationFolderId=ACTIVE!773   # optional
export OfficeAgent__IManageConnections__0__TemplatesFolder=ACTIVE!780    # optional
export OfficeAgent__IManageConnections__0__TemplateLocation=Organization # optional
```

`GrantType` picks `Password` (the default), `RefreshToken` or `ClientCredentials`. A missing
server URL, customer id, client id, or the credentials the grant needs stops the server at
startup and names the setting. Secrets are read from configuration only and never appear in a
tool result. [MCP server settings](mcp-server.md#settings) lists every key.

### Registering the application in iManage

An iManage administrator registers OfficeAgent as an application in **Control Center** >
**Applications**, with a client id and, for a confidential client, a client secret. On
multi-tenant iManage Cloud, iManage registers the application. Then:

1. Create a dedicated service account for OfficeAgent. Give it access only to the libraries,
   workspaces and folders it should edit, with read and write rights, and import or create
   rights on the creation folder if you set one. It needs no administrator role.
2. Grant the application to that account under **Access**.
3. Read the customer id from the server's `/api` response, or ask the administrator.
4. Configure `IManageOAuthTokenProvider` with the server URL, client id, secret, and the
   account's user id and password.

| Grant | When |
| --- | --- |
| `Password` (default) | Unattended hosts. The token acts as the service account. |
| `RefreshToken` | A refresh token obtained once by signing in through the authorization-code flow. The application must have **Allow Refresh Token** set to **Yes**. |
| `ClientCredentials` | Only where your iManage environment enables this grant for the application. |

The provider requests the `user` scope, which acts with the account's own document
permissions. Never request `admin` for document editing. To use another credential flow,
implement `IIManageAccessTokenProvider` and register it directly.

iManage throttling and transient failures are not retried automatically. If a save times out,
reopen the document and compare its latest version with the intended output before retrying.
## The Google Drive provider

```bash
dotnet add package OfficeAgent.Core
dotnet add package OfficeAgent.GoogleDrive
dotnet add package Microsoft.Extensions.DependencyInjection
```

`OfficeAgent.GoogleDrive` ships an `IDocumentProvider` over the Drive API v3 REST endpoints, with
no Google client library:

```csharp
using OfficeAgent.GoogleDrive;

services.AddSingleton(new HttpClient());
services.AddGoogleDriveServiceAccountAuthentication(o =>
{
    o.KeyJson = File.ReadAllText(configuration["Drive:KeyFile"]!); // secret store / mounted file, not source
});
services.AddGoogleDriveDocumentProvider(
    connectionId: "drive",
    o =>
    {
        o.AllowedExtensions = new[] { ".docx", ".pptx", ".xlsx" };
        o.CreationFolderId = configuration["Drive:CreationFolderId"]!;
    });
```

- **Office files only.** The provider edits `.docx`, `.pptx` and `.xlsx` files stored in Drive.
  Native Google Docs, Sheets and Slides are refused by MIME type with `extension-not-allowed`
  and a message that says why. They are never converted: converting a Google Doc to Word and
  back loses comments and suggested edits. Folders, shortcuts and files in the trash are refused
  too.
- **A document is registered by a file id or a URL.** A source is a Drive file id, a Drive URL
  (`https://drive.google.com/file/d/{id}/view`, `https://drive.google.com/open?id={id}`), or the
  Docs editor URL Drive opens an Office file in (`https://docs.google.com/document/d/{id}/edit`).
  Any file the configured identity can reach is registrable. Shared drives work: every request
  sends `supportsAllDrives=true`.
- **Versions are head revision ids.** A reference's version is the file's `headRevisionId`,
  which changes with every content upload and not with a rename or a sharing change.
- **`Replace` is checked, not locked.** Drive API v3 has no server-side precondition on a
  content update: v3 files carry no ETag. `Replace` therefore reads the head revision
  immediately before uploading and raises `DocumentVersionConflictException` on a mismatch. A
  write that lands in the moment between that check and the upload is overwritten. Drive keeps
  it as an earlier revision in the file's version history, where it can be restored. A file
  locked with a read-only content restriction is refused with `write-rejected` before upload.
- **Creation never overwrites, and checks the name.** Drive allows two files with the same
  name in a folder, and an upload always makes a new file, so nothing is ever overwritten.
  `NewVersion`, `NewDocument` and creation look for the name in the destination folder first
  and refuse a taken name with `AlreadyExists`. Two writers that create the same name at the
  same moment can still both succeed, leaving two files with that name.
- **Creation has one destination.** Set `CreationFolderId` to allow new documents in that folder.
  `NewVersion` and `NewDocument` write beside the source file.
- **Listing reads the creation folder and the templates folder.** Registration has no folder
  scope, so `ListAsync` lists below `CreationFolderId`, or below `TemplatesFolder` (a Drive
  folder id) for a templates listing. It calls `files.list` with the query
  `'{folder}' in parents and trashed = false`, quotes and backslashes escaped, 200 files per
  request, on shared drives with `corpora=drive`. `Folder` is a path of folder names, each found
  with one name query. Each entry's source is the Drive file id. Native Google files, shortcuts
  and files over `MaximumBytes` are not listed. Without a creation folder a listing reports
  `ConfigurationError`.
- **Uploads are single requests.** The provider does not implement resumable uploads, which
  Google recommends for files over 5 MB. A failed single request is not retried, so keep
  `MaximumBytes` (100 MB by default) as low as your documents allow.
- **Registrations are pluggable.** `InMemoryGoogleDriveRegistrationStore` is the default;
  `JsonFileGoogleDriveRegistrationStore` survives restarts; implement
  `IGoogleDriveRegistrationStore` for shared storage.

### Google Drive authentication

`GoogleServiceAccountAccessTokenProvider` signs a JWT with the service account's RSA key and
exchanges it at `https://oauth2.googleapis.com/token` (the JWT bearer grant). Pass the key file's
JSON as `KeyJson`, or `ClientEmail` and `PrivateKey` from a secret store. The key file's own
`token_uri` is ignored, so the signed assertion only ever goes to `TokenUri`. To use another
credential flow, implement `IGoogleDriveAccessTokenProvider`.

There are two ways to give a service account access:

- **Share with the service account (narrowest).** Leave `Subject` empty. The account acts as
  itself and reaches only the files, folders and shared drives shared with its email address.
  Share the folders it should edit with it as an editor, or add it to a shared drive as a
  content manager. No administrator setup is needed beyond enabling the Drive API.
- **Domain-wide delegation.** Set `Subject` to a Workspace user's email address and the account
  acts as that user. A Workspace super administrator must authorize the service account's client
  id for the exact scope under Security > API controls > Domain-wide delegation. This reaches
  everything the user can reach, so prefer sharing where it is enough.

The scope is `GoogleDriveScopes.Drive` (`https://www.googleapis.com/auth/drive`) by default. It is
the narrowest scope that can open and save files the app did not create. `drive.file` covers
only files the app created or a user opened with it through a picker, so it suits a connection
that only edits documents it created itself, and cannot register existing files.

## The Box provider

```bash
dotnet add package OfficeAgent.Core
dotnet add package OfficeAgent.Box
dotnet add package Microsoft.Extensions.DependencyInjection
```

`OfficeAgent.Box` ships an `IDocumentProvider` over the Box REST API, with no Box SDK:

```csharp
using OfficeAgent.Box;

services.AddSingleton(new HttpClient());
services.AddBoxClientCredentialsAuthentication(o =>
{
    o.ClientId     = configuration["Box:ClientId"]!;
    o.ClientSecret = configuration["Box:ClientSecret"]!;   // secret store / env, not source
    o.SubjectType  = BoxSubjectType.Enterprise;
    o.SubjectId    = configuration["Box:EnterpriseId"]!;
});
services.AddBoxDocumentProvider(
    connectionId: "box",
    o =>
    {
        o.AllowedExtensions = new[] { ".docx", ".pptx", ".xlsx" };
        o.CreationFolderId = configuration["Box:CreationFolderId"]!;
    });
```

- **A document is registered by a file id or a URL.** A source is a numeric Box file id or a
  Box file URL (`https://app.box.com/file/{id}`, or your enterprise's `*.app.box.com` host).
  Shared links (`/s/…`) are refused; register the file itself.
- **Versions are ETags, and Box enforces them.** `Replace` uploads a new file version with
  `If-Match` on the opened ETag. Box refuses the upload with 412 when the file changed, which
  raises `DocumentVersionConflictException`, so a write between the provider's check and its
  upload is never overwritten. Box keeps the previous version in the file's history.
- **Creation never overwrites.** `NewVersion`, `NewDocument` and creation upload a new file. Box
  refuses a name already used in the folder with 409, which is `AlreadyExists`. `NewVersion`
  picks the next free `.v{n}` name with Box's upload preflight check, which uploads nothing.
- **Creation has one destination.** Set `CreationFolderId` to allow new documents in that folder.
  `NewVersion` and `NewDocument` write beside the source file.
- **Listing reads the creation folder and the templates folder.** Registration has no folder
  scope, so `ListAsync` lists below `CreationFolderId`, or below `TemplatesFolder` (a numeric
  folder id) for a templates listing. It calls `GET /folders/{id}/items` with marker paging, 200
  items per request. `Folder` is a path of folder names, found by reading each level's items.
  Each entry's source is the Box file id. Web links and files over `MaximumBytes` are not
  listed. Without a creation folder a listing reports `ConfigurationError`.
- **A locked file says so.** A file locked by someone else is refused with `write-rejected` and
  a message that asks them to unlock it or suggests saving as a new document. A Box Shield
  access policy that blocks a download or upload is `access-denied`, and the message names the
  policy as the cause.
- **Downloads keep the token at Box.** Box answers a content request with a redirect to a
  short-lived download URL. The provider follows it without the bearer token when the
  `HttpClient` does not follow redirects itself; .NET drops the `Authorization` header on
  redirects it follows.
- **Uploads are single requests.** The provider does not implement Box chunked uploads, which
  Box recommends above 50 MB. `MaximumBytes` defaults to 50 MB.
- **Registrations are pluggable.** `InMemoryBoxRegistrationStore` is the default;
  `JsonFileBoxRegistrationStore` survives restarts; implement `IBoxRegistrationStore` for shared
  storage.

### Box authentication

`BoxClientCredentialsAccessTokenProvider` uses Box's server authentication with the client
credentials grant, against `https://api.box.com/oauth2/token`. Create a Custom App with
"Server Authentication (Client Credentials Grant)" in the Box Developer Console, and have a Box
administrator authorize it in the Admin Console. Changing the app's settings needs the app
authorized again.

| Subject | `SubjectType` / `SubjectId` | App configuration |
| --- | --- | --- |
| The app's service account (narrowest) | `Enterprise` / the enterprise id | App Access Only. Collaborate the service account on each folder it should edit, as Editor. |
| An app user | `User` / the app user's id | Enable "Generate user access tokens". |
| A managed user | `User` / the user's id | App + Enterprise Access, and "Generate user access tokens". |

Under application scopes, enable "Read all files and folders stored in Box" and "Write all files
and folders stored in Box". The provider needs both and no others. Tokens last about an hour; the
provider asks for a new one when the old one is about to expire. To use another credential flow,
implement `IBoxAccessTokenProvider`.

In the MCP server, configure Google Drive with `OfficeAgent__GoogleDriveConnections__N__*` and Box
with `OfficeAgent__BoxConnections__N__*`. A Drive connection reads its service account key from
`KeyFilePath`, or from `KeyJson` supplied through the environment, and acts as itself unless
`Subject` names a Workspace user. A Box connection takes `ClientId`, `ClientSecret`, `SubjectType`
and `SubjectId`. Both create documents only when `CreationFolderId` is set, and both take
`TemplatesFolder` and `TemplateLocation` for a templates folder. A missing or unreadable key, a
missing Box credential, or a templates folder that is not a folder id stops the server at
startup. See [MCP server settings](mcp-server.md#settings).

## The Azure Blob Storage provider

```bash
dotnet add package OfficeAgent.AzureBlob
```

`OfficeAgent.AzureBlob` keeps documents in Azure Blob Storage. One connection is one container
and an optional prefix:

```csharp
using OfficeAgent.AzureBlob;

services.AddAzureBlobDocumentProvider("contracts", o =>
{
    o.ServiceUri    = "https://contoso.blob.core.windows.net";
    o.ContainerName = "legal";
    o.Prefix        = "contracts/";
    o.AllowedExtensions = new[] { ".docx" };
    o.RegistrationIndexPath = "/data/contracts-index.json"; // optional
});
```

- **Authentication.** With `ServiceUri`, the provider signs in with `DefaultAzureCredential`:
  the managed identity of an App Service, Container App or VM in Azure, and your developer
  sign-in locally. Set `ManagedIdentityClientId` for a user-assigned identity. Give the identity
  the *Storage Blob Data Contributor* role on the container, not the account. The alternative
  is `ConnectionString`, read from a secret store or the environment; do not put an account key
  in a configuration file. For any other credential, a SAS, or the Azurite emulator, pass your
  own client: `AddAzureBlobDocumentProvider("contracts", sp => new BlobContainerClient(...), o => ...)`.
- **Names stay inside the prefix.** A document is registered by its blob name relative to the
  prefix (`2026/acme.docx`), or by its URL in the same container. A leading `/`, a backslash, a
  `.` or `..` segment, a URL for another container, or a name outside the prefix is refused
  with `AccessDenied`. Stored names are checked again on every open and save.
- **Versions are ETags.** `Replace` uploads with `If-Match`, so a change since the document was
  opened raises `DocumentVersionConflictException`. The blob's content type and metadata are kept.
- **New blobs never overwrite.** `NewVersion` writes `contract.v2.docx` beside the source,
  `NewDocument` writes the `NewName` you give, and `CreateAsync` writes under the prefix. Each
  uses `If-None-Match: *`, and a taken name returns `AlreadyExists`.
- **Failed writes say what happened.** A refusal such as a lease, a 4xx status, or a missing
  container stores nothing and is reported as such. A server error, a timeout or a lost
  response is `OutcomeUnknown`. Each upload carries an `officeagent_write_id` metadata value,
  so when the SDK retries a request whose first attempt landed, the provider recognises its own
  write instead of reporting a conflict.
- **Limits come first.** The extension is checked before the service is called, and the size
  is read from the blob's properties before any download. The download is pinned to the ETag
  just checked.
- **Registrations** stay in memory unless you set `RegistrationIndexPath`. For several host
  instances, implement `IAzureBlobRegistrationStore` over shared storage and pass it to the
  `AzureBlobDocumentProvider` constructor. Removing a registration never deletes the blob.
- **Listing stays inside the prefix.** `ListAsync` calls `GetBlobsByHierarchy` with the `/`
  delimiter, 200 items per request, one prefix level at a time, so it never reads outside the
  prefix. `Folder` is a path relative to the prefix, and the same escapes are refused with
  `AccessDenied`. Each entry's source is the blob name relative to the prefix. Set
  `TemplatesFolder` to a folder inside the prefix, such as `templates`, for a templates listing.
  The identity needs list access on the container; *Storage Blob Data Contributor* includes it.

## The Amazon S3 provider

```bash
dotnet add package OfficeAgent.AmazonS3
```

`OfficeAgent.AmazonS3` keeps documents in Amazon S3. One connection is one bucket and an
optional key prefix:

```csharp
using OfficeAgent.AmazonS3;

services.AddAmazonS3DocumentProvider("contracts", o =>
{
    o.BucketName = "contoso-legal";
    o.KeyPrefix  = "contracts/";
    o.Region     = "eu-west-1";
    o.AllowedExtensions = new[] { ".docx" };
});
```

- **Authentication.** The provider uses the AWS SDK's default credential chain: the task role
  on ECS, the pod identity on EKS, or the instance role on EC2, and your profile locally. Do not
  put access keys in configuration. The role needs `s3:GetObject` and `s3:PutObject` on the
  prefix; conditional writes need both. For custom credentials, endpoint or retry settings,
  pass your own client: `AddAmazonS3DocumentProvider("contracts", sp => new AmazonS3Client(...), o => ...)`.
- **Keys stay inside the prefix.** A document is registered by its key relative to the prefix,
  or as `s3://bucket/key` for the same bucket. The same escapes as for Azure Blob are refused
  with `AccessDenied`.
- **Conditional writes.** `Replace` sends `If-Match` with the ETag that was checked, and every
  new object is written with `If-None-Match: *`. S3 answers a failed condition with
  `412 PreconditionFailed`, which is a version conflict on `Replace` and `AlreadyExists` for a
  new object. A `409 ConditionalRequestConflict`, from a concurrent request on the same key,
  stores nothing: it is a version conflict on `Replace` and `WriteRejected` for a new object.
  A 5xx answer, including `503 SlowDown`, is `OutcomeUnknown`. Writes carry an
  `officeagent-write-id` metadata value for the same retry check as Azure Blob.
- **Versioning.** On a bucket with versioning enabled, a reference's `Version` is the S3 version
  id, and an open downloads exactly that version. Otherwise `Version` is the ETag.
- **Limits, registrations and removal** work as for Azure Blob. Implement
  `IAmazonS3RegistrationStore` for shared registrations. A single PUT holds at most 5 GB, so
  keep `MaximumBytes` below that.
- **Listing stays inside the prefix.** `ListAsync` calls `ListObjectsV2` with the key prefix,
  the `/` delimiter and S3's continuation token, 200 keys per request, one prefix level at a
  time. `Folder`, sources and `TemplatesFolder` work as for Azure Blob. Listing needs
  `s3:ListBucket` on the bucket, which you can limit to the prefix with an `s3:prefix` condition.

**S3-compatible stores.** Set `ServiceUrl` (and usually `ForcePathStyle = true`). The store must
honour both conditional headers, or `Replace` and creation lose their protection.
[Cloudflare R2](https://developers.cloudflare.com/r2/api/s3/api/) supports `If-Match` and
`If-None-Match` on `PutObject`, evaluated atomically with the write; it has no bucket
versioning, so versions are ETags. MinIO supports both headers on `PutObject`; use a release
that includes the fix for conditional writes without read quorum, merged in October 2025
([minio/minio#21653](https://github.com/minio/minio/pull/21653)). Check any other store before
relying on it.

In the MCP server, configure Azure Blob Storage with `OfficeAgent__AzureBlobConnections__N__*`
(`ContainerName` and `Prefix`, with `ServiceUri` for `DefaultAzureCredential` or a
`ConnectionString` from the environment) and Amazon S3 with `OfficeAgent__AmazonS3Connections__N__*`
(`BucketName`, `KeyPrefix`, `Region`, and `ServiceUrl` with `ForcePathStyle` for an S3-compatible
store). S3 credentials come from the AWS default chain, never from OfficeAgent settings. Both
take `TemplatesFolder`, relative to the prefix, and `TemplateLocation`. See
[MCP server settings](mcp-server.md#settings).

## Listing the cloud connections

The iManage, Google Drive, Box, Azure Blob and Amazon S3 providers implement
`IDocumentListingProvider` and `IConnectionTemplateSource` with the same rules and bounds:

- **A page is bounded.** `MaximumResults` is clamped to 1 to 500. One `ListAsync` call makes at
  most 10 requests to the service, finding the query's folder included, so a page can hold
  fewer entries than asked for, even none, while a continuation token is returned. A folder
  path that would take more than that to find is refused with `InvalidArgument`.
- **Recursion is bounded.** A recursive listing reads subfolders breadth first, at most 5 levels
  below the listed folder, and holds at most 50 subfolders to read later. Further subfolders
  found while 50 are waiting are not listed. Each entry's `Folder` is its path below the
  connection's root, or the templates folder, in the form `Folder` accepts.
- **The allow-list decides.** Only extensions in `AllowedExtensions` are listed, and
  `Extensions` only narrows them. Macro-enabled formats (`.docm`, `.dotm`, `.xlsm`, `.xltm`,
  `.xlam`, `.pptm`, `.potm`, `.ppsm`, `.ppam`, `.sldm`) are never listed. Files over
  `MaximumBytes` are not listed, because registration would refuse them.
- **Templates.** A templates listing reads `TemplatesFolder` and marks every entry as a
  template; an entry elsewhere is a template when its extension is `.dotx`, `.potx` or `.xltx`.
  A connection without a templates folder returns an empty page without calling the service.
  `TemplateLocation` defaults to `Connection` when `TemplatesFolder` is set.
- **Tokens cannot leave the listing.** A continuation token is encrypted and signed with a key
  held in the process, and bound to the connection and the query. An edited token, a token from
  another query, or one issued before the server restarted is refused with `InvalidArgument`;
  list again without it. Several server instances behind a load balancer do not share tokens,
  so route a listing's pages to one instance.
- **Errors match the other calls.** A refused listing is `AccessDenied`, a missing folder,
  container or bucket is `NotFound`, and other failures are `IO`.

## Other providers

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

## Create a new document

`RegisterAsync` needs a document that already exists. `CreateAsync` starts one:

```csharp
var created = await client.CreateAsync("contracts", "brief.docx");
// created.Committed is true; created.Document is the new document's reference.

// …or with an initial plan, applied before anything is written:
var report = await client.CreateAsync("contracts", "report.docx", new DocumentPlan
{
    Operations = new PlanOperation[]
    {
        new InsertOp
        {
            Target   = new TextSpanAnchor { ParaId = "auto-0000", Expect = "" },
            Position = InsertPosition.Before,
            Text     = "Quarterly Report"
        }
    }
});
```

The Word module supplies a valid package with one empty paragraph, addressed as
`auto-0000`, over a style catalogue the `styleId` examples name: `Normal`,
`Heading1`, `Heading2`, `Heading3`, `ListParagraph`, and the `TableGrid` table
style, plus the document defaults that set the base font. Without those a
`styleId` would write a reference to a style the document does not define, and
Word would silently render the paragraph as `Normal` — the plan committing while
the document looked nothing like what was asked for. The optional plan is applied
in memory before storage is touched. Providers use their native create-if-absent operation, so an existing
name returns `AlreadyExists` and is never overwritten.

Storage creation and registration are separate steps. If storage succeeds but
registration fails, the provider leaves the new document
intact and reports the possibly unregistered filename instead of deleting remote
content or retrying an ambiguous write.

The filesystem provider implements creation and writes new documents directly
under the connection root. A SharePoint connection implements creation when its
`CreationDriveId` and `CreationFolderItemId` identify the destination folder. An iManage,
Google Drive or Box connection implements creation when its `CreationFolderId` names the
destination folder. An Azure Blob Storage or Amazon S3 connection creates documents under its
prefix, and never over an existing name.

Apart from the chunk retries above, Graph throttling and transient failures are not
retried automatically. If a write times out or receives a throttling response, reconcile storage state before
retrying: Graph may have accepted the bytes even when the response was lost.
Microsoft Graph does not support replacing a sensitivity-labeled file with
application permissions. An app-only connection cannot use `Replace` for that
case; use On-Behalf-Of/delegated access or a separately validated workflow, and
include a labeled file in preflight tests.

### Sensitivity labels and protected files

Inspection lists a document's Microsoft Purview sensitivity label as a `sensitivityLabel`
node, with the label id, its name when the file records one, and whether a person or a default
applied it. A document with no label has no node. Labels are read from the `MSIP_Label_*`
custom properties and from `docMetadata/LabelInfo.xml`, which current versions of Office write.

OfficeAgent never changes a label: no plan operation can target one, and saving leaves both
stores as they were. Labelled files are otherwise edited normally. Whether an agent may touch a
"Highly Confidential" document is the host's decision, and the node gives it what it needs to
decide.

A file a label or password has encrypted is refused with `encrypted-document` when it is
opened, before any plan runs. OfficeAgent does not decrypt, and decryption is not planned.

## List documents and templates

A provider that implements `IDocumentListingProvider` can list what a connection holds, so an
agent can find a document or a template instead of being told its source. Listing registers
nothing. Every entry's `Source` is exactly what `RegisterAsync` accepts on the same connection,
so it can be registered, opened or edited as listed:

```csharp
var page = await client.ListAsync("contracts", new DocumentListQuery
{
    Folder = "2026",              // empty for the connection root
    NameContains = "acme",        // case-insensitive
    Extensions = new[] { ".docx" },
    Recursive = true,
    MaximumResults = 100          // clamped to 1..500
});

foreach (var entry in page.Items)
    Console.WriteLine($"{entry.Folder}/{entry.Name} -> {entry.Source}");

// page.ContinuationToken fetches the next page with the same query.
var doc = await client.RegisterAsync("contracts", page.Items[0].Source);
```

`Extensions` narrows the connection's `AllowedExtensions` and never widens it. Macro-enabled
files (`.docm`, `.dotm`, `.pptm`, `.potm`, `.xlsm`, `.xltm` and the rest) are never listed,
even when a connection allows them, and neither is a file over the connection's size limit.
A continuation token is opaque and belongs to the query that produced it: reusing it with
another folder or filter is refused with `InvalidArgument`. `client.DescribeListings()` reports,
for every connection, whether it can be listed and where its templates are. A connection whose
provider does not list fails with `ConfigurationError`.

The built-in providers list as follows:

| Provider | Source | Folder | Bounds |
| --- | --- | --- | --- |
| Filesystem | The path relative to the root, with `/` separators, such as `2026/acme.docx` | A folder below the root | At most 8 folders deep, 5,000 entries examined per call, 10,000 entries read per folder |
| Session (memory) | The document's id | None; the connection has no folders | One page of the documents held |
| SharePoint | The `driveId/itemId` pair | A folder URL, a `driveId/itemId` pair, or a path such as `Contracts/2026` under the creation folder | At most 5 folders deep, 100 folders per walk, 10 Graph requests per call |

The filesystem provider never leaves its root. Links, junctions and other reparse points are
skipped rather than followed, a folder reached through one is refused with `AccessDenied`, and
an entry's folder is shown relative to the listing root, never as an absolute path. Hidden and
system files, names that start with a dot (the `.officeagent` index and staging files) and
Office lock files (`~$name.docx`) are not listed. The continuation token names the last entry
examined; it is only compared with the names found under the listed folder, so a forged token
cannot reach anything outside it.

The SharePoint provider walks a folder's children through Graph and filters names on its own
side. It does not use Graph search, whose index lags behind new files and spans the whole drive
rather than one folder. Graph's own `@odata.nextLink` travels inside the continuation token and
is followed only when it points at the configured `GraphBaseUrl`, so the bearer token never goes
anywhere else. A connection with no creation folder lists only a folder named by URL or pair.
OneDrive drive ids that do not start with `b!` must be written as `drives/{driveId}/items/{itemId}`
in `Folder`, because a bare pair without that prefix is read as a path.

### Templates folders

A connection can name a templates folder: `TemplatesFolder` on the filesystem options is a
folder relative to the root, such as `Templates`, and on the SharePoint options a folder URL or
a `driveId/itemId` pair. `DocumentListQuery.Templates` lists it instead of `Folder` (and
`Folder` is then a path inside it), and every document listed from it has `IsTemplate` set.
A `.dotx`, `.potx` or `.xltx` file is reported as a template wherever it is, but only a
connection whose `AllowedExtensions` include those extensions lists or registers them; the
default stays `.docx`.

`TemplateLocation` says what kind of location it is: `Connection` (the default once a templates
folder is set), `Personal`, `Workgroup` or `Organization`. A filesystem connection with a
location and no templates folder is a template location as a whole, which is how the MCP
server adds Office's own template folders. Set `Organization` on a SharePoint connection whose
templates folder is an organization assets library; OfficeAgent does not find that library on
its own. Both providers implement `IConnectionTemplateSource`, which reports the two values.

### Read-only connections

`FileSystemDocumentProviderOptions.ReadOnly` makes a connection refuse every save, creation and
registration removal with `AccessDenied` before anything is changed, so a tool reports
`writeOutcome: notWritten`. Listing, registration and opening still work. A read-only connection
creates no directory and keeps its registrations in memory instead of writing
`.officeagent/index.json`, so a templates folder people share is never written to.

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
| `NewVersion` | Writes the result as `{base}.v{n}.{ext}` beside the source file, mints a fresh opaque id, and registers the new path. The source is preserved. Use `NewName` to override the destination's file name. |
| `NewDocument` | Same as `NewVersion`, named for clarity when the intent is "create a sibling document". A caller-chosen `DestinationItemId` is rejected - ids are provider-assigned. |
| `Replace` | Default. Atomically overwrites the registered file in place (same id, same path) after an optimistic-concurrency check against `ExpectedVersion` (or the source reference's `Version`). |

Every save returns a canonical `DocumentReference` for the result, including its new id and version. A version mismatch raises `DocumentVersionConflictException` instead of overwriting newer content.

## Default change mode

A `changeText` operation that does not state a `mode` takes the connection's `DefaultChangeMode`. The global default is `Tracked` - a contract nobody thought to redline still arrives redlined - but a connection over generated or machine-owned documents can say otherwise once, in configuration, instead of relying on the agent to repeat it on every call:

```csharp
services.AddFileSystemDocumentProvider("reports", "/srv/officeagent/reports", o =>
{
    o.AllowedExtensions = new[] { ".docx", ".pptx" };
    o.DefaultChangeMode = ChangeMode.Direct;
});
```

This matters most for decks: PresentationML has no redline vocabulary, so under `Tracked` every deck edit that omits a mode is refused outright. A connection serving `.pptx` usually wants `Direct`.

The setting fills a gap; it never overrides an operation that names its own `mode`. It applies to plans arriving through `OfficeAgentTools` - `preview_plan`, `apply_plan`, `edit_document`, `edit_documents`, `create_document` - where an absent `mode` is genuinely absent. A host building a `DocumentPlan` in code sets `ChangeTextOp.Mode` itself, and `OfficeAgentClient.DefaultChangeModeFor(connectionId)` reports the connection's policy if it wants to follow it. A custom `IDocumentProvider` that does not implement `IConnectionEditingDefaults` takes the global default.

### Default change granularity

Beside it, `DefaultChangeGranularity` sets how much a tracked `changeText` that does not state a
`granularity` marks. `Span`, the default, marks the whole matched text as deleted and the whole
replacement as inserted. `Minimal` marks only the words that differ, so a host can have every
agent's redlines read like a reviewer's without changing any prompt:

```csharp
services.AddFileSystemDocumentProvider("contracts", "/srv/officeagent/contracts", o =>
{
    o.DefaultChangeGranularity = ChangeGranularity.Minimal;
});
```

It works like `DefaultChangeMode`: an operation that names its own `granularity` keeps it, it
applies to plans arriving through `OfficeAgentTools`, and a host building a `DocumentPlan` in code
sets `ChangeTextOp.Granularity` itself, with `OfficeAgentClient.DefaultChangeGranularityFor(connectionId)`
reporting the connection's setting. The SharePoint, iManage, Google Drive, Box, Azure Blob,
Amazon S3 and session-memory providers have the same setting; the MCP server reads it as
`DefaultChangeGranularity` on each connection. A custom provider
opts in by implementing `IConnectionGranularityDefaults`. See
[`changeText`](document-plans.md#mark-only-the-words-that-change) for what `Minimal` marks.

`Span` stays the default throughout 1.x, because the [compatibility policy](compatibility.md)
promises that a plan keeps producing the same document. `Minimal` becomes the default in 2.0.

## Formula calculation

A connection over workbooks can have every preview and commit through it calculate formulas,
so saved files carry current values. It needs a registered evaluator
(`AddExcelFormulaCalculation` from `OfficeAgent.Excel.Calculation`):

```csharp
services.AddFileSystemDocumentProvider("finance", "/srv/officeagent/finance", o =>
{
    o.AllowedExtensions = new[] { ".xlsx" };
    o.DefaultChangeMode = ChangeMode.Direct;
    o.CalculateFormulas = true;
    o.RejectFormulaErrors = true;
});
```

The same two settings exist on the SharePoint, iManage, Google Drive, Box, Azure Blob, Amazon S3
and session-memory providers. A custom
`IDocumentProvider` opts in by implementing `IConnectionCalculationDefaults`; one that does
not leaves calculation off. See [formula calculation](excel.md#formula-calculation) for what
is calculated and what is written.

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
| `ConfigurationError` | Registry has zero or more than one provider for the requested `(provider, connectionId)`, or the connection's provider cannot create documents. |
| `IO` | Underlying storage I/O error. |
| `AlreadyExists` | A document with the requested name is already there; the provider refuses to overwrite it. Retry with a different name. |

These flatten cleanly into the agent-tool surface: `OfficeAgent.AgentFramework` returns each one as a stable wire code (`not-found`, `access-denied`, `content-too-large`, `extension-not-allowed`, `version-conflict`, `invalid-argument`, `configuration-error`, `io-error`, `already-exists`) so the LLM can act on it.
