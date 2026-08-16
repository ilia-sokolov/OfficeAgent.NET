# MCP server

`OfficeAgent.Mcp` exposes the OfficeAgent workflow as [Model Context Protocol](https://modelcontextprotocol.io) tools, so any MCP-capable agent can inspect and edit real Word documents and PowerPoint decks without taking a .NET dependency. It is the same engine and tool contract as `OfficeAgent.AgentFramework`: typed plans, preview-before-apply, tracked changes by default in Word, and all-or-nothing commits.

> This page is the configuration reference. For step-by-step wiring of specific clients - Claude Code, Codex, Copilot Studio, Microsoft 365 Copilot - and the identity checklist, see [Deployment & client setup](deployment.md).

One binary, two transports:

| Mode | Command | Hosting |
| --- | --- | --- |
| stdio | `officeagent-mcp --stdio` | Local: the MCP client starts the server as a child process and speaks JSON-RPC over stdin/stdout. |
| streamable HTTP | `officeagent-mcp` | Cloud or shared: ASP.NET Core serves the MCP endpoint at `/` and a health probe at `/healthz`. |

## Install and run

```bash
dotnet tool install --global OfficeAgent.Mcp
OfficeAgent__FileSystemConnections__0__ConnectionId=documents \
OfficeAgent__FileSystemConnections__0__RootPath=/absolute/path/to/documents \
officeagent-mcp --stdio
```

PowerShell:

```powershell
$env:OfficeAgent__FileSystemConnections__0__ConnectionId = "documents"
$env:OfficeAgent__FileSystemConnections__0__RootPath = "C:\officeagent-documents"
officeagent-mcp --stdio
```

The server deliberately refuses to start with zero connections. Create the root
directory first and use an absolute path.

Or from source:

```bash
dotnet run --project src/OfficeAgent.Mcp -- --stdio
```

During initialization the server advertises its tools and the OfficeAgent prompt guidance as MCP server instructions, so a connected client passes the `(connectionId, documentId)` contract and the safety loop to its model.

## Local hosting (stdio)

A typical MCP client entry (Claude Desktop, VS Code, and most agent SDKs use this shape):

```json
{
  "mcpServers": {
    "officeagent": {
      "command": "officeagent-mcp",
      "args": ["--stdio"],
      "env": {
        "OfficeAgent__FileSystemConnections__0__ConnectionId": "documents",
        "OfficeAgent__FileSystemConnections__0__RootPath": "/Users/me/Documents/agent-workspace",
        "OfficeAgent__FileSystemConnections__0__AllowedExtensions__0": ".docx",
        "OfficeAgent__FileSystemConnections__0__AllowedExtensions__1": ".pptx"
      }
    }
  }
}
```

In stdio mode logs go to stderr; stdout carries only JSON-RPC frames.

## Cloud hosting (streamable HTTP)

The default mode is a regular ASP.NET Core app: configure it with environment variables, bind with `ASPNETCORE_URLS`, and point your platform's liveness probe at `/healthz`.

```bash
export OfficeAgent__FileSystemConnections__0__ConnectionId=documents
export OfficeAgent__FileSystemConnections__0__RootPath=/data/documents
export ASPNETCORE_URLS=http://0.0.0.0:8080
officeagent-mcp
```

Notes for production:

- **Put authentication in front.** The open-source server ships no auth layer; run it behind your reverse proxy, API gateway, or service mesh and authenticate there.
- **Registrations need durability across restarts.** Filesystem connections persist registrations in `{root}/.officeagent/index.json` automatically. SharePoint connections default to in-memory; set `RegistrationIndexPath` (single instance) or implement `ISharePointRegistrationStore` over shared storage (multiple instances).
- **Secrets stay out of appsettings.** Supply `ClientSecret` via environment variable or your secret store.

## Configuration reference

Everything binds from the `OfficeAgent` section - `appsettings.json`, `OfficeAgent__`-prefixed environment variables, or command line:

| Key | Default | Meaning |
| --- | --- | --- |
| `Transport` | `http` | `http` or `stdio` (the `--stdio` flag also forces stdio). |
| `AllowRegistration` | `true` | Expose `register_document` / `remove_document` / `open_document` / `edit_document` / `list_connections` - every tool that takes a connection-relative source. Unlike the in-process tools (opt-in), the MCP server defaults to on: an MCP client has no other channel to stage document ids. Set to `false` to pin agents to ids the host distributes itself. |
| `AllowCreation` | `false` | Expose `create_document` when at least one connection allows a creatable extension (`.docx` or `.pptx`); SharePoint must also have a configured creation destination. Independent of `AllowRegistration`, so a host can permit creation without permitting arbitrary registration/removal. |
| `FileSystemConnections[n]:ConnectionId` | - | Connection id agents address documents under. |
| `FileSystemConnections[n]:RootPath` | - | Root directory; registrations must stay under it, and new documents are created in it. |
| `FileSystemConnections[n]:MaximumBytes` | 100 MB | Size cap per document. |
| `FileSystemConnections[n]:AllowedExtensions` | `[".docx"]` | Extension allow-list. |
| `FileSystemConnections[n]:DefaultChangeMode` | `Tracked` | Change mode for a plan operation that does not state one. Set `Direct` for generated or machine-owned documents, or for a connection serving `.pptx` - a deck refuses `Tracked`. |
| `SharePointConnections[n]:ConnectionId` | - | Connection id agents address documents under. Documents are registered by URL or `driveId/itemId`, so the connection is not tied to one drive. |
| `SharePointConnections[n]:AuthMode` | `appOnly` | `onBehalfOf` (act as the signed-in user; hosted HTTP only) or `appOnly` (shared app identity). |
| `SharePointConnections[n]:TenantId` / `ClientId` / `ClientSecret` | - | Entra app registration. For `onBehalfOf` this is the middle-tier API app. |
| `SharePointConnections[n]:OnBehalfOfScope` | Graph `.default` | Downstream Graph scope the OBO exchange requests. |
| `SharePointConnections[n]:AppOnlyScope` | `https://graph.microsoft.com/.default` | Graph scope requested by `appOnly`. Override together with `GraphBaseUrl` and `LoginAuthority` for a sovereign cloud. |
| `SharePointConnections[n]:RegistrationIndexPath` | in-memory | JSON file that makes registrations survive restarts. |
| `SharePointConnections[n]:CreationDriveId` / `CreationFolderItemId` | empty | Optional Graph drive and destination-folder item ids. Set both to allow `create_document` in this connection; registration remains cross-drive. |
| `SharePointConnections[n]:GraphBaseUrl` / `LoginAuthority` | Graph v1.0 / public Entra | Override for sovereign clouds. |
| `SharePointConnections[n]:MaximumBytes` / `AllowedExtensions` | 100 MB / `[".docx"]` | Same caps as filesystem connections. |
| `SharePointConnections[n]:DefaultChangeMode` | `Tracked` | As above, per library. |

### Acting as the signed-in user (On-Behalf-Of)

With `AuthMode: onBehalfOf`, the server exchanges each caller's inbound bearer token for a Graph token that carries that user's identity, so SharePoint permissions are enforced **per user** instead of through a shared app identity. The HTTP host captures the inbound `Authorization` header automatically; the MCP client must therefore present a user token whose audience is your middle-tier API (the `ClientId`). This is the right choice for Copilot Studio and Microsoft 365 Copilot agents where many users share one hosted server. It does not apply to stdio hosting (no inbound user token).

A SharePoint connection in `appsettings.json`:

```json
{
  "OfficeAgent": {
    "SharePointConnections": [
      {
        "ConnectionId": "legal",
        "AuthMode": "onBehalfOf",
        "TenantId": "00000000-0000-0000-0000-000000000000",
        "ClientId": "00000000-0000-0000-0000-000000000000",
        "RegistrationIndexPath": "/data/officeagent/legal-index.json",
        "CreationDriveId": "b!9a3f...",
        "CreationFolderItemId": "01ABCDEF..."
      }
    ]
  }
}
```

with `OfficeAgent__SharePointConnections__0__ClientSecret` supplied from the environment.

## Tools

The MCP toolset is the projection of [the agent-integration surface](agent-integration.md): `inspect_document`, `find_in_document`, `preview_plan`, and `apply_plan`; `AllowRegistration` independently adds `register_document` / `remove_document` plus the composites `open_document` / `edit_document`, while `AllowCreation` adds `create_document` when at least one connection allows a creatable extension - `.docx` or `.pptx` (SharePoint also requires its creation destination). Either opt-in adds `list_connections`, which returns `{connectionId, provider, canCreateDocuments}` entries. That boolean means the connection is configured for at least one creatable format; it is not a format list, a permission check, or a readiness probe.

The schemas are strict. Every field shown in a tool signature is required on the
wire, including fields that have semantic defaults. Send `fidelity: "content"`,
`paragraphOffset: 0`, `paragraphLimit: 200`, boolean search flags as `false`,
`saveMode: "Replace"`, and `newName: ""` where those defaults are wanted. An
empty `planJson` is valid only for `create_document`; preview, apply, and edit
require an operations array or plan object.

Plan reports always contain `isValid`, `committed`, `sourceDocumentId`,
`outputConnectionId`, `outputDocumentId`, `outputVersion`, `outputName`,
`outputContentType`, `changes`, and `errors`. Values that do not apply are
`null`; `changes` and `errors` are arrays. Clients must decide success from
`isValid`, `committed`, and `errors`, not merely from the presence of an output
id.

Inspection returns `snapshot` as a scalar etag. To detect drift in Word text-host
XML or PowerPoint slide/notes XML, copy it into the submitted plan as
`"snapshot": { "eTag": "<inspect snapshot>" }`; the server does not add it
automatically. The etag does not cover properties, comments, sections,
media/image bytes, masters, or layouts; their own anchors and version checks
still apply.

The security model carries over: edit tools use opaque ids and never expose
credentials. If registration or source-addressed composite tools are enabled,
the caller can supply a path, SharePoint URL, or `driveId/itemId`, and that
source may appear in client/model context. A filesystem source cannot escape
its connection root, and a SharePoint source resolves only to documents the
connection identity can reach (the intersection of delegated app scopes and
user access under On-Behalf-Of). `create_document` writes only under the
filesystem root or configured SharePoint folder and never overwrites a name;
`remove_document` drops a registration without deleting content.

## A complete loop, from any MCP client

1. `list_connections({})` → `[{ connectionId: "documents", provider: "filesystem", canCreateDocuments: true }, …]`
2. `register_document({ "connectionId": "documents", "source": "contract.docx" })` → `{ documentId: "…" }` *(SharePoint accepts a document URL or `driveId/itemId`.)*
3. `find_in_document({ "connectionId": "documents", "documentId": "…", "pattern": "Acme Corp", "regex": false, "wholeWord": false, "caseSensitive": false })` → content-verified anchors
4. `preview_plan({ "connectionId": "documents", "documentId": "…", "planJson": "[...]" })` → before/after report, no write
5. `apply_plan({ "connectionId": "documents", "documentId": "…", "planJson": "[...]", "saveMode": "Replace", "newName": "" })` → `{ committed: true, outputDocumentId: "…", outputName: "contract.docx", … }`
6. `remove_document({ "connectionId": "documents", "documentId": "…" })` when the registration is no longer needed

Step 5 writes back to `contract.docx` itself (default `Replace` mode) after an optimistic version check, so the returned id is the one passed in. Pass `saveMode: "NewVersion"` to leave the source untouched and write `contract.v2.docx` beside it instead.

The composite tools collapse that loop. Steps 2–5 become one call when the edit is already known:

```jsonc
edit_document({
  "connectionId": "documents",
  "source": "contract.docx",
  "planJson": "[ { \"op\": \"changeText\", \"target\": { \"find\": \"Acme Corp\" }, \"with\": \"Globex Inc.\" } ]",
  "saveMode": "Replace",
  "newName": ""
})
```

→ `{ committed: true, sourceDocumentId: "…", outputDocumentId: "…", outputName: "contract.docx" }`. Text matching more than once comes back as `ambiguous-anchor` with the candidates listed, and nothing is written; re-issue with `"match": <index>`. When the document has to be read first, `open_document({ "connectionId": "documents", "source": "contract.docx", "fidelity": "content", "paragraphOffset": 0, "paragraphLimit": 200 })` replaces steps 2–3 and returns the registration and inspection together.

To start from nothing instead, replace step 2 with `create_document({ "connectionId": "documents", "name": "brief.docx", "planJson": "" })` → `{ committed: true, outputDocumentId: "…" }`. The new document holds one empty paragraph addressed as `auto-0000`, so `planJson` can carry an initial plan targeting it. Plan-validation errors happen before the write; provider errors may occur after storage accepted the file, so do not retry the same name blindly.

## SharePoint operational limits

The provider uses Microsoft Graph's single-request content upload API, whose
limit is 250 MB; `MaximumBytes` must not exceed that value for SharePoint.
Uploads do not currently perform automatic `Retry-After` handling. On throttling
or a timeout, first reconcile by reopening the registered item (or checking the
requested new name) before retrying, because Graph may have accepted a write
whose response was lost. Reuse the same plan only after refreshing its snapshot
and version.

Microsoft Graph does not support replacing the content of a sensitivity-labeled
file with application permissions. An `appOnly` connection therefore cannot use
`Replace` for such a file. Use an On-Behalf-Of/delegated workflow or a separately
validated process, and include a labeled document in deployment preflight tests.
