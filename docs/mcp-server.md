# MCP server

`OfficeAgent.Mcp` exposes the OfficeAgent workflow as [Model Context Protocol](https://modelcontextprotocol.io) tools, so any MCP-capable agent can inspect and edit real Word documents without taking a .NET dependency. It is the same engine and tool contract as `OfficeAgent.AgentFramework`: typed plans, preview-before-apply, tracked changes by default, and all-or-nothing commits.

> This page is the configuration reference. For step-by-step wiring of specific clients - Claude Code, Codex, Copilot Studio, Microsoft 365 Copilot - and the identity checklist, see [Deployment & client setup](deployment.md).

One binary, two transports:

| Mode | Command | Hosting |
| --- | --- | --- |
| stdio | `officeagent-mcp --stdio` | Local: the MCP client starts the server as a child process and speaks JSON-RPC over stdin/stdout. |
| streamable HTTP | `officeagent-mcp` | Cloud or shared: ASP.NET Core serves the MCP endpoint at `/` and a health probe at `/healthz`. |

## Install and run

```bash
dotnet tool install --global OfficeAgent.Mcp
officeagent-mcp --stdio
```

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
        "OfficeAgent__FileSystemConnections__0__RootPath": "/Users/me/Documents/agent-workspace"
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
| `AllowCreation` | `false` | Expose `create_document` when at least one connection allows `.docx`: any filesystem connection qualifies, while SharePoint must also have a configured creation destination. Independent of `AllowRegistration`, so a host can permit creation without permitting arbitrary registration/removal. |
| `FileSystemConnections[n]:ConnectionId` | - | Connection id agents address documents under. |
| `FileSystemConnections[n]:RootPath` | - | Root directory; registrations must stay under it, and new documents are created in it. |
| `FileSystemConnections[n]:MaximumBytes` | 100 MB | Size cap per document. |
| `FileSystemConnections[n]:AllowedExtensions` | `[".docx"]` | Extension allow-list. |
| `FileSystemConnections[n]:DefaultChangeMode` | `Tracked` | Change mode for a plan operation that does not state one. Set `Direct` for generated or machine-owned documents, or for a connection serving `.pptx` - a deck refuses `Tracked`. |
| `SharePointConnections[n]:ConnectionId` | - | Connection id agents address documents under. Documents are registered by URL or `driveId/itemId`, so the connection is not tied to one drive. |
| `SharePointConnections[n]:AuthMode` | `appOnly` | `onBehalfOf` (act as the signed-in user; hosted HTTP only) or `appOnly` (shared app identity). |
| `SharePointConnections[n]:TenantId` / `ClientId` / `ClientSecret` | - | Entra app registration. For `onBehalfOf` this is the middle-tier API app. |
| `SharePointConnections[n]:OnBehalfOfScope` | Graph `.default` | Downstream Graph scope the OBO exchange requests. |
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

The MCP toolset is the projection of [the agent-integration surface](agent-integration.md): `inspect_document`, `find_in_document`, `preview_plan`, and `apply_plan`; `AllowRegistration` independently adds `register_document` / `remove_document` plus the composites `open_document` / `edit_document`, while `AllowCreation` adds `create_document` when at least one connection allows `.docx` (SharePoint also requires its creation destination). Either opt-in adds `list_connections`, which returns `{connectionId, provider, canCreateDocuments}` entries. Tool results are the same structured JSON, including stable error codes such as `already-exists` and `configuration-error`.

The security model carries over: agents see opaque ids, never credentials; a filesystem source cannot escape its connection's root, and a SharePoint source resolves only to documents the connection's identity can already reach (per-user under On-Behalf-Of); `create_document` writes only under the filesystem root or configured SharePoint folder and never overwrites an existing name; `remove_document` drops a registration without deleting content; `apply_plan` never returns document bytes through the model - the host (or a download endpoint) retrieves the saved revision by id.

## A complete loop, from any MCP client

1. `list_connections()` → `[{ connectionId: "documents", provider: "filesystem", canCreateDocuments: true }, …]` - discover connection capabilities
2. `register_document("documents", "contract.docx")` → `{ documentId: "…" }` *(a SharePoint connection takes a document URL or `driveId/itemId` instead of a path)*
3. `find_in_document("documents", id, "Acme Corp")` → content-verified anchors
4. `preview_plan(…)` → before/after report, no write
5. `apply_plan(…)` → `{ committed: true, outputDocumentId: "…", outputName: "contract.docx", outputContentType: "…" }`
6. `remove_document("documents", id)` when the registration is no longer needed

Step 5 writes back to `contract.docx` itself (default `Replace` mode) after an optimistic version check, so the returned id is the one passed in. Pass `saveMode: "NewVersion"` to leave the source untouched and write `contract.v2.docx` beside it instead.

The composite tools collapse that loop. Steps 2–5 become one call when the edit is already known:

```jsonc
edit_document("documents", "contract.docx",
  "[ { \"op\": \"changeText\", \"target\": { \"find\": \"Acme Corp\" }, \"with\": \"Globex Inc.\" } ]")
```

→ `{ committed: true, sourceDocumentId: "…", outputDocumentId: "…", outputName: "contract.docx" }`. Text matching more than once comes back as `ambiguous-anchor` with the candidates listed, and nothing is written; re-issue with `"match": <index>`. When the document has to be read first, `open_document("documents", "contract.docx")` replaces steps 2–3 and returns the registration and the inspection together.

To start from nothing instead, replace step 2 with `create_document("documents", "brief.docx", "")` → `{ committed: true, outputDocumentId: "…" }`. The new document holds one empty paragraph addressed as `auto-0000`, so the third argument can carry an initial plan targeting it. Plan-validation errors happen before the write; provider errors may occur after storage accepted the file, so do not retry the same name blindly.
