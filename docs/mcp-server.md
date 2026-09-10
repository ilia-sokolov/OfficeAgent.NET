# MCP server

`OfficeAgent.Mcp` exposes the OfficeAgent workflow as [Model Context Protocol](https://modelcontextprotocol.io) tools, so any MCP-capable agent can inspect and edit Word documents, PowerPoint decks, and Excel workbooks without taking a .NET dependency. It is the same engine and tool contract as `OfficeAgent.AgentFramework`: typed plans, preview-before-apply, tracked changes by default in Word, and all-or-nothing commits.

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

With no configuration at all the server starts on an in-memory session connection and
says so on stderr - useful for a first look, but its documents last only while the process
runs. To edit documents on disk, create the root directory first and use an absolute path.

Or from source:

```bash
dotnet run --project src/OfficeAgent.Mcp -- --stdio
```

During initialization the server advertises its tools and the OfficeAgent prompt guidance as MCP server instructions, so a connected client passes the `(connectionId, documentId)` contract and the safety loop to its model.

## Local hosting (stdio)

Claude Desktop and many agent SDKs use an `mcpServers` object:

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
        "OfficeAgent__FileSystemConnections__0__AllowedExtensions__1": ".pptx",
        "OfficeAgent__FileSystemConnections__0__DefaultChangeMode": "Direct"
      }
    }
  }
}
```

That mixed connection defaults omitted modes to `Direct` so deck edits work. Send
`"mode": "Tracked"` explicitly for Word edits that must remain reviewable.

VS Code uses `servers` in `.vscode/mcp.json`:

```json
{
  "servers": {
    "officeagent": {
      "type": "stdio",
      "command": "officeagent-mcp",
      "args": ["--stdio"],
      "env": {
        "OfficeAgent__FileSystemConnections__0__ConnectionId": "documents",
        "OfficeAgent__FileSystemConnections__0__RootPath": "${userHome}/Documents/agent-workspace"
      }
    }
  }
}
```

See [deployment and client setup](deployment.md) for exact Claude Code and Codex locations.

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

Everything binds from the `OfficeAgent` section - `appsettings.json`, a configuration file named with `--config`, `OfficeAgent__`-prefixed environment variables, or the command line, in that order of increasing precedence.

### A configuration file

Connections do not have to be written as indexed environment variables. `--config` names a JSON file carrying the same `OfficeAgent` section, where a list of extensions is simply a list:

```bash
officeagent-mcp --stdio --config ./officeagent.json
```

```json
{
  "OfficeAgent": {
    "AllowCreation": true,
    "FileSystemConnections": [
      {
        "ConnectionId": "documents",
        "RootPath": "C:\\officeagent-documents",
        "AllowedExtensions": [ ".docx", ".pptx" ],
        "DefaultChangeMode": "Direct"
      }
    ]
  }
}
```

A copy tested through the production configuration binder lives at
[`samples/config/word-and-powerpoint.json`](../samples/config/word-and-powerpoint.json).
Change its `RootPath` before use.

With no `--config`, the server reads `OFFICEAGENT_CONFIG`, then
`%APPDATA%\OfficeAgent\config.json` (`~/.config/OfficeAgent/config.json` elsewhere).
`OfficeAgentConfiguration.ResolvePath` never discovers its additional config file from the
working directory: a stdio server starts in whatever directory its client chose. Standard
.NET configuration still loads `appsettings.json` when the host is started from a directory
that contains one. A file named explicitly and not found is an error rather than a fallback.

The file sits *below* the environment, so `OfficeAgent__` variables and the command line still override it and existing deployments are unaffected. The two merge per key rather than per connection: `OfficeAgent__FileSystemConnections__0__RootPath` re-roots the file's first connection instead of adding a second one.

Declaring `AllowedExtensions` replaces the default rather than adding to it, so a connection can be restricted to `.pptx` alone.

A connection whose value cannot be read - a `MaximumBytes` that is not a number, say - is reported by name. The configuration binder's own behaviour is to skip such an entry, which would otherwise surface much later as "no connections configured".

### Settings

| Key | Default | Meaning |
| --- | --- | --- |
| `Transport` | `http` | `http` or `stdio` (the `--stdio` flag also forces stdio). |
| `AllowRegistration` | `true` | Expose `register_document` / `remove_document` / `open_document` / `edit_document` / `list_connections` - every tool that takes a connection-relative source. Unlike the in-process tools (opt-in), the MCP server defaults to on: an MCP client has no other channel to stage document ids. Set to `false` to pin agents to ids the host distributes itself. |
| `AllowCreation` | `false` | Expose `create_document` when at least one connection allows a creatable extension (`.docx`, `.pptx`, or `.xlsx`); SharePoint must also have a configured creation destination. Independent of `AllowRegistration`, so a host can permit creation without permitting arbitrary registration/removal. The zero-configuration fallback enables it for its session connection. |
| `AllowInlineContent` | `false` | Expose `create_document_content` / `inspect_document_content` / `edit_document_content`, which carry the document as base64 in both directions. Needs no connection. Best for a single self-contained call; see [Documents with no storage](#documents-with-no-storage) before using it for multi-step editing. |
| `EphemeralConnectionId` | empty | Id of a session connection whose documents the server holds in memory for the life of the process - `"session"` by convention. Adds `import_document_content` / `export_document_content` and makes the ordinary connection-addressed tools usable with no storage. With no configuration of any kind, the effective value is `session`. |
| `EphemeralMaximumTotalBytes` | 100 MB | Total the session connection may hold at once. The store is process memory, so this bound is what keeps a long session from growing without limit. |
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

Word assembly adds the read-only `preview_document_merge` and, when `AllowCreation` is
enabled, `merge_documents`. They authorize every input connection; commit also authorizes
the destination. See [Word document assembly](document-assembly.md) for the JSON contracts.

The MCP toolset is the projection of [the agent-integration surface](agent-integration.md): `inspect_document`, `find_in_document`, `preview_plan`, `apply_plan`, and the read-only `compare_documents`; `AllowRegistration` independently adds `register_document` / `remove_document` plus the composites `open_document` / `edit_document`, while `AllowCreation` adds `create_document` and `populate_template_batch` when at least one connection allows a creatable extension - `.docx`, `.pptx`, or `.xlsx` (SharePoint also requires its creation destination). Either opt-in adds `list_connections`, which returns `{connectionId, provider, canCreateDocuments}` entries. That boolean means the connection is configured for at least one creatable format; it is not a format list, a permission check, or a readiness probe. The higher-level workflow contracts and their limits are documented in [template population and comparison](document-workflows.md).

Every tool named above addresses a document by `(connectionId, documentId)` and is offered
only when a connection exists to name. `AllowInlineContent` adds a separate set that
carries the document instead of an id - see [Documents with no storage](#documents-with-no-storage).

The schemas are strict. Every field shown in a tool signature is required on the
wire, including fields that have semantic defaults. Send `fidelity: "content"`,
`paragraphOffset: 0`, `paragraphLimit: 200`, boolean search flags as `false`,
`saveMode: "Replace"`, and `newName: ""` where those defaults are wanted. An
empty `planJson` is valid only for `create_document`; preview, apply, and edit
require an operations array or plan object.

Plan reports always contain `isValid`, `committed`, `receipt`, `sourceDocumentId`,
`outputConnectionId`, `outputDocumentId`, `outputVersion`, `outputName`,
`outputContentType`, `changes`, and `errors`. Values that do not apply are
`null`; `changes` and `errors` are arrays. Clients must decide success from
`isValid`, `committed`, and `errors`, not merely from the presence of an output
id.

On apply and inline preview, `receipt` contains the outcome, one apply timestamp, resolved
revision identity, SHA-256 hashes for the effective plan and exact document bytes, and the
saved document reference when applicable. Its `actor` is null unless the host registers an
`IAuditActorProvider`; it can never be supplied through plan JSON.

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

## Documents with no storage

### The zero-configuration default

Started with nothing configured, the server does not exit. It mints a session connection
called `session` with document creation enabled — equivalent to setting
`EphemeralConnectionId=session` and `AllowCreation=true` — and prints a line to stderr
saying so:

```
No storage is configured, so OfficeAgent started with an in-memory session connection
called 'session' and document creation enabled. Documents created there last only while
this server runs and are written nowhere. ...
```

That notice matters more than the fallback does. A typo in a connection variable produces
exactly the same state as configuring nothing, and a server that quietly ran on a session
connection while the operator believed it was pointed at their documents would be the worst
outcome — so it is always said out loud, and `list_connections` returns `session` alone.

The registration tools (`register_document`, `open_document`, `edit_document`) are not
offered in this mode: they take a path or a URL, and a session connection has neither.
Documents arrive through `import_document_content` or are made with `create_document`.

Anything the host *did* configure is left alone — the fallback applies only when nothing at
all was set, and a session connection you configured yourself still respects `AllowCreation`.

### Choosing between the two

There are two ways to run without configuring storage, and they are not interchangeable.

**A session connection holds the documents here.** `EphemeralConnectionId` adds a
connection whose documents live in the server process. The agent addresses them by opaque
id exactly as it would documents in storage — `inspect_document`, `find_in_document`,
`preview_plan`, `apply_plan`, `create_document` all work unchanged — and two extra tools
move bytes across the boundary: `import_document_content` puts a document in,
`export_document_content` takes the result out.

```json
{
  "OfficeAgent": {
    "EphemeralConnectionId": "session",
    "AllowCreation": true
  }
}
```

**Inline content carries the document in the call.** `AllowInlineContent` adds
`create_document_content`, `inspect_document_content` and `edit_document_content`, which
take the document as base64 and hand the edited document straight back, holding nothing.

Both can be on at once. Which to use is not a matter of taste:

| | Session connection | Inline content |
| --- | --- | --- |
| What the agent passes | a short opaque id | the whole document, base64 |
| Cost of the *n*th edit | unchanged | the whole file again, both ways |
| Multi-step editing | reliable | **unreliable** — see below |
| Server holds state | yes, for the session | no |
| Result reaches the host | `export_document_content` | returned by every call |

**Prefer the session connection whenever more than one edit is coming.** Chaining inline
edits requires the model to reproduce the document exactly to make the second call, and
models do not do that reliably. One exploratory run produced the following result:

| | Session connection | Inline content |
| --- | --- | --- |
| Result | succeeded | failed at step 2 |
| Wall clock | 40 s | 722 s |
| Largest tool input | 444 bytes | 3,076 bytes |
| Failed calls | 0 | 3 |

This is an observation rather than a general benchmark: the original run did not preserve
the model version and raw traces needed for reproduction. The inline run failed because the
model reproduced 2,928 characters of base64 with one
character wrong, then retried the same string twice. A document that arrives altered is
refused as `invalid-argument` saying so, rather than as an unexplained error — but it
cannot be repaired from the agent's side, because the agent's copy is the damaged one.

Inline content remains the right tool for a single self-contained call: create a document
and hand it back, or apply one known edit to content the host already holds.

Session documents are not persisted. They live as long as the server process, are written
to no storage, and are gone when it stops — export before finishing, or tell the user the
result was not saved.

### Inline content in detail

`AllowInlineContent` adds three tools that carry the document itself rather than an id. It
is one of the ways to run without storage; the other, and the better one for more than a
single call, is the session connection above - which is also what the server falls back to
when nothing at all is configured.

It is an ordinary setting, so it can be given either way — as an environment variable:

```bash
claude mcp add \
  --env OfficeAgent__AllowInlineContent=true \
  --transport stdio \
  officeagent -- officeagent-mcp --stdio
```

or in [a configuration file](#a-configuration-file):

```json
{
  "OfficeAgent": {
    "AllowInlineContent": true
  }
}
```

```bash
officeagent-mcp --stdio --config ./officeagent.json
```

The usual precedence applies: the environment overrides the file, so a deployment that ships
a file with `AllowInlineContent` on can still turn it off per host with
`OfficeAgent__AllowInlineContent=false`.

That is the whole configuration. With no connection configured the connection-addressed
tools are not offered at all - no `inspect_document`, `apply_plan`, `register_document`,
`create_document`, or `list_connections` - so the agent sees only the three tools that can
work:

| Tool | Takes | Returns |
| --- | --- | --- |
| `create_document_content` | `name`, optional `planJson` | `contentBase64` of a new document, authored by the plan |
| `inspect_document_content` | `contentBase64` | The same payload as `inspect_document` |
| `edit_document_content` | `contentBase64`, `planJson`, `preview` | `contentBase64` of the edited document |

The loop is: create or receive base64 → optionally inspect → edit → hand the returned
`contentBase64` back to the host. Nothing is stored, so the returned document is the only
copy; the next edit takes the newest base64, and passing an older one silently discards the
work in between. `contentBase64` is `null` for a preview and for a failed plan; determine
which occurred from `isValid`, `committed`, and `errors`. Targets may name text directly
(`{ "find": "Acme Corp" }`), so inspecting first is optional.

Both modes can be on at once. A host that configures connections *and* sets
`AllowInlineContent` gets both toolsets, and the agent picks by whether it holds an id or
bytes.

**What it costs.** These tools put the whole package in the model's context in both
directions: the document going in, the edited document coming back. That is tokens in
proportion to file size, and it puts the complete file - not just the text the inspect
tools already return - in front of the model provider. It is off by default for that
reason, and a deployment that configured a connection so documents would never travel that
way should leave it off. Where it earns its place is a host that has no storage to offer:
documents arriving as chat attachments, a sandbox with no writable root, or a client that
already holds the bytes.

The change mode still applies. A Word document edited inline defaults to `Tracked`, as any
connection would; a deck defaults to `Direct`, because the bytes say it is a deck and
PresentationML has no revision markup to honour `Tracked` with. An explicit
`"mode": "Tracked"` on a deck is still refused.

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
