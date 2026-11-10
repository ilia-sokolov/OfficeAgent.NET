# MCP server

`OfficeAgent.Mcp` exposes the OfficeAgent workflow as [Model Context Protocol](https://modelcontextprotocol.io) tools, so any MCP-capable agent can inspect and edit Word documents, PowerPoint decks, and Excel workbooks without taking a .NET dependency. It is the same engine and tool contract as `OfficeAgent.AgentFramework`: typed plans, preview-before-apply, tracked changes by default in Word, and all-or-nothing commits.

> This page is the configuration reference. For step-by-step wiring of specific clients - Claude Code, Codex, Google ADK, Copilot Studio, Microsoft 365 Copilot, Microsoft Foundry, ChatGPT - and the identity checklist, see [Deployment & client setup](deployment.md).

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

Or from source. Build once, then start the built server. `dotnet run` is no use to an MCP
client, because it writes build output to the stream the client reads:

```bash
dotnet build src/OfficeAgent.Mcp -c Release -f net10.0
```

```bash
dotnet src/OfficeAgent.Mcp/bin/Release/net10.0/OfficeAgent.Mcp.dll --stdio
```

No .NET installed? Use a [standalone download](#standalone-downloads), or install from PyPI:

```bash
uvx officeagent-mcp --stdio
```

The `officeagent-mcp` package carries the same self-contained server as the standalone
downloads, so it needs no .NET either. `uvx` runs it without a separate install step;
`pip install officeagent-mcp` puts an `officeagent-mcp` command on the path instead. Wheels
exist for Windows (x64, Arm64), macOS 15 or later (Intel, Apple silicon), and Linux with glibc
2.27 or later (x64, Arm64); Alpine and other musl-based systems are not covered. Because pip
and uv install the executable themselves, it carries no download mark, so Windows SmartScreen
and macOS Gatekeeper do not question it the way they do a browser download. Windows 11's Smart
App Control, on machines where it is turned on, can still block it, because the executable has
no publisher signature yet. In an MCP client's configuration, use
`"command": "uvx", "args": ["officeagent-mcp", "--stdio"]` with the same environment
variables as above.

Each wheel is built from a release's standalone archive after the archive is checked against
the release's `SHA256SUMS` and its build attestation, so it installs exactly the attested
executable. PyPI holds its own publish attestation for every wheel.

During initialization the server advertises its tools and the OfficeAgent prompt guidance as MCP server instructions, so a connected client passes the `(connectionId, documentId)` contract and the safety loop to its model.

## Standalone downloads

Every release from 1.2.0 attaches the server as one self-contained executable per platform, so it
runs where .NET is not installed. Get them from the
[latest release](https://github.com/ilia-sokolov/OfficeAgent.NET/releases/latest):

| Platform | Archive | Claude Desktop bundle |
| --- | --- | --- |
| Windows x64 | `officeagent-mcp-win-x64.zip` | `officeagent-mcp-win-x64.mcpb` |
| Windows Arm64 | `officeagent-mcp-win-arm64.zip` | `officeagent-mcp-win-arm64.mcpb` |
| macOS, Apple silicon | `officeagent-mcp-osx-arm64.tar.gz` | `officeagent-mcp-osx-arm64.mcpb` |
| macOS, Intel | `officeagent-mcp-osx-x64.tar.gz` | `officeagent-mcp-osx-x64.mcpb` |
| Linux x64 | `officeagent-mcp-linux-x64.tar.gz` | |
| Linux Arm64 | `officeagent-mcp-linux-arm64.tar.gz` | |

### Claude Desktop

Open the `.mcpb` file for your machine with Claude Desktop, or drag it onto Settings >
Extensions, then pick the folder OfficeAgent may work in. The bundle runs the server over stdio
with one filesystem connection, `documents`:

| Setting | Default | Effect |
| --- | --- | --- |
| Documents folder | Your Documents folder | The connection root. `.docx`, `.pptx`, and `.xlsx` files inside it are reachable; nothing outside it is. |
| Allow creating new documents | Off | Sets `AllowCreation`, which adds `create_document` and the template tools. |
| Calculate formulas in workbooks | Off | Sets the connection's `CalculateFormulas`, so formulas written to a workbook are saved with their results. See [formula calculation](excel.md#formula-calculation). |

Word edits are tracked changes unless the agent asks for `Direct`. PowerPoint has no tracked
changes, so the server refuses a tracked text edit to a deck and tells the agent to send it again
with `"mode": "Direct"`.

### Claude Code plugin

`plugins/officeagent` is a Claude plugin that bundles the server with the two agent skills. It
starts the server with
`uv run --project ${CLAUDE_PLUGIN_ROOT} --locked --isolated officeagent-mcp --stdio`, against
the `pyproject.toml` and `uv.lock` in the plugin folder, which pin `officeagent-mcp` to one
version and each platform's wheel to its SHA-256. [uv](https://docs.astral.sh/uv/) must be
installed. The plugin roots the server's one folder connection at `.`, the project folder
Claude Code starts it in. Creating documents and calculating formulas are plugin options that
stay off until the user turns them on. Try it with:

```bash
claude --plugin-dir ./plugins/officeagent
```

The server is a local process, so it runs in Claude Code and in Cowork sessions on the user's
computer; in claude.ai chat only the skills load. The plugin's skills are copies of `skills/`;
after editing a skill, run `python scripts/sync_claude_plugin.py` to refresh them.

### Other clients

Extract the archive and give the client the executable's absolute path where the examples on this
page say `officeagent-mcp`:

```json
{
  "mcpServers": {
    "officeagent": {
      "command": "/opt/officeagent-mcp-linux-x64/officeagent-mcp",
      "args": ["--stdio"],
      "env": {
        "OfficeAgent__FileSystemConnections__0__ConnectionId": "documents",
        "OfficeAgent__FileSystemConnections__0__RootPath": "/home/me/officeagent-documents"
      }
    }
  }
}
```

Every setting on this page applies, and started without `--stdio` the executable serves
streamable HTTP like the tool does.

### Unsigned executables

The executables carry no publisher signature. The macOS builds have the ad-hoc signature the
.NET SDK applies, which Apple silicon requires to run them at all, but not a Developer ID, so:

- On macOS, Gatekeeper refuses to start an executable extracted from an archive a browser
  downloaded. Download with `curl`, or clear the quarantine flag on the extracted file with
  `xattr -d com.apple.quarantine officeagent-mcp`.
- On Windows, SmartScreen or Smart App Control can block an executable from an archive downloaded
  from the internet. Run `Unblock-File officeagent-mcp-win-x64.zip` before extracting it, or check
  **Unblock** in the zip file's properties.

Remove the flag only from a download you have verified, as below.

### Verify a download

Each archive and bundle is covered by the release's `SHA256SUMS`, a GitHub build provenance
attestation, and a CycloneDX SBOM, `officeagent-mcp-<runtime>.<version>.cdx.json`, which lists
the packages and the .NET runtime pack compiled into the executable:

```bash
gh attestation verify officeagent-mcp-linux-x64.tar.gz --repo ilia-sokolov/OfficeAgent.NET
sha256sum --check --ignore-missing SHA256SUMS
```

### Runtime updates

A standalone executable contains the .NET runtime it was built with, and its SBOM names that
runtime's version. A .NET security update reaches it only with the next OfficeAgent release. The
.NET tool runs on whichever runtime is installed, and the container image follows its base image,
so prefer one of those where you already keep .NET up to date.

## Local hosting (stdio)

Claude Desktop can install the server as a [bundle](#claude-desktop) with no JSON to edit. Configured
by hand, Claude Desktop and many agent SDKs use an `mcpServers` object:

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
- **Registrations need durability across restarts.** Filesystem connections persist registrations in `{root}/.officeagent/index.json` automatically. SharePoint, iManage, Google Drive, Box, Azure Blob and Amazon S3 connections default to in-memory; set `RegistrationIndexPath` (single instance) or implement the provider's registration store over shared storage (multiple instances).
- **Secrets stay out of appsettings.** Supply `ClientSecret`, `Password`, `RefreshToken`, `KeyJson` and `ConnectionString` via environment variable or your secret store. They never appear in a tool result, `list_connections`, or a startup error.

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
| `AllowRegistration` | `true` | Expose `register_document` / `remove_document` / `open_document` / `edit_document` / `edit_documents` / `refresh_document` / `list_connections`, every tool that takes a connection-relative source, and `list_documents` / `find_templates`, which find one. Unlike the in-process tools (opt-in), the MCP server defaults to on: an MCP client has no other channel to stage document ids. Set to `false` to pin agents to ids the host distributes itself. |
| `AllowCreation` | `false` | Expose `create_document` when at least one connection allows a creatable extension (`.docx`, `.pptx`, or `.xlsx`); SharePoint, iManage, Google Drive and Box connections must also have a configured creation folder. Independent of `AllowRegistration`, so a host can permit creation without permitting arbitrary registration/removal. The zero-configuration fallback enables it for its session connection. |
| `AllowInlineContent` | `false` | Expose `create_document_content` / `inspect_document_content` / `edit_document_content`, which carry the document as base64 in both directions. Needs no connection. Best for a single self-contained call; see [Documents with no storage](#documents-with-no-storage) before using it for multi-step editing. |
| `EphemeralConnectionId` | empty | Id of a session connection whose documents the server holds in memory for the life of the process - `"session"` by convention. Adds `import_document_content` / `export_document_content` and makes the ordinary connection-addressed tools usable with no storage. With no configuration of any kind, the effective value is `session`. |
| `EphemeralMaximumTotalBytes` | 100 MB | Total the session connection may hold at once. The store is process memory, so this bound is what keeps a long session from growing without limit. |
| `EphemeralCalculateFormulas` / `EphemeralRejectFormulaErrors` | `false` / `false` | Calculation for the session connection, as `FileSystemConnections[n]:CalculateFormulas` and `RejectFormulaErrors` do for a folder. Applies to the zero-configuration session connection too. See [formula calculation](excel.md#formula-calculation). |
| `StrictToolSchemas` | `true` | List every tool parameter as required, including those with a default, as OpenAI strict function calling requires. Set `false` for Microsoft Copilot Studio, which asks the user for every required input the model leaves out and cannot send an empty string: under strict schemas it asks for a `newName` before every in-place save. With `false`, parameters that have a default are optional; nothing else in the schemas changes. |
| `RevisionAuthorFromActor` | `false` | Author tracked changes and comments with the signed-in user's display name on an `onBehalfOf` SharePoint connection. The plan's author stays on the receipt. Without a signed-in user nothing changes. See [identity](deployment.md#authentication--identity). |
| `RequirePlanSnapshot` | `false` | Refuse a `preview_plan`, `apply_plan` or `edit_document` plan without a non-empty `snapshot.eTag`, with `snapshot-required`, before the document is read. `edit_documents` is refused outright, because one plan cannot carry a snapshot for several documents. A plan bound to its snapshot is refused as `stale-snapshot` once the document changes, so a retried call cannot apply the same plan twice - for example, append a table row a second time. Turn it on so that holds even when the model forgets to copy the snapshot. |
| `FileSystemConnections[n]:ConnectionId` | - | Connection id agents address documents under. |
| `FileSystemConnections[n]:RootPath` | - | Root directory; registrations must stay under it, and new documents are created in it. |
| `FileSystemConnections[n]:MaximumBytes` | 100 MB | Size cap per document. |
| `FileSystemConnections[n]:AllowedExtensions` | `[".docx"]` | Extension allow-list. |
| `FileSystemConnections[n]:DefaultChangeMode` | `Tracked` | Change mode for a plan operation that does not state one. Set `Direct` for generated or machine-owned documents, or for a connection serving `.pptx` - a deck refuses `Tracked`. |
| `FileSystemConnections[n]:DefaultChangeGranularity` | `Span` | How much a tracked `changeText` that does not state a `granularity` marks: `Span` the whole matched text, `Minimal` only the words that differ. `Span` stays the server's default in 1.x and `Minimal` becomes it in 2.0. See [mark only the words that change](document-plans.md#mark-only-the-words-that-change). |
| `FileSystemConnections[n]:CalculateFormulas` | `false` | Calculate the formulas of a workbook previewed or committed through this connection, so the saved file carries current values. Only allowlisted functions are calculated; the tool result's `calculation` section names every formula left uncalculated and why. See [formula calculation](excel.md#formula-calculation). |
| `FileSystemConnections[n]:RejectFormulaErrors` | `false` | With `CalculateFormulas`, refuse a plan that introduces a formula error or a circular reference with `formula-error` instead of saving it with the errors reported. |
| `FileSystemConnections[n]:TemplatesFolder` | empty | A folder below the root that holds templates, such as `Templates`. `list_documents` with `templates: true` lists it and `find_templates` searches it. It must stay inside the root. |
| `FileSystemConnections[n]:TemplateLocation` | empty | `Connection` (the default once `TemplatesFolder` is set), `Personal`, `Workgroup` or `Organization`. Without a `TemplatesFolder`, any of these makes the whole root a template location. |
| `FileSystemConnections[n]:ReadOnly` | `false` | Refuse every save, creation and registration removal with `access-denied`, and write no `.officeagent` index under the root. Use it for a templates folder people share. A read-only connection is never offered to `create_document`. |
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
| `SharePointConnections[n]:DefaultChangeGranularity` | `Span` | As above, per library. |
| `SharePointConnections[n]:CalculateFormulas` / `RejectFormulaErrors` | `false` / `false` | As above, per library. |
| `SharePointConnections[n]:TemplatesFolder` / `TemplateLocation` | empty / empty | A templates folder, as a folder URL or `driveId/itemId` pair, and its kind, as for filesystem connections. Set `Organization` for an organization assets library; it is not found automatically. |
| `SharePointConnections[n]:ReadOnly` | `false` | Refuse every save and creation with `access-denied` before Graph is called, whatever the identity may write. Registrations still persist. Use it for a templates library people share, beside a writable connection; it cannot also name `CreationDriveId` or `CreationFolderItemId`, and is never offered to `create_document`. |
| `OfficeTemplates:Enabled` | `false` | Add the Office template folders found on this machine as read-only filesystem connections `office-personal-templates` and `office-workgroup-templates`. See [Office template locations](#office-template-locations). |
| `IManageConnections[n]:ConnectionId` | - | Connection id for an iManage Work connection. Documents are registered by id (`ACTIVE!12345.1`) or an iManage Work link on `ServerUrl`. |
| `IManageConnections[n]:ServerUrl` / `CustomerId` | - | The https iManage server, such as `https://cloudimanage.com`, and the customer id from its `/api` response. Both required. |
| `IManageConnections[n]:ClientId` / `ClientSecret` | - | The application registered in Control Center. `ClientSecret` is a secret: supply it through the environment. |
| `IManageConnections[n]:GrantType` | `Password` | `Password` (needs `Username` and `Password`), `RefreshToken` (needs `RefreshToken`) or `ClientCredentials` (needs `ClientSecret`). `Password` and `RefreshToken` are secrets. |
| `IManageConnections[n]:CreationFolderId` / `CreationDocumentClass` | empty | Folder or workspace id (`ACTIVE!773`) that receives new documents, and an optional document class. Without a folder, the connection cannot create documents. |
| `IManageConnections[n]:SaveComment` / `LockDuringSave` | `Saved by OfficeAgent.` / `true` | Comment on each saved version; lock the document while a save checks and uploads. |
| `IManageConnections[n]:ApiBasePath` / `TokenPath` | `/work/api/v2` / `/auth/oauth2/token` | Paths on `ServerUrl`, for a server that moves them. |
| `GoogleDriveConnections[n]:ConnectionId` | - | Connection id for a Google Drive connection. Documents are registered by file id or Drive URL. |
| `GoogleDriveConnections[n]:KeyFilePath` / `KeyJson` | - | The service account key: a path to the key file, or its JSON from the environment. Set exactly one. Never put the key in a manifest. |
| `GoogleDriveConnections[n]:Subject` / `Scope` | empty / Drive | A Workspace user to act as through domain-wide delegation, and the OAuth scope. |
| `GoogleDriveConnections[n]:CreationFolderId` | empty | Drive folder that receives new documents. |
| `GoogleDriveConnections[n]:ApiBaseUrl` / `UploadBaseUrl` | Drive API v3 | Override only for a proxy. |
| `BoxConnections[n]:ConnectionId` | - | Connection id for a Box connection. Documents are registered by file id or Box file URL. |
| `BoxConnections[n]:ClientId` / `ClientSecret` | - | The Box app's client credentials. `ClientSecret` is a secret. |
| `BoxConnections[n]:SubjectType` / `SubjectId` | `Enterprise` / - | Act as the app's service account (`Enterprise`, with the enterprise id) or a user (`User`, with the user id). |
| `BoxConnections[n]:CreationFolderId` | empty | Numeric Box folder id that receives new documents. |
| `BoxConnections[n]:MaximumBytes` | 50 MB | Box's ceiling for single-request uploads. |
| `BoxConnections[n]:ApiBaseUrl` / `UploadBaseUrl` | Box API 2.0 | Override only for a proxy. |
| `AzureBlobConnections[n]:ConnectionId` / `ContainerName` / `Prefix` | - / - / empty | One container and an optional prefix. Documents are registered by blob name relative to the prefix and created under it. |
| `AzureBlobConnections[n]:ServiceUri` / `ConnectionString` | - | Set exactly one: `ServiceUri` signs in with `DefaultAzureCredential` (managed identity in Azure); `ConnectionString` carries the account key and comes from the environment. |
| `AzureBlobConnections[n]:ManagedIdentityClientId` | empty | Client id of a user-assigned managed identity. |
| `AmazonS3Connections[n]:ConnectionId` / `BucketName` / `KeyPrefix` | - / - / empty | One bucket and an optional key prefix. Documents are registered by key relative to the prefix and created under it. Credentials come from the AWS default chain. |
| `AmazonS3Connections[n]:Region` | AWS default chain | Such as `eu-west-1`. |
| `AmazonS3Connections[n]:ServiceUrl` / `ForcePathStyle` | empty / `false` | An S3-compatible store that supports conditional writes, such as MinIO or Cloudflare R2. |
| `RegistrationIndexPath` | in-memory | On every iManage, Google Drive, Box, Azure Blob and Amazon S3 connection: a JSON file that makes registrations survive restarts. |
| `MaximumBytes` / `AllowedExtensions` | 100 MB / `[".docx"]` | On every iManage, Google Drive, Box, Azure Blob and Amazon S3 connection, as for filesystem connections. |
| `DefaultChangeMode` / `DefaultChangeGranularity` | `Tracked` / `Span` | On every iManage, Google Drive, Box, Azure Blob and Amazon S3 connection, as above. |
| `CalculateFormulas` / `RejectFormulaErrors` | `false` / `false` | On every iManage, Google Drive, Box, Azure Blob and Amazon S3 connection, as above. |
| `TemplatesFolder` | empty | On every iManage, Google Drive, Box, Azure Blob and Amazon S3 connection: the folder a templates listing reads. An iManage folder id (`ACTIVE!780`), a Drive folder id, a numeric Box folder id, or for Azure Blob and Amazon S3 a folder relative to the prefix (`templates`). Empty means the connection has no templates. |
| `TemplateLocation` | `Connection` when `TemplatesFolder` is set | On the same connections: where the templates folder sits, `Connection`, `Personal`, `Workgroup` or `Organization`. Setting it without `TemplatesFolder`, or `None` with one, stops the server at startup. |
| `ReadOnly` | `false` | On the same connections: refuse every save and creation with `access-denied` before the service is called. Registrations still persist. An iManage, Google Drive or Box connection that is `ReadOnly` cannot also name `CreationFolderId`; a read-only connection is never offered to `create_document`. |

A missing or malformed setting on any of these connections stops the server at startup with a
message that names the connection and the environment variable to set, such as
`OfficeAgent__IManageConnections__0__ServerUrl`. A message never quotes a secret. A
`TemplateLocation` that is not one of the four names, a filesystem `TemplatesFolder` that is not
a folder below the root, a SharePoint `TemplatesFolder` that is neither a URL nor a pair, and a
`ReadOnly` connection that names a creation folder stop it the same way. A templates folder is set like any other key, for example
`OfficeAgent__AzureBlobConnections__0__TemplatesFolder=templates` and
`OfficeAgent__AzureBlobConnections__0__TemplateLocation=Organization`.

Templates people share belong on a read-only connection of their own, beside the writable
connection agents create and edit documents on. For a SharePoint templates library:

```text
OfficeAgent__SharePointConnections__0__ConnectionId=documents
OfficeAgent__SharePointConnections__0__CreationDriveId=<drive-id>
OfficeAgent__SharePointConnections__0__CreationFolderItemId=<folder-item-id>
OfficeAgent__SharePointConnections__1__ConnectionId=templates
OfficeAgent__SharePointConnections__1__ReadOnly=true
OfficeAgent__SharePointConnections__1__TemplatesFolder=https://contoso.sharepoint.com/sites/brand/Templates
OfficeAgent__SharePointConnections__1__TemplateLocation=Organization
OfficeAgent__SharePointConnections__1__AllowedExtensions__0=.dotx
OfficeAgent__SharePointConnections__1__AllowedExtensions__1=.potx
```

`find_templates` finds a template on `templates`, and `create_document` with
`connectionId=documents`, `templateConnectionId=templates` and its `templateSource` files the
new document on `documents`. Nothing an agent does can change the template: every save to
`templates` is refused with `access-denied`. Set the credential keys on both connections as
usual.

### Office template locations

With `OfficeAgent__OfficeTemplates__Enabled=true`, the server adds the template folders Word,
PowerPoint and Excel show under File > New > Personal and Shared, as read-only filesystem
connections that `find_templates` searches:

| Connection | Where it comes from | `templateLocation` |
| --- | --- | --- |
| `office-personal-templates` | On Windows, the `PersonalTemplates` value under `HKCU\Software\Microsoft\Office\16.0\<App>\Options` for Word, PowerPoint and Excel; when none is set, `Documents\Custom Office Templates`. On macOS, `~/Library/Group Containers/UBF8T346G9.Office/User Content/Templates`. | `personal` |
| `office-workgroup-templates` | On Windows, the `SharedTemplates` value under `HKCU\Software\Policies\Microsoft\Office\16.0\Common\General`, then under `HKCU\Software\Microsoft\Office\16.0\Common\General`. | `workgroup` |

When the applications name different personal folders, each becomes its own connection,
numbered `office-personal-templates-1`, `-2` and so on. A folder that does not exist is skipped
with a line on stderr, and a connection id already configured is left as configured. These
connections allow `.docx`, `.dotx`, `.pptx`, `.potx`, `.xlsx` and `.xltx`, never a macro-enabled
file, and refuse every write: an agent can list, open and start from a template there, but
never change it or write the registration index beside it.

Not covered: Microsoft's online templates (the featured gallery in File > New), templates
managed by third-party tools that keep their own store, and the SharePoint organization assets
library, which is not discovered and must be configured as a SharePoint connection with
`TemplatesFolder` and `TemplateLocation: Organization`.

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

An iManage connection is configured the same way, with its secrets from the environment:

```json
{
  "OfficeAgent": {
    "AllowCreation": true,
    "IManageConnections": [
      {
        "ConnectionId": "matters",
        "ServerUrl": "https://cloudimanage.com",
        "CustomerId": "1",
        "ClientId": "officeagent",
        "Username": "svc-officeagent",
        "CreationFolderId": "ACTIVE!773",
        "DefaultChangeGranularity": "Minimal"
      }
    ]
  }
}
```

with `OfficeAgent__IManageConnections__0__ClientSecret` and `OfficeAgent__IManageConnections__0__Password`
supplied from the environment.


Structural validation is not native Office acceptance and neither is visual parity. See
[what "it worked" means](getting-started.md#what-it-worked-means-here) before reporting a
document as verified.

## Tools

Word assembly adds the read-only `preview_document_merge` and, when `AllowCreation` is
enabled, `merge_documents`. They authorize every input connection; commit also authorizes
the destination. See [Word document assembly](document-assembly.md) for the JSON contracts.

Every tool carries MCP annotations: a title, and explicit read-only, destructive, idempotent
and open-world hints. The read-only tools change nothing outside the call. `apply_plan`,
`edit_document`, `edit_documents` and `refresh_document` are destructive because they save over a document in place; the other writing
tools only add documents or registrations. The [wire contract](wire-contract.md#tools) lists
every tool's annotations.

The MCP toolset is the projection of [the agent-integration surface](agent-integration.md): `describe_capabilities`, `inspect_document`, `find_in_document`, `preview_plan`, `apply_plan`, `compare_documents`, and `preview_document_merge`; `AllowRegistration` independently adds `register_document` / `remove_document`, the composites `open_document` / `edit_document` / `edit_documents` / `refresh_document`, and the read-only `list_documents` / `find_templates`, while `AllowCreation` adds `create_document`, the template trio `discover_template` / `preview_template_batch` / `populate_template_batch`, and `merge_documents` when at least one connection allows a creatable extension - `.docx`, `.pptx`, or `.xlsx` (SharePoint, iManage, Google Drive and Box also require their creation folder). Either opt-in adds `list_connections`, which returns `{connectionId, provider, canCreateDocuments, allowedExtensions, defaultChangeMode, defaultChangeGranularity, calculatesFormulas, listable, templatesFolder, templateLocation, readOnly}` entries. `listable` says `list_documents` works on the connection; `templateLocation` is `organization`, `workgroup`, `personal`, `connection` or null; `readOnly` is null where configuration does not say. `provider` is `filesystem`, `sharepoint`, `imanage`, `googledrive`, `box`, `azureblob`, `amazons3` or `session`. `canCreateDocuments` means the connection is configured for at least one creatable format; it is not a permission check or a readiness probe. `register_document` and `open_document` describe only the sources of the providers that are configured. The higher-level workflow contracts and their limits are documented in [template population and comparison](document-workflows.md) and [Word document assembly](document-assembly.md).

`list_documents(connectionId, folder, nameContains, extensions, recursive, templates, maximumResults, continuation)`
lists a connection without registering anything and returns
`{connectionId, items: [{source, name, folder, contentType, length, lastModified, isTemplate}], continuation}`.
`extensions` is a comma list such as `.docx,.dotx`; `templates: true` lists the connection's
templates folder; a non-empty `continuation` fetches the next page of the same query. Each
`source` is what `open_document`, `edit_document` and `register_document` accept. A connection
whose provider cannot list returns `configuration-error`. `find_templates(format, nameContains, maximumResults)`
searches every connection with a template location the caller can read and returns
`{templates: [{connectionId, source, name, folder, format, location}], connectionsSearched, connectionsFailed: [{connectionId, code, message}]}`,
ranked by how well the name matches, then organization, workgroup, personal and connection
locations, then name. One connection that fails is reported in `connectionsFailed` and does not
fail the call. Both demand read access and are read-only. Listing bounds per provider are in
[document providers](document-providers.md#list-documents-and-templates).

Discovery and batch preflight write nothing and demand read access only, but they arrive
with `AllowCreation` rather than separately: a preflight exists to precede a commit, so
offering one where nothing can be committed would advertise a dead end. The capability each
tool demands is enforced independently of which opt-in reveals it.

Every tool named above addresses a document by `(connectionId, documentId)` and is offered
only when a connection exists to name. `AllowInlineContent` adds a separate set that
carries the document instead of an id - see [Documents with no storage](#documents-with-no-storage).

The schemas are strict. Every field shown in a tool signature is required on the
wire, including fields that have semantic defaults. Send `fidelity: "content"`,
`paragraphOffset: 0`, `paragraphLimit: 200`, `sheetId: 0`, `range: ""`,
`maximumCells: 1000`, boolean search flags as `false`, `spreadsheetValueView: "both"`,
`saveMode: "Replace"`, and `newName: ""` where those defaults are wanted. An
empty `planJson` is valid only for `create_document`; preview, apply, and edit
require an operations array or plan object.

Plan reports always contain `isValid`, `committed`, `writeOutcome`, `possibleOutput`, `receipt`, `sourceDocumentId`,
`outputConnectionId`, `outputDocumentId`, `outputVersion`, `outputName`,
`outputContentType`, `changes`, and `errors`. Values that do not apply are
`null`; `changes` and `errors` are arrays. Clients must decide success from
`isValid`, `committed`, and `errors`, not merely from the presence of an output
id. `writeOutcome` is `committed`, `notWritten`, `unknown`, or
`writtenNotRegistered`; only `notWritten` proves storage was unchanged. The uncertain
outcomes include `possibleOutput` and must be reconciled before retrying.

On apply and inline preview, `receipt` contains the outcome, one apply timestamp, resolved
revision identity, SHA-256 hashes for the effective plan and exact document bytes, and the
saved document reference when applicable. Its `actor` is null unless the host registers an
`IAuditActorProvider`; it can never be supplied through plan JSON.

Inspection returns `snapshot` as a scalar etag. To detect covered drift, copy it into the submitted plan as
`"snapshot": { "eTag": "<inspect snapshot>" }`; the server does not add it
automatically. Word covers body/header/footer/footnote/endnote XML; PowerPoint covers
slide and notes XML; Excel covers workbook/shared-string/worksheet/table/note XML. The
etag does not cover all package parts, including media bytes, so provider version checks
and operation-specific validation still apply.

The security model carries over: edit tools use opaque ids and never expose
credentials. If registration or source-addressed composite tools are enabled,
the caller can supply a path, SharePoint URL, `driveId/itemId`, iManage, Drive
or Box id or link, or blob name or object key, and that source may appear in
client/model context. A filesystem source cannot escape its connection root, a
blob name or key cannot leave its prefix, and a SharePoint, iManage, Drive or
Box source resolves only to documents the connection identity can reach (for
SharePoint, the intersection of delegated app scopes and user access under
On-Behalf-Of). `create_document` writes only under the filesystem root, the
configured folder, or the container or bucket prefix, and never overwrites a name;
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

The registration tools (`register_document`, `open_document`, `edit_document`, `edit_documents`, `refresh_document`) are not
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
the model version and raw traces needed for reproduction. It is kept because the failure mode
it shows is real, not because the numbers are measurements. For measurements that can be
reproduced, see [benchmarks](benchmarks.md); engine performance and model behaviour are
separate questions and the figures above are the latter. The inline run failed because the
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
2. `register_document({ "connectionId": "documents", "source": "contract.docx" })` → `{ documentId: "…" }` *(SharePoint accepts a document URL or `driveId/itemId`; iManage a document id or link; Drive and Box a file id or URL; Azure Blob and S3 a name relative to the prefix.)*
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

To start from the company's template, add `templateSource` (a source exactly as `register_document` takes it) and, when the template lives elsewhere, `templateConnectionId`: `create_document({ "connectionId": "documents", "name": "invoice-0042.docx", "planJson": "…", "templateConnectionId": "templates", "templateSource": "Invoice.dotx" })`. The new document is a copy of the template with the plan applied, and the template is only read. Inspect the template first for its anchors and snapshot; they hold for the copy. A `.dotx`, `.potx` or `.xltx` becomes a normal `.docx`, `.pptx` or `.xlsx`, so `name` must use that extension or the call fails with `invalid-argument` before anything is written. A macro-enabled template fails with `macro-enabled-document`. The template connection needs register and read access, and its `AllowedExtensions` must include the template's extension.

## SharePoint operational limits

A save under 4 MiB is a single Graph request; a larger one goes through a Graph
upload session in 10 MiB chunks, so `MaximumBytes` is the only size limit that
applies (Graph itself accepts files up to 250 GB). A failed chunk is retried up
to three times, honouring `Retry-After`; a single-request upload is not retried.
A save that fails before its last chunk stored nothing and reports
`write-rejected`. On `outcome-unknown`, from throttling, a timeout or a failure on
the last chunk, first reconcile by reopening the registered item (or checking the
requested new name) before retrying, because Graph may have accepted a write
whose response was lost. Reuse the same plan only after refreshing its snapshot
and version. See [document providers](document-providers.md#the-sharepoint-provider) for the
details.

Microsoft Graph does not support replacing the content of a sensitivity-labeled
file with application permissions. An `appOnly` connection therefore cannot use
`Replace` for such a file. Use an On-Behalf-Of/delegated workflow or a separately
validated process, and include a labeled document in deployment preflight tests.
