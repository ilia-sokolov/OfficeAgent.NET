# Deployment & client setup

This is the power-user manual for connecting the OfficeAgent MCP server to the
agents that can use it: **Claude Code**, **OpenAI Codex CLI**, **Microsoft
Copilot Studio**, and **Microsoft 365 Copilot**. For what the server *is* and the
full configuration reference, see [the MCP server guide](mcp-server.md); this
page is about wiring each client.

Your deployment follows from which clients you need:

| Client | Transport it speaks | Hosting you provide |
| --- | --- | --- |
| Claude Code | stdio **or** remote HTTP | None (it launches the server) - or point at a hosted URL |
| OpenAI Codex CLI | stdio **or** streamable HTTP | None - or point at a hosted URL |
| Copilot Studio | streamable HTTP only | **A public HTTPS endpoint** |
| Microsoft 365 Copilot | streamable HTTP only (declarative-agent action) | **A public HTTPS endpoint** |

If your client list is only the two CLI agents, **Option A** (local, zero infra)
is enough. The moment Copilot Studio or Microsoft 365 Copilot is in scope, you
need **Option B** (one hosted server) - and the CLI agents can use that same
hosted server too.

---

## Option A - Local (stdio), for Claude Code and Codex

The CLI agent launches `officeagent-mcp` as a child process. Nothing is exposed
to the network; credentials stay in your local environment. ~10 minutes, no
admin.

**1. Install the tool.**

```bash
dotnet tool install --global OfficeAgent.Mcp
```

That puts `officeagent-mcp` on your PATH. (No global install? Use
`dotnet /path/to/OfficeAgent.Mcp.dll --stdio` as the command instead.)

**2a. Claude Code.**

```bash
claude mcp add \
  --env OfficeAgent__FileSystemConnections__0__ConnectionId=documents \
  --env OfficeAgent__FileSystemConnections__0__RootPath=/Users/me/Documents/agent-workspace \
  --transport stdio \
  officeagent -- officeagent-mcp --stdio
```

The `--` separates Claude Code's flags from the launched command. Project-scoped
equivalent in `.mcp.json` (checked into the repo):

```json
{
  "mcpServers": {
    "officeagent": {
      "type": "stdio",
      "command": "officeagent-mcp",
      "args": ["--stdio"],
      "env": {
        "OfficeAgent__FileSystemConnections__0__ConnectionId": "documents",
        "OfficeAgent__FileSystemConnections__0__RootPath": "${HOME}/Documents/agent-workspace"
      }
    }
  }
}
```

**2b. Codex CLI** - add to `~/.codex/config.toml` (or project-scoped
`.codex/config.toml` in a trusted project):

```toml
[mcp_servers.officeagent]
command = "officeagent-mcp"
args = ["--stdio"]
env = { "OfficeAgent__FileSystemConnections__0__ConnectionId" = "documents", "OfficeAgent__FileSystemConnections__0__RootPath" = "/Users/me/Documents/agent-workspace" }
```

**3. Verify.** `claude mcp list` (Claude Code) or `codex mcp list` (Codex) should
show the server connected; inside a session `/mcp` lists its tools.

> SharePoint over stdio uses `appOnly` - there is no inbound user token locally, so
> the On-Behalf-Of mode is a hosted-HTTP feature (Option B).

---

## Option B - Hosted (HTTP), for all four clients

Deploy the server once over HTTPS; every client attaches to that URL. This is
required for Copilot Studio and Microsoft 365 Copilot, and the CLI agents can
use it too.

### B1. Host the server

Without `--stdio`, `officeagent-mcp` is a normal ASP.NET Core app: MCP endpoint
at `/`, health probe at `/healthz`, configuration from `OfficeAgent__` environment
variables. A SharePoint connection acting as the **signed-in user** (recommended
for multi-user agents - see [identity](#authentication--identity)):

```bash
export OfficeAgent__SharePointConnections__0__ConnectionId=legal
export OfficeAgent__SharePointConnections__0__AuthMode=onBehalfOf
export OfficeAgent__SharePointConnections__0__TenantId=<tenant-id>
export OfficeAgent__SharePointConnections__0__ClientId=<middle-tier API app id>
export OfficeAgent__SharePointConnections__0__ClientSecret=<from your secret store>
export OfficeAgent__SharePointConnections__0__RegistrationIndexPath=/data/legal-index.json
export ASPNETCORE_URLS=http://0.0.0.0:8080
officeagent-mcp
```

Build and run the repository's container image for a local filesystem-backed
deployment:

```bash
docker build -t officeagent-mcp .
docker run --rm -p 8080:8080 \
  -v /absolute/path/to/documents:/data/documents \
  -e OfficeAgent__FileSystemConnections__0__ConnectionId=documents \
  -e OfficeAgent__FileSystemConnections__0__RootPath=/data/documents \
  officeagent-mcp
```

Deploy the same image to Azure Container Apps, App Service, or another container
platform. `/healthz` is a **liveness** endpoint: it proves that the process and
configuration started, not that a provider is reachable or authorized. Add a
deployment-time provider preflight (for example, register and inspect a harmless
test document) when readiness matters.

### B2. Put authentication in front

**The open-source server ships no authentication or per-caller authorization
layer - do not expose it bare.** Front it with an API gateway (Azure API
Management, a reverse proxy, Front Door) that terminates TLS, authenticates the
caller, and authorizes that caller for the requested OfficeAgent connection.
Authentication alone is insufficient: without a connection ACL, every accepted
caller can address every configured connection and registered id.

Use an API key only with `appOnly`. Use OAuth 2.0 for `onBehalfOf`, because the
server needs an API-audience user access token to exchange for Graph access.

> In On-Behalf-Of mode, validate issuer, audience, signature, expiry, and required
> delegated scope at the edge, then forward the original `Authorization` bearer
> token to OfficeAgent. The token audience must be the middle-tier API, not
> Microsoft Graph. The current token provider does not implement Conditional
> Access claims-challenge round trips; workflows that require one fail and must
> be resumed after the client obtains a suitable token.

### B3. Connect each client to the hosted URL

The static API-key/bearer examples below assume an `appOnly` SharePoint
connection. For `onBehalfOf`, configure the client and gateway for interactive
OAuth so each request carries that signed-in user's middle-tier API token.

**Claude Code** (remote):

```bash
claude mcp add \
  --header "x-api-key: ${OFFICEAGENT_KEY}" \
  --transport http \
  officeagent https://officeagent.example.com/
```

**Codex CLI** (remote) - in `~/.codex/config.toml`; keep the token in an env var,
not the file:

```toml
[mcp_servers.officeagent]
url = "https://officeagent.example.com/"
bearer_token_env_var = "OFFICEAGENT_TOKEN"
```

**Copilot Studio:**

1. Make sure the agent uses **generative orchestration** (Settings) - MCP tools require it.
2. **Tools → Add a tool → New tool → Model Context Protocol.**
3. Enter a name, description, and the **server URL** (your gateway endpoint).
4. Pick the auth type matching your gateway: API key for `appOnly`, or OAuth 2.0 for `onBehalfOf` (dynamic client registration is supported).
5. Add it - the wizard lists the advertised tools; toggle off any you don't want (for example `register_document`).

**Microsoft 365 Copilot** (declarative agent - MCP support is GA via the
Microsoft 365 Agents Toolkit; the toolkit writes the manifests, no hand-editing):

1. In VS Code, open **Microsoft 365 Agents Toolkit → Create a new Declarative Agent.**
2. **Add an Action → Start with an MCP Server**, and give it your hosted server URL.
3. The toolkit fills in `declarativeAgent.json`, `ai-plugin.json`, and `manifest.json` (the MCP server is wrapped as an API-plugin action).
4. **Provision** and start debugging to sideload the agent into Microsoft 365 Copilot for testing; publish through the Agent Store or admin deployment.

> Client configuration surfaces (Codex TOML keys, the Copilot Studio wizard, the
> Agents Toolkit flow) change quickly. If a step doesn't match, check the current
> vendor docs linked at the bottom - the server side is stable; only the client
> UIs move.

---

## Authentication & identity

Two layers, kept separate:

1. **Edge auth (who may call the server).** The gateway in B2 - API key or OAuth. Always required for a hosted deployment.
2. **SharePoint identity (whose permissions apply to documents).** Set per connection with `AuthMode`:

| `AuthMode` | Acts as | Use when |
| --- | --- | --- |
| `onBehalfOf` | The signed-in user | Multi-user hosted agents. Effective access is the intersection of the app's consented delegated Graph permissions and the user's SharePoint permissions. Requires OAuth, an inbound token for the middle-tier API (`ClientId`), and a delegated Graph scope such as `Files.ReadWrite.All`. Hosted HTTP only. |
| `appOnly` (default) | A shared app identity | Unattended agents. Every caller shares that identity, so isolate audiences with separate deployments and identities where necessary. Prefer `Sites.Selected`, then explicitly assign the app to each allowed site with the required `write` role. |

`Sites.Selected` admin consent does not grant access by itself. A SharePoint or
Graph administrator must also create a site permission for the application on
each intended site and grant `write` (or another sufficient role). Verify that
assignment before starting the server.

### Pre-flight checklist for a hosted deployment

- [ ] TLS terminated and an auth gateway in front of the server (never exposed bare); its policy maps callers to permitted connection ids.
- [ ] `AuthMode` chosen deliberately; for `onBehalfOf`, the gateway forwards the caller's bearer token unchanged.
- [ ] Hosted clients that use `onBehalfOf` are configured for interactive OAuth, not a shared API key or service token.
- [ ] For `appOnly`, the app registration uses `Sites.Selected`, admin consent is complete, and the app has an explicit `write` assignment on every intended site.
- [ ] Client secret supplied from a secret store / environment, never committed.
- [ ] `RegistrationIndexPath` set if registrations must survive restarts. Protect the JSON file as sensitive application state, include it in backups, and use it from only one process. For multiple instances, implement `ISharePointRegistrationStore` over shared storage.
- [ ] `MaximumBytes` and `AllowedExtensions` reviewed against your documents (defaults: 100 MB, `.docx`). Add `.pptx` for connections that serve decks.
- [ ] Every filesystem root is owned by the service/trusted administrators. Its ACL denies untrusted principals permission to create, rename, or replace directory entries; container volumes are not shared with untrusted writers.
- [ ] `DefaultChangeMode` set per connection if `Tracked` is not what you want when a plan omits `mode` - notably `Direct` for a connection serving `.pptx`, since a deck refuses tracked changes.
- [ ] `AllowRegistration` left on only if agents should stage their own document ids; otherwise set `false`. When on, the server exposes `register_document`, `remove_document`, and the source-addressed composites `open_document` and `edit_document`. All four reach the same documents under the same connection boundary - the composites only save round trips, they do not widen it.
- [ ] `AllowCreation` enabled only if agents should create files. Each SharePoint creation connection has both `CreationDriveId` and `CreationFolderItemId`; `list_connections` reports `canCreateDocuments` per connection whenever registration or creation discovery is enabled.

---

## Which recipe should I pick?

- **Only Claude Code and/or Codex, single developer:** Option A. No hosting, no admin.
- **Copilot Studio or Microsoft 365 Copilot in scope, or a shared team capability:** Option B. The CLI agents can point at the hosted server too.
- **Both audiences:** run both - the same binary, Option A locally for the dev loop and Option B hosted for the Microsoft agents.

Note that the Microsoft-agent path always needs **one tenant-admin step** (consent for the Graph permissions, and Agent Store / admin deployment for M365 Copilot), so it is never purely self-service the way the CLI path is.

---

## References

- [OfficeAgent MCP server guide](mcp-server.md) - transports, full configuration reference, security model.
- [SharePoint provider & authentication modes](document-providers.md) - `AuthMode`, On-Behalf-Of setup, registration stores.
- Claude Code MCP: <https://code.claude.com/docs/en/mcp>
- Codex MCP: <https://developers.openai.com/codex/mcp>
- Copilot Studio MCP: <https://learn.microsoft.com/en-us/microsoft-copilot-studio/mcp-add-existing-server-to-agent>
- Microsoft 365 Copilot declarative agents with MCP: <https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/build-mcp-plugins>
