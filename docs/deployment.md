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
dotnet tool install --global OfficeAgent.Mcp --prerelease
```

That puts `officeagent-mcp` on your PATH. (No global install? Use
`dotnet /path/to/OfficeAgent.Mcp.dll --stdio` as the command instead.)

**2a. Claude Code.**

```bash
claude mcp add officeagent \
  --env OfficeAgent__FileSystemConnections__0__ConnectionId=documents \
  --env OfficeAgent__FileSystemConnections__0__RootPath=/Users/me/Documents/agent-workspace \
  -- officeagent-mcp --stdio
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

Deploy as a container or app service (Azure Container Apps, App Service, or
anywhere). Point your platform's liveness probe at `/healthz`.

### B2. Put authentication in front

**The open-source server ships no auth layer - do not expose it bare.** Front it
with an API gateway (Azure API Management, a reverse proxy, Front Door) that
terminates TLS and requires an API key or OAuth. Copilot Studio's MCP wizard can
send no-auth, API key (header or query), or OAuth 2.0, so the gateway pattern
maps directly onto every client below.

> If you use the On-Behalf-Of identity mode, the gateway must **forward the
> caller's `Authorization` bearer token through to the server** (don't strip or
> replace it) - that user token is what the server exchanges for a per-user Graph
> token.

### B3. Connect each client to the hosted URL

**Claude Code** (remote):

```bash
claude mcp add --transport http officeagent https://officeagent.example.com/ \
  --header "x-api-key: ${OFFICEAGENT_KEY}"
```

**Codex CLI** (remote) - in `~/.codex/config.toml`; keep the token in an env var,
not the file:

```toml
[mcp_servers.officeagent]
transport = { type = "streamable_http", url = "https://officeagent.example.com/" }
bearer_token_env_var = "OFFICEAGENT_TOKEN"
```

**Copilot Studio:**

1. Make sure the agent uses **generative orchestration** (Settings) - MCP tools require it.
2. **Tools → Add a tool → Model Context Protocol.**
3. Enter a name, description, and the **server URL** (your gateway endpoint).
4. Pick the auth type matching your gateway (API key in a header is simplest; OAuth 2.0 is supported, including dynamic client registration).
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
| `onBehalfOf` | The signed-in user | Multi-user hosted agents (Copilot Studio, M365 Copilot). Each user is limited to what they can already open in SharePoint. Requires the inbound user token's audience to be your middle-tier API (`ClientId`), the user's consent to it, and a matching delegated Graph permission (e.g. `Files.ReadWrite.All`). Hosted HTTP only. |
| `appOnly` (default) | A shared app identity | Unattended / back-office agents with no user present. **Scope the app registration narrowly with `Sites.Selected`** so it can reach only the intended sites. |

### Pre-flight checklist for a hosted deployment

- [ ] TLS terminated and an auth gateway in front of the server (never exposed bare).
- [ ] `AuthMode` chosen deliberately; for `onBehalfOf`, the gateway forwards the caller's bearer token unchanged.
- [ ] For `appOnly`, the app registration uses `Sites.Selected`, not tenant-wide Graph permissions.
- [ ] Client secret supplied from a secret store / environment, never committed.
- [ ] `RegistrationIndexPath` set if you need registrations to survive restarts (filesystem connections persist automatically; SharePoint defaults to in-memory). For multiple instances, implement `ISharePointRegistrationStore` over shared storage.
- [ ] `MaximumBytes` and `AllowedExtensions` reviewed against your documents (defaults: 100 MB, `.docx`).
- [ ] `AllowRegistration` left on only if agents should stage their own document ids; otherwise set `false`. When on, the server exposes `register_document`, `remove_document`, and `list_connections` (the last returns the configured `{connectionId, provider}` pairs so agents can discover where to register documents).

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
