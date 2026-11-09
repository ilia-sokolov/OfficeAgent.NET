# Deployment & client setup

This is the power-user manual for connecting the OfficeAgent MCP server to the
agents that can use it: **Claude Code**, **OpenAI Codex CLI**, **Google ADK**,
**Microsoft Copilot Studio**, **Microsoft 365 Copilot**, **Microsoft Foundry Agent
Service**, and **ChatGPT**. For what the server *is* and the
full configuration reference, see [the MCP server guide](mcp-server.md); this
page is about wiring each client.

Your deployment follows from which clients you need:

| Client | Transport it speaks | Hosting you provide |
| --- | --- | --- |
| Claude Code | stdio **or** remote HTTP | None (it launches the server) - or point at a hosted URL |
| OpenAI Codex CLI | stdio **or** streamable HTTP | None - or point at a hosted URL |
| Google ADK | stdio **or** streamable HTTP | None (the agent launches the server) - or point at a hosted URL |
| Copilot Studio | streamable HTTP only | **A public HTTPS endpoint** |
| Microsoft 365 Copilot | streamable HTTP only (declarative-agent action) | **A public HTTPS endpoint** |
| Microsoft Foundry Agent Service | streamable HTTP only | **An HTTPS endpoint**, public or in your virtual network |
| ChatGPT | streamable HTTP only | **A public HTTPS endpoint** with OAuth |

If every client runs on your machine - Claude Code, Codex, or an ADK agent -
**Option A** (local, zero infra) is enough. The moment Copilot Studio, Microsoft
365 Copilot, Foundry, or ChatGPT is in scope, you need **Option B** (one hosted
server) - and the local agents can use that same hosted server too.

---

## Option A - Local (stdio), for Claude Code, Codex and Google ADK

The agent launches `officeagent-mcp` as a child process. Nothing is exposed
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

**2c. Google ADK** (Python) - install ADK with its MCP extra. Without the extra, ADK 2.x
installs no MCP client.

```bash
pip install "google-adk[mcp]"
```

```python
from google.adk.agents import LlmAgent
from google.adk.tools.mcp_tool import McpToolset
from google.adk.tools.mcp_tool.mcp_session_manager import StdioConnectionParams
from mcp import StdioServerParameters

officeagent = McpToolset(
    connection_params=StdioConnectionParams(
        server_params=StdioServerParameters(
            command="officeagent-mcp",
            args=["--stdio"],
            env={
                "OfficeAgent__FileSystemConnections__0__ConnectionId": "documents",
                "OfficeAgent__FileSystemConnections__0__RootPath": "/Users/me/Documents/agent-workspace",
            },
        ),
        timeout=30,
    ),
)

root_agent = LlmAgent(
    model="gemini-flash-latest",
    name="document_editor",
    instruction="Inspect and edit Office documents with the OfficeAgent tools.",
    tools=[officeagent],
)
```

The launched server gets `PATH` and a few other basics, plus `env`. Other variables in your
shell do not reach it, so put every `OfficeAgent__` setting in `env`. `timeout` is how long ADK
waits for each response, 5 seconds by default. Outside `adk web`, call
`await officeagent.close()` when you are done, so the server process ends.

**3. Verify.** `claude mcp list` (Claude Code) or `codex mcp list` (Codex) should
show the server connected; inside a session `/mcp` lists its tools. In ADK,
`await officeagent.get_tools()` returns them.

> SharePoint over stdio uses `appOnly` - there is no inbound user token locally, so
> the On-Behalf-Of mode is a hosted-HTTP feature (Option B).

---

## Option B - Hosted (HTTP), for every client

Deploy the server once over HTTPS; every client attaches to that URL. This is
required for Copilot Studio, Microsoft 365 Copilot, Foundry, and ChatGPT, and the
local agents can use it too.

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
platform. The [Azure template](../deploy/azure/README.md) deploys it to Container Apps with
Entra authentication in front, which covers B2 for that platform. `/healthz` is a **liveness** endpoint: it proves that the process and
configuration started, not that a provider is reachable or authorized. Add a
deployment-time provider preflight (for example, register and inspect a harmless
test document) when readiness matters.

### B2. Put authentication in front

**The standalone `officeagent-mcp` executable ships no authentication layer - do not expose it
bare.** Front it with an API gateway (Azure API Management, a reverse proxy, Front Door) that
terminates TLS and authenticates the caller. Authentication alone is insufficient: the host must
also register an `IConnectionAccessPolicy` that authorizes the trusted principal for each
connection and `read`, `register`, `create`, `edit`, or `delete` capability. The policy runs
before provider lookup, and `list_connections` omits connections the caller cannot use.

The repository's [HostedGateway reference](../samples/HostedGateway/) shows that composition with
ASP.NET Core authentication, two callers, and disjoint filesystem connections. Its bearer handler
is intentionally a local demonstration; replace it with validated JWT authentication. A hosted
product still needs tenant provisioning, quotas, encryption, secret rotation, retention controls,
regional deployment, billing, abuse controls, and operational support.

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

**Google ADK** (remote) - keep the key in an environment variable:

```python
import os
from google.adk.tools.mcp_tool import McpToolset
from google.adk.tools.mcp_tool.mcp_session_manager import StreamableHTTPConnectionParams

officeagent = McpToolset(
    connection_params=StreamableHTTPConnectionParams(
        url="https://officeagent.example.com/",
        headers={"x-api-key": os.environ["OFFICEAGENT_KEY"]},
        timeout=30,
    ),
)
```

Use the header name your gateway checks. For `onBehalfOf`, send
`Authorization: Bearer <token>` with the user's token for the middle-tier API instead. To
set it per user, pass `header_provider`, a function ADK calls before each tool call to get
extra headers.

**Copilot Studio:**

The steps below use Copilot Studio's wizard. For a definition an administrator can version and
review, import the [connector](../deploy/copilot-studio/README.md) instead.

1. Start the server with `OfficeAgent__StrictToolSchemas=false`. Copilot Studio asks the user for every required tool input the model leaves out and cannot send an empty string, so with the default strict schemas it asks for a `newName` before every in-place save and never completes one. For agents that edit documents people rely on, also set `OfficeAgent__RequirePlanSnapshot=true` so a retried call cannot apply the same plan twice. See [the MCP server settings](mcp-server.md#settings).
2. Make sure the agent uses **generative orchestration** (Settings) - MCP tools require it.
3. **Tools → Add a tool → New tool → Model Context Protocol.**
4. Enter a name, description, and the **server URL** (your gateway endpoint).
5. Pick the auth type matching your gateway: API key for `appOnly`, or OAuth 2.0 for `onBehalfOf` (dynamic client registration is supported).
6. Add it - the wizard lists the advertised tools; toggle off any you don't want (for example `register_document`).

**Microsoft 365 Copilot** (declarative agent - MCP support is GA via the
Microsoft 365 Agents Toolkit; the toolkit writes the manifests, no hand-editing):

1. In VS Code, open **Microsoft 365 Agents Toolkit → Create a new Declarative Agent.**
2. **Add an Action → Start with an MCP Server**, and give it your hosted server URL.
3. The toolkit fills in `declarativeAgent.json`, `ai-plugin.json`, and `manifest.json` (the MCP server is wrapped as an API-plugin action).
4. **Provision** and start debugging to sideload the agent into Microsoft 365 Copilot for testing; publish through the Agent Store or admin deployment.

To start from a working project instead, use the
[M365ContractReviewer sample](../samples/M365ContractReviewer/): a contract redline reviewer whose
action is the hosted server, with Entra single sign-on for On-Behalf-Of, review instructions based
on the word-document-review skill, and provisioning steps.

**Microsoft Foundry Agent Service:**

Foundry calls only remote MCP servers, so it needs this hosted server. A private endpoint in
your virtual network needs Foundry's Standard Agent Setup with private networking.

Pick the authentication that matches your SharePoint `AuthMode`:

| Foundry authentication | Use with | Notes |
| --- | --- | --- |
| Key-based (a **Custom keys** project connection) | `appOnly`, behind a gateway that checks the key | Set the credential name to the header your gateway reads, such as `x-api-key`. Anyone with access to the Foundry project can read the key. |
| Microsoft Entra: project managed identity or agent identity | `appOnly`, behind a gateway that accepts that identity's token | There is no signed-in user, so `onBehalfOf` cannot work. |
| OAuth identity passthrough, with **custom OAuth** | `onBehalfOf`, such as the [Azure template](../deploy/azure/README.md) | Each user signs in once per connection. Use your own Entra app registration. |

For OAuth identity passthrough against the Azure template, enter a client id and secret of an
app registration allowed to request `access_as_user`, the Entra v2 authorize and token URLs
for your tenant (the token URL also serves as the refresh URL), and the scopes
`api://<client-id>/access_as_user offline_access`, separated by a space. Then add the redirect
URL Foundry shows to that app registration. Users need the **Foundry Agent Consumer** role on
the project, and must be in the project's tenant.

1. Start the server with `OfficeAgent__RequirePlanSnapshot=true` for agents that edit documents people rely on, as for Copilot Studio. Leave `StrictToolSchemas` at its default. If the agent stops to ask for an input it left out, such as `newName`, set `OfficeAgent__StrictToolSchemas=false`.
2. In the Foundry portal, open **Tools**, select **Connect a tool**, then **Custom** and **MCP**.
3. Enter a name and the **server URL** (your gateway endpoint), and choose the authentication from the table.
4. Add the tool to an agent and try it in the playground.

In code, with the Azure AI Projects SDK for Python (`pip install "azure-ai-projects>=2.0.0" azure-identity`):

```python
from azure.ai.projects import AIProjectClient
from azure.ai.projects.models import MCPTool, PromptAgentDefinition
from azure.identity import DefaultAzureCredential

with AIProjectClient(
    endpoint="https://<resource>.services.ai.azure.com/api/projects/<project>",
    credential=DefaultAzureCredential(),
    allow_preview=True,
) as project:
    agent = project.agents.create_version(
        agent_name="officeagent",
        definition=PromptAgentDefinition(
            model="gpt-5-mini",  # a model deployed in your project
            instructions="Inspect and edit Office documents with the OfficeAgent tools.",
            tools=[
                MCPTool(
                    server_label="officeagent",
                    server_url="https://officeagent.example.com/",
                    require_approval="always",
                    project_connection_id="officeagent",  # the connection that holds the auth
                )
            ],
        ),
    )
```

In .NET, `ResponseTool.CreateMcpTool` in `Azure.AI.Extensions.OpenAI` builds the same tool;
set its `ProjectConnectionId`.

- `require_approval` defaults to `always`. Each call then returns an `mcp_approval_request`, and runs only after you answer with an `mcp_approval_response` in the same conversation. Foundry also takes `never`, or lists of tools under `never` or `always`, so read-only tools can run without approval.
- `allowed_tools` limits the tools the agent sees. Leaving it out, or passing an empty list, allows all of them.
- Foundry refuses a tool schema that uses `anyOf` or `allOf`, or gives a parameter several types. OfficeAgent's schemas use none of these, with either `StrictToolSchemas` setting.
- A call that is not streamed fails after 100 seconds. Split work that could take longer, such as `edit_documents` over many files.

**ChatGPT:**

Custom MCP servers are available to ChatGPT Business and Enterprise/Edu workspaces, on the web
only. Admins and owners can add them; on Enterprise/Edu, admins can let other members add them
too. Only admins and owners can publish one to the workspace.

ChatGPT authenticates with OAuth 2.1 or not at all. It cannot send an API key or a custom header,
and it does not support machine-to-machine grants. So:

- An API-key gateway does not work with ChatGPT, and the server must not run without authentication (see B2). Use OAuth.
- ChatGPT follows the MCP authorization spec. Your gateway must publish protected resource metadata at `/.well-known/oauth-protected-resource`, or name it in a `WWW-Authenticate` header on `401`. Your authorization server must support PKCE with `S256`, and one of client ID metadata documents, dynamic client registration, or a client you register for ChatGPT. ChatGPT sends `resource=<server URL>` and expects the token's audience to match. Check all of this against your gateway and identity provider before you rely on ChatGPT.
- With `onBehalfOf`, the token must be the user's Entra token for the middle-tier API, because the server exchanges it. A gateway that issues its own tokens works only with `appOnly`.
- Include `offline_access` in the scopes, or users have to sign in again when the first token expires.

1. Start the server with `OfficeAgent__RequirePlanSnapshot=true` for agents that edit documents people rely on. Leave `StrictToolSchemas` at its default. If ChatGPT stops to ask for an input the model left out, such as `newName`, set `OfficeAgent__StrictToolSchemas=false`.
2. In ChatGPT, open **Plugins**, select the plus button, then **Add custom MCP server**.
3. Enter a name and a description.
4. Under **Connection**, enter the **server URL** (your gateway endpoint). OfficeAgent serves MCP at `/`, not `/mcp`.
5. Configure OAuth, review the risk warning, and select **I understand and want to continue**.
6. Select **Create as a plugin** and review the tools it found.
7. Start a new conversation, type `@`, and pick the plugin. ChatGPT may ask the user to confirm a write; the server marks its writing tools in their MCP annotations.

After upgrading the server, open the plugin and select **Refresh** so ChatGPT reads the new tool
list. Enterprise/Edu admins can turn off tools, such as `register_document`, with
**Configure Actions** when publishing.

> Client configuration surfaces (Codex TOML keys, ADK's connection classes, the Copilot Studio
> wizard, the Agents Toolkit flow, the Foundry portal, ChatGPT's Plugins page) change quickly. If a step doesn't match, check the current
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

On an `onBehalfOf` connection, receipts name the signed-in user. Once Entra has accepted the
caller's token in the On-Behalf-Of exchange, the server takes `oid`, `iss` and `name` from it and
records them as the receipt's `actor`. Nothing is read from a header alone or from plan JSON, so
stdio, filesystem and `appOnly` calls have no actor. A document created by a plan, with no earlier
read in that request, has no actor either, because the exchange has not happened yet.

Tracked changes and comments still show `OfficeAgent`, or the plan's author, unless you set
`OfficeAgent__RevisionAuthorFromActor=true`. They then show the user's display name, and the
plan's author stays on the receipt.

`Sites.Selected` admin consent does not grant access by itself. A SharePoint or
Graph administrator must also create a site permission for the application on
each intended site and grant `write` (or another sufficient role). Verify that
assignment before starting the server.

### Pre-flight checklist for a hosted deployment

- [ ] TLS terminated and authenticated requests reach a host with `IConnectionAccessPolicy`; its policy maps trusted principals to permitted connection capabilities.
- [ ] `AuthMode` chosen deliberately; for `onBehalfOf`, the gateway forwards the caller's bearer token unchanged.
- [ ] Hosted clients that use `onBehalfOf` are configured for interactive OAuth, not a shared API key or service token.
- [ ] For `appOnly`, the app registration uses `Sites.Selected`, admin consent is complete, and the app has an explicit `write` assignment on every intended site.
- [ ] Client secret supplied from a secret store / environment, never committed.
- [ ] `RegistrationIndexPath` set if registrations must survive restarts. Protect the JSON file as sensitive application state, include it in backups, and use it from only one process. For multiple instances, implement `ISharePointRegistrationStore` over shared storage.
- [ ] `MaximumBytes` and `AllowedExtensions` reviewed against your documents (defaults: 100 MB, `.docx`). Add `.pptx` for decks and `.xlsx` for workbooks.
- [ ] Every filesystem root is owned by the service/trusted administrators. Its ACL denies untrusted principals permission to create, rename, or replace directory entries; container volumes are not shared with untrusted writers.
- [ ] `DefaultChangeMode` set per connection if `Tracked` is not what you want when a plan omits `mode` - notably `Direct` for a connection serving `.pptx`, since a deck refuses tracked changes.
- [ ] `AllowRegistration` left on only if agents should stage their own document ids; otherwise set `false`. When on, the server exposes `register_document`, `remove_document`, and the source-addressed composites `open_document`, `edit_document`, `edit_documents` and `refresh_document`, and the read-only `list_documents` and `find_templates`, which list what a connection holds. All of them reach the same documents under the same connection boundary - the composites only save round trips and the listing tools only find sources, they do not widen it.
- [ ] `AllowCreation` enabled only if agents should create files. Each SharePoint creation connection has both `CreationDriveId` and `CreationFolderItemId`; `list_connections` reports `canCreateDocuments` per connection whenever registration or creation discovery is enabled.

---

## Which recipe should I pick?

- **Only Claude Code, Codex, or an ADK agent, single developer:** Option A. No hosting, no admin.
- **Copilot Studio, Microsoft 365 Copilot, Foundry, or ChatGPT in scope, or a shared team capability:** Option B. The local agents can point at the hosted server too.
- **Both audiences:** run both - the same binary, Option A locally for the dev loop and Option B hosted for the other agents.

Note that the Microsoft-agent path always needs **one tenant-admin step** (consent for the Graph permissions, and Agent Store / admin deployment for M365 Copilot), so it is never purely self-service the way the CLI path is. ChatGPT needs a workspace admin or owner to add or publish the server.

---

## References

- [OfficeAgent MCP server guide](mcp-server.md) - transports, full configuration reference, security model.
- [SharePoint provider & authentication modes](document-providers.md) - `AuthMode`, On-Behalf-Of setup, registration stores.
- Claude Code MCP: <https://code.claude.com/docs/en/mcp>
- Codex MCP: <https://developers.openai.com/codex/mcp>
- Copilot Studio MCP: <https://learn.microsoft.com/en-us/microsoft-copilot-studio/mcp-add-existing-server-to-agent>
- Microsoft 365 Copilot declarative agents with MCP: <https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/build-mcp-plugins>
- Foundry Agent Service MCP tool: <https://learn.microsoft.com/en-us/azure/foundry/agents/how-to/tools/model-context-protocol>
- Foundry MCP server authentication: <https://learn.microsoft.com/en-us/azure/foundry/agents/how-to/mcp-authentication>
- Google ADK MCP tools: <https://adk.dev/tools-custom/mcp-tools/>
- ChatGPT custom MCP servers: <https://help.openai.com/en/articles/12584461>
- ChatGPT connection and authentication: <https://developers.openai.com/plugins/deploy/connect-chatgpt>, <https://developers.openai.com/plugins/build/auth>
