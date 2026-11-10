# M365ContractReviewer - a Microsoft 365 Copilot contract redline reviewer

A declarative agent for Microsoft 365 Copilot, in Microsoft 365 Agents Toolkit layout, whose one
action is a hosted OfficeAgent MCP server. Given a link to a Word contract in SharePoint or
OneDrive, it reads the contract, proposes redlines as tracked changes, comments on clauses that
need a human decision, and saves the contract in place after you confirm.

The agent's instructions, in [`appPackage/instruction.txt`](appPackage/instruction.txt), are
based on the [word-document-review skill](../../skills/word-document-review/SKILL.md): read the
open comments and pending revisions first, anchor every edit on text the server found, put the
review in one plan, preview it, mark every edit `Tracked`, and recover from each refusal code in
the way the skill describes.

| File | Purpose |
| --- | --- |
| `appPackage/manifest.json` | Microsoft 365 app manifest, schema 1.30 |
| `appPackage/declarativeAgent.json` | The agent: name, conversation starters, instructions, action. Schema 1.9 |
| `appPackage/ai-plugin.json` | The action: a `RemoteMCPServer` runtime at `${{OFFICEAGENT_MCP_URL}}`, with tools discovered from the server at runtime. Plugin schema 2.4 |
| `appPackage/instruction.txt` | The review instructions |
| `m365agents.yml` | Provisioning: creates the app, registers Microsoft Entra single sign-on for the server, builds, validates and sideloads the package |
| `env/.env.dev` | Configuration for the `dev` environment. The server URL and app ids go here; nothing is committed filled in |

## Prerequisites

- **A hosted OfficeAgent server over HTTPS,** with a SharePoint connection that acts as the
  signed-in user (`AuthMode: onBehalfOf`). The [Azure deployment](../../deploy/azure/README.md)
  creates one, or see [Option B in deployment](../../docs/deployment.md#option-b---hosted-http-for-every-client).
  The agent uses `list_connections` and `open_document`, so leave `AllowRegistration` on (the
  default). `RequirePlanSnapshot=true`, which the Azure template sets, is recommended.
- **The middle-tier API app registration** the server's SharePoint connection uses as
  `ClientId`, set up as in [Before you deploy](../../deploy/azure/README.md#before-you-deploy):
  an Application ID URI, an `access_as_user` scope, v2 access tokens, and delegated Microsoft
  Graph permissions such as `Files.ReadWrite.All`.
- **Tenant admin consent** for those Graph permissions. On-Behalf-Of cannot get a Graph token
  for a user until the permissions are consented.
- **A Microsoft 365 tenant with Microsoft 365 Copilot,** where your account shows **Custom App
  Upload Enabled** and **Copilot Access Enabled** in the Agents Toolkit **Accounts** pane.
- **Visual Studio Code** with the
  [Microsoft 365 Agents Toolkit](https://aka.ms/M365AgentsToolkit) 6.12.0 or later.
- Rights to edit the app registration (redirect URI, Application ID URI, authorized client), or
  someone who has them.

## How sign-in works

Microsoft 365 Copilot signs the user in with Microsoft Entra single sign-on and calls the server
with an access token for the middle-tier API. The server exchanges that token for a Microsoft
Graph token for the same user (On-Behalf-Of), so the agent can open only the documents the user
can open, and SharePoint's permissions apply. Receipts name the user; tracked changes and
comments show `OfficeAgent` unless the server sets `RevisionAuthorFromActor=true`, which puts
the user's name on them. See [identity](../../docs/deployment.md#authentication--identity).

## Provision and sideload

1. **Configure.** In `env/.env.dev`, set:
   - `OFFICEAGENT_MCP_URL` to the server's MCP endpoint, such as `https://<your-server>/`;
   - `OFFICEAGENT_API_CLIENT_ID` to the middle-tier API app's client id.

   Keep the filled-in file out of version control.
2. **Open** this folder in Visual Studio Code. In the Agents Toolkit **Accounts** pane, sign in
   to Microsoft 365.
3. **Provision.** In the **Lifecycle** pane, select **Provision**. The toolkit creates the app,
   registers Entra single sign-on for the server (`oauth/register`), and writes
   `MCP_DA_AUTH_ID_OFFICEAGENT` and `OFFICEAGENT_SSO_APPLICATION_ID_URI` into `env/.env.dev`. It
   then builds the package with those values and sideloads the agent for your account.
4. **Update the app registration** in the
   [Microsoft Entra admin center](https://entra.microsoft.com/), as
   [Microsoft's Entra SSO guide](https://learn.microsoft.com/microsoft-365/copilot/extensibility/plugin-authentication-entra-sso)
   describes:
   - add the `OFFICEAGENT_SSO_APPLICATION_ID_URI` value to `identifierUris` in the app's
     manifest, beside the existing `api://<client-id>`;
   - under **Authentication**, add the **Web** redirect URI
     `https://teams.microsoft.com/api/platform/v1.0/oAuthConsentRedirect`;
   - under **Expose an API**, add the client application
     `ab3be6b7-f5df-413d-ac2d-abf1e3fd9c0b` (the Microsoft Enterprise token store) for the
     `access_as_user` scope.
5. **Accept the new audience** at the server's front door. Tokens now arrive with the new
   Application ID URI as their audience. On the Azure deployment, add it to the Container App's
   allowed token audiences (`allowedAudiences` in `deploy/azure/main.bicep`, or **Authentication**
   in the portal). Behind another gateway, add it to the audiences that gateway validates.
6. **Use it.** Go to [Microsoft 365 Copilot](https://m365.cloud.microsoft/chat), open
   **Agents**, and select **Contract Redline Reviewerdev**. The first call asks you to sign in.

To make the agent available to others, select **Publish** in the **Lifecycle** pane; a tenant
admin then approves it in the Microsoft 365 admin center.

To change the instructions or starters later, edit `appPackage/` and select **Provision**
again.

## Try it

Upload [`services-agreement.docx`](../documents/services-agreement.docx) to a SharePoint
document library you can edit, copy its link, and ask:

> Review the contract at `<link>`. Change the initial term in clause 5 from twelve months to
> twenty-four months as a tracked change, and comment on anything in clause 4 that legal should
> decide.

The agent should open the contract and report what is already pending: Priya Raman's open
comment on the clause 4 liability cap, and the tracked change of the interest rate in clause 3
from 2% to 3%. It should then list its proposal, for example:

- clause 5: `twelve months` -> `twenty-four months`, tracked;
- clause 4: a comment that the liability cap needs legal approval before signature;

and ask you to confirm. After you confirm, it applies the plan and saves the contract in place.

Open the contract in Word and open **Review > Reviewing Pane**. You should see:

- in clause 5, `twelve` struck through and `twenty-four` inserted before `months`, as a
  revision you can accept or reject. The instructions ask for `"granularity": "Minimal"`, so
  only the changed word is marked;
- a new comment on clause 4;
- Priya's comment and the clause 3 revision unchanged, and clause 4's own text, which also says
  twelve months, unchanged.

SharePoint's version history keeps the version from before the edit.

## Troubleshooting

| Symptom | Likely cause |
| --- | --- |
| Provisioning stops at `oauth/register` | `OFFICEAGENT_MCP_URL` or `OFFICEAGENT_API_CLIENT_ID` is empty in `env/.env.dev`. |
| The agent cannot sign in, or every call returns `401` | Step 4 or 5 was skipped: the new Application ID URI is not on the app registration or not an allowed audience. |
| Sign-in works, but the server reports `access-denied` | The user cannot open that file, or the Graph permissions are not consented. |
| `write-rejected`, with a message that the document is locked | The file is checked out or open for exclusive editing. Check it in or close it and try again, or save as a new version. |
| `version-conflict` | Someone saved the file after the agent opened it. Ask the agent to open it again and rebuild its proposal. |
| `encrypted-document` | The file is protected by a sensitivity label or a password. OfficeAgent does not decrypt; use an unprotected copy. |
| The agent edits without `Tracked` | It should not; the instructions require it. Report the prompt and the reply, and reject the change in Word. |

## Not in this sample

- No secrets, tenant ids or server URL are committed. Provisioning writes ids into
  `env/.env.dev`; keep that copy local.
- It does not host the server. See [Option B in deployment](../../docs/deployment.md#option-b---hosted-http-for-every-client).
- It is configured for dynamic tool discovery, so it uses whatever tools the server offers. To
  pin a fixed set, see
  [dynamic tool discovery](https://learn.microsoft.com/microsoft-365/copilot/extensibility/plugin-dynamic-tool-discovery).
