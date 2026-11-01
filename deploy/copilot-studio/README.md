# Copilot Studio connector

Connector definitions for a hosted OfficeAgent.NET MCP server. They give an administrator
something to version, review and import. Copilot Studio's MCP onboarding wizard is simpler for
one agent, and is covered in [deployment and client setup](../../docs/deployment.md).

| File | Use with |
| --- | --- |
| `officeagent-oauth.swagger.yaml` | `onBehalfOf`, acting as the signed-in user. This matches the [Azure template](../azure/README.md) |
| `officeagent-apikey.swagger.yaml` | `appOnly` behind a gateway that checks an API key, with a shared identity. The Azure template does not set this up |

## Before you import

You need a hosted server. The server itself has no authentication, so it must sit behind the Azure
template's authentication or your own gateway. It needs these settings; the Azure template sets them:

| Setting | Value | Why |
| --- | --- | --- |
| `OfficeAgent__StrictToolSchemas` | `false` | Copilot Studio asks the user for every required input the model leaves out |
| `OfficeAgent__RequirePlanSnapshot` | `true` | A retried call cannot apply the same plan twice |

Edit the file you are using and replace the placeholders:

| Placeholder | Value |
| --- | --- |
| `YOUR_SERVER_HOST`, `YOUR_GATEWAY_HOST` | The host of your server, with no `https://`, from the deployment's `serverUrl` output |
| `YOUR_TENANT_ID` | Your Entra tenant id |
| `YOUR_API_CLIENT_ID` | Client id of the middle-tier API app |

For the API-key file, change `X-Api-Key` if your gateway reads another header.

## Import it

1. In Copilot Studio, open your agent's **Tools** page and select **Add a tool**.
2. Select **New tool**, then **Custom connector**. Power Apps opens.
3. Select **New custom connector**, then **Import OpenAPI file**, and choose your edited file.
4. Select **Continue** and finish the connector. For OAuth, enter a client id and secret for
   Copilot Studio to sign in with, then copy the **Redirect URL** Power Apps shows into that app
   registration, and allow it to request the `access_as_user` scope.
5. Back on **Add a tool**, select the connector and create a connection.
6. Select **Add to agent**, then turn off any tools the agent should not have. Agents need
   **generative orchestration**.

Access to the server goes through a Power Platform connector, so your data policies apply to it.

## Public catalog

These are custom connectors for your own environment. Listing OfficeAgent in Microsoft's public
connector catalog is a separate step. It would need a hosted endpoint owned by the publisher,
publisher verification, and Microsoft's review. There is no shared hosted OfficeAgent server,
because each organisation runs its own, so that does not apply yet.

Microsoft's current steps are in
[Connect your agent to an existing MCP server](https://learn.microsoft.com/microsoft-copilot-studio/mcp-add-existing-server-to-agent).
