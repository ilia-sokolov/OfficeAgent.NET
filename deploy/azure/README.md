# Deploy to Azure Container Apps

A repeatable deployment of the OfficeAgent.NET MCP server for hosted clients such as
Copilot Studio and Microsoft 365 Copilot. It acts as the signed-in user against SharePoint.

[![Deploy to Azure](https://aka.ms/deploytoazurebutton)](https://portal.azure.com/#create/Microsoft.Template/uri/https%3A%2F%2Fraw.githubusercontent.com%2Filia-sokolov%2FOfficeAgent.NET%2Fmain%2Fdeploy%2Fazure%2Fazuredeploy.json)

## What it creates

| Resource | Purpose |
| --- | --- |
| Container App | The server, from the pinned image in the `image` parameter |
| Container Apps authentication | Requires a valid Entra token on every call except `/healthz`, and forwards it unchanged |
| Container Apps environment and Log Analytics workspace | Hosting and logs, kept for 30 days |
| Storage account with a file share at `/data` | The registration index, so registered documents survive restarts |

It runs one replica, because the registration index is a single file, and it costs money while
it runs. Set `minReplicas` to `0` to scale to zero; the first call after idle is then slow.

## Before you deploy

Create the middle-tier API app registration in Entra ID:

1. Register an application and note its **Application (client) ID**.
2. Under **Expose an API**, set the Application ID URI to `api://<client-id>` and add a scope
   named `access_as_user`.
3. Under **API permissions**, add the delegated Microsoft Graph permissions your documents need,
   such as `Files.ReadWrite.All`, and grant admin consent.
4. Under **Certificates & secrets**, create a client secret.
5. Under **Token configuration**, make sure the token includes the `name` claim if you want
   receipts to carry display names.

Every user in the tenant who can get a token for this app can call the server. Each one acts as
themselves, so SharePoint's own permissions still apply. To limit who can call it, turn on
**Assignment required** on the app's enterprise application and assign users or groups.

## Deploy

Use the **Deploy to Azure** button above, or the Azure CLI:

```bash
az group create --name officeagent --location westeurope
```

```bash
az deployment group create \
  --resource-group officeagent \
  --template-file deploy/azure/main.bicep \
  --parameters apiClientId=<client-id> apiClientSecret=<client-secret>
```

The deployment's `serverUrl` output is the MCP endpoint. A call without a token should return
`401`, and `/healthz` should return `200`.

## Parameters

| Parameter | Default | Meaning |
| --- | --- | --- |
| `apiClientId` | required | Client id of the app registration |
| `apiClientSecret` | required | Its secret, stored as a Container Apps secret |
| `image` | the pinned release | Server image. Keep it on an exact version |
| `connectionId` | `sharepoint` | Connection id agents address documents under |
| `minReplicas` | `1` | `1` stays warm, `0` scales to zero |
| `revisionAuthorFromActor` | `false` | Author tracked changes and comments with the signed-in user's name. See [identity](../../docs/deployment.md#authentication--identity) |
| `tenantId` | the subscription's tenant | Tenant of the signed-in users |
| `name`, `location` | `officeagent`, the resource group's | Naming and region |

## Connect a client

For Copilot Studio, use the [connector and its README](../copilot-studio/README.md).

The template sets `StrictToolSchemas=false` and `RequirePlanSnapshot=true`, the settings Copilot
Studio needs. See [deployment and client setup](../../docs/deployment.md) for other clients.

## What this does not do

- It does not authorize per connection. The stock server image allows every authenticated caller
  to use the connection, as themselves. A host needing finer control registers an
  `IConnectionAccessPolicy`; see the
  [HostedGateway sample](../../samples/HostedGateway/).
- It does not implement Conditional Access claims challenges. See the note in
  [deployment](../../docs/deployment.md#b2-put-authentication-in-front).
- It does not set up private networking, backups, or secret rotation.

After changing `main.bicep`, regenerate `azuredeploy.json`, which the button uses:

```bash
az bicep build --file deploy/azure/main.bicep --outfile deploy/azure/azuredeploy.json
```
