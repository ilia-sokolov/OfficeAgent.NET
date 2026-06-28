# AgentEdit - MAF + OfficeAgent.NET (Azure OpenAI)

An interactive console agent that edits a Microsoft Word document in response to natural-language requests. Uses [Microsoft Agent Framework](https://www.nuget.org/packages/Microsoft.Agents.AI/) (MAF 1.x), Azure OpenAI as the model backend, and OfficeAgent.NET as the tool layer.

The host registers one document with a provider connection to receive an opaque id; the LLM only ever sees that id and addresses operations by `(connectionId, documentId)`. The agent cannot register, delete, or otherwise touch storage - the provider is a registry of references the host owns.

The sample picks its storage provider automatically: by default a sandboxed **filesystem** provider (zero configuration), or the **SharePoint** provider when SharePoint configuration is supplied (see below). The agent code, prompt, and edit loop are identical either way - only the connection wiring changes.

## What it shows

- Wiring `OfficeAgentTools` into a MAF `ChatClientAgent` as `AIFunction[]`.
- `Microsoft.Extensions.AI`'s `UseFunctionInvocation()` executing the tool calls automatically.
- `OfficeAgentClient.RegisterAsync` minting an opaque id for a host-supplied path.
- Iterative edits across turns - each successful `apply_plan` returns an `outputDocumentId` the sample threads forward.
- Host-side final-document delivery: `export <output.docx>` resolves the active opaque id with `OpenReadAsync` and writes the bytes without routing them through the model.
- `OfficeAgentTools.SystemPromptGuidance` teaching the model the storage, anchor, and safety contracts.

## Run

```bash
export AZURE_OPENAI_ENDPOINT='https://<your-resource>.openai.azure.com'
export AZURE_OPENAI_DEPLOYMENT='<your-chat-deployment>'    # e.g. gpt-4o-mini
export AZURE_OPENAI_API_KEY='<key>'                         # optional; omit to use DefaultAzureCredential

dotnet run --project samples/AgentEdit
```

The sample generates `./sample.docx` on first run (a small contract fixture with a content control, a clause, and a milestones table). Override with `AGENT_DOC=/path/to/your.docx`. The provider's storage root is a per-run temp directory; override with `AGENT_STORAGE_DIR=/path/to/dir`.

## Run against SharePoint

Set `AGENT_SHAREPOINT_DOC` to switch the sample to the SharePoint provider. It then registers that **existing** document by its SharePoint/OneDrive URL or its `driveId/itemId` pair (it does not create content in your library), and runs the same edit loop against it. The document can live in any drive the configured identity can reach.

```bash
export AZURE_OPENAI_ENDPOINT='https://<your-resource>.openai.azure.com'
export AZURE_OPENAI_DEPLOYMENT='<your-chat-deployment>'

# The doc to edit: a SharePoint/OneDrive URL ...
export AGENT_SHAREPOINT_DOC='https://contoso.sharepoint.com/:w:/s/legal/EaBc123...'
# ... or a driveId/itemId pair:
# export AGENT_SHAREPOINT_DOC='b!<graph-drive-id>/01ABCDEF...'

# Authentication - the app-only (client-credentials) flow:
export AGENT_SHAREPOINT_TENANT_ID='<tenant-id>'
export AGENT_SHAREPOINT_CLIENT_ID='<app-registration-id>'
export AGENT_SHAREPOINT_CLIENT_SECRET='<secret>'             # prefer a secret store / env over source

# Optional: override the connection id (default "sharepoint")
export AGENT_SHAREPOINT_CONNECTION_ID='legal'

dotnet run --project samples/AgentEdit
```

A non-destructive save writes the edited revision as a versioned sibling derived from the source name (`<source>.v2.docx` - e.g. `contract.docx` → `contract.v2.docx`) in the same library; the source stays untouched. `export <output.docx>` works the same way - the host resolves the opaque id through the provider and writes the bytes locally. This sample uses app-only application identity (client credentials); for a multi-user host that should act as the signed-in user, configure the SharePoint provider's On-Behalf-Of mode instead (see [docs/document-providers.md](../../docs/document-providers.md)).

## Try these prompts

- *"What's in the document? Give me a one-paragraph summary."*
- *"Replace every occurrence of 'Acme Corp' with 'Globex Inc.' as a tracked change."*
- *"Highlight 'Effective date' in yellow."*
- *"Add a row to the milestones table: Beta, 2026-08-15."*
- *"Set the document title to 'Service Agreement v2'."*
- *"Fill the ClientName content control with 'Globex'."*
- *"Accept all the tracked changes you made."*

Type `export <output.docx>` to retrieve the current revision from provider storage and write it to disk, or `quit` to exit. In a web or chat host, the equivalent code sends that stream through the platform's download/attachment API.
