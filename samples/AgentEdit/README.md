# AgentEdit - MAF + OfficeAgent.NET (Azure OpenAI)

An interactive console agent that edits a Microsoft Word document in response to natural-language requests. Uses [Microsoft Agent Framework](https://www.nuget.org/packages/Microsoft.Agents.AI/) (MAF 1.x), Azure OpenAI as the model backend, and OfficeAgent.NET as the tool layer.

The host stages one document under a sandboxed filesystem provider's root and registers it to receive an opaque id; the LLM only ever sees that id and addresses operations by `(connectionId, documentId)`. The agent cannot register, delete, or otherwise touch storage - the provider is a registry of references the host owns.

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

## Try these prompts

- *"What's in the document? Give me a one-paragraph summary."*
- *"Replace every occurrence of 'Acme Corp' with 'Globex Inc.' as a tracked change."*
- *"Highlight 'Effective date' in yellow."*
- *"Add a row to the milestones table: Beta, 2026-08-15."*
- *"Set the document title to 'Service Agreement v2'."*
- *"Fill the ClientName content control with 'Globex'."*
- *"Accept all the tracked changes you made."*

Type `export <output.docx>` to retrieve the current revision from provider storage and write it to disk, or `quit` to exit. In a web or chat host, the equivalent code sends that stream through the platform's download/attachment API.
