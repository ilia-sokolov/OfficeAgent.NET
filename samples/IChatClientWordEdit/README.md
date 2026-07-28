# IChatClientWordEdit - direct IChatClient integration

This minimal console sample exposes OfficeAgent.NET's four bounded Word tools
directly to a `Microsoft.Extensions.AI.IChatClient`. It does not use
`ChatClientAgent` or Microsoft Agent Framework.

The sample:

1. creates a synthetic Word document in a temporary workspace;
2. registers it under an opaque document id;
3. asks the model to change `60 days` to `30 days` as a tracked change;
4. logs each OfficeAgent tool call;
5. captures `outputDocumentId` from `apply_plan`; and
6. retrieves the saved document without sending its bytes through the model.

The source fixture is checked after the edit and the temporary workspace is
deleted. Only `reviewed-contract.docx`, or the output path you supply, remains.

## Prerequisites

- .NET 8 SDK
- An Azure OpenAI deployment whose model supports function calling
- Azure CLI login with `az login`
- The signed-in identity must have the `Cognitive Services OpenAI User` role on
  the Azure AI resource

## Run

Set the endpoint and deployment name, then run the project from the repository
root.

PowerShell:

```powershell
$env:AZURE_OPENAI_ENDPOINT = "https://<resource>.openai.azure.com/"
$env:AZURE_OPENAI_DEPLOYMENT = "<deployment-name>"
$env:AZURE_TOKEN_CREDENTIALS = "AzureCliCredential"
dotnet run --project samples/IChatClientWordEdit
```

Bash:

```bash
export AZURE_OPENAI_ENDPOINT="https://<resource>.openai.azure.com/"
export AZURE_OPENAI_DEPLOYMENT="<deployment-name>"
export AZURE_TOKEN_CREDENTIALS="AzureCliCredential"
dotnet run --project samples/IChatClientWordEdit
```

Pass a path after `--` to choose the output location:

```bash
dotnet run --project samples/IChatClientWordEdit -- ./my-reviewed-contract.docx
```

The trace should show discovery, preview, and apply. A model can use
`inspect_document` alone when it already has the exact target, or add
`find_in_document` when it needs a content-verified anchor.

Open the output in Microsoft Word to see one tracked replacement: deletion of
`60 days` and insertion of `30 days`.

## Why the result is normalized

Depending on the function adapter path, an OfficeAgent JSON result can surface
as a JSON object or as a JSON string inside a `JsonElement`.
`NormalizeToolResult` converts both forms to a structured value. This lets the
model consume the tool result and lets the host reliably capture
`outputDocumentId` without parsing the assistant's prose.
