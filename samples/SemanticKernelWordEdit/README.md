# SemanticKernelWordEdit - OfficeAgent tools in Semantic Kernel

This console sample gives a [Semantic Kernel](https://learn.microsoft.com/semantic-kernel/)
`Kernel` OfficeAgent.NET's Word tools as a plugin and lets the model make one tracked edit. It is
the Semantic Kernel counterpart of [IChatClientWordEdit](../IChatClientWordEdit/), and uses the
same document and task as the Python samples,
[python-openai-agents](../python-openai-agents/) and [python-langgraph](../python-langgraph/).

`OfficeAgentTools.AsAIFunctions()` returns `Microsoft.Extensions.AI` functions. Semantic Kernel
takes each one as a kernel function with `AsKernelFunction()`, and
`Plugins.AddFromFunctions("officeagent", ...)` makes them one plugin:

```csharp
AIFunction[] tools = new OfficeAgentTools(office).AsAIFunctions();
IKernelBuilder builder = Kernel.CreateBuilder();
builder.Plugins.AddFromFunctions("officeagent", tools.Select(tool => tool.AsKernelFunction()));
```

The sample:

1. copies [`services-agreement.docx`](../documents/) into a temporary workspace;
2. registers it under an opaque document id;
3. puts `OfficeAgentTools.SystemPromptGuidance` in the system message, with the connection and
   document id;
4. asks the model to change the initial term in clause 5 from `twelve months` to
   `twenty-four months` as a tracked change, leaving clause 4, which also says twelve months,
   as it is;
5. logs each tool call through an `IAutoFunctionInvocationFilter` and captures
   `outputDocumentId` from `apply_plan`; and
6. retrieves the saved document without sending its bytes through the model, and reports
   whether it contains the tracked insertion and deletion.

The source copy is checked after the edit and the temporary workspace is deleted. Only
`reviewed-services-agreement.docx`, or the output path you supply, remains.

## Prerequisites

- .NET 10 SDK (the repository's `global.json` selects it)
- One of:
  - an Azure OpenAI deployment whose model supports function calling, with either an API key
    or an Azure CLI login (`az login`) by an identity that has the
    `Cognitive Services OpenAI User` role on the resource; or
  - an OpenAI API key.

The model is configured through environment variables only:

| Variable | Used for |
| --- | --- |
| `AZURE_OPENAI_ENDPOINT` | Azure OpenAI resource, such as `https://<resource>.openai.azure.com/`. When it is set, Azure OpenAI is used. |
| `AZURE_OPENAI_DEPLOYMENT` | Azure OpenAI deployment name. |
| `AZURE_OPENAI_API_KEY` | Optional. Without it the sample signs in with `DefaultAzureCredential`. |
| `OPENAI_API_KEY` | OpenAI, when `AZURE_OPENAI_ENDPOINT` is not set. |
| `OPENAI_MODEL` | Optional OpenAI model; the default is `gpt-5.6-luna`. |

Azure OpenAI is called through its v1 API, `<endpoint>/openai/v1/`, with the OpenAI client that
Semantic Kernel's OpenAI connector uses.

## Run

Set the variables, then run the project from the repository root.

PowerShell, Azure OpenAI with `az login`:

```powershell
$env:AZURE_OPENAI_ENDPOINT = "https://<resource>.openai.azure.com/"
$env:AZURE_OPENAI_DEPLOYMENT = "<deployment-name>"
$env:AZURE_TOKEN_CREDENTIALS = "AzureCliCredential"
dotnet run --project samples/SemanticKernelWordEdit
```

Bash, OpenAI:

```bash
export OPENAI_API_KEY="<your key>"
dotnet run --project samples/SemanticKernelWordEdit
```

Pass a path after `--` to choose the output location:

```bash
dotnet run --project samples/SemanticKernelWordEdit -- ./my-reviewed-contract.docx
```

To check the plugin without a model or a key, list its functions and call `inspect_document`
through the kernel:

```bash
dotnet run --project samples/SemanticKernelWordEdit -- --list-functions
```

## What you should see

The trace shows each function the model called, then the change report and the check of the
saved file:

```text
tool: inspect_document
tool: find_in_document
tool: preview_plan
tool: apply_plan
changes: [{"verb":"changeText", ... "before":"continues for twelve months","after":"continues for twenty-four months", ...}]
...
Saved .../reviewed-services-agreement.docx
Tracked insertion of 'twenty-four months': True
Tracked deletion of 'twelve months':       True
```

Which functions the model calls varies; `apply_plan` must commit. Open the output in Microsoft
Word and open **Review > Reviewing Pane**: clause 5 has one deletion, `twelve months`, and one
insertion, `twenty-four months`. The document's existing clause 3 revision and the comment on
clause 4 are unchanged.

## Notes

- The tools keep their OfficeAgent names inside the plugin. Semantic Kernel's OpenAI connector
  sends them to the model with the plugin name in front, as `officeagent-inspect_document`,
  while `SystemPromptGuidance` names them without it.
- `AllowParallelCalls = false` keeps the model to one call at a time, so it previews before it
  applies.
- An OfficeAgent result can arrive as a JSON object or as JSON text inside a string; the filter
  reads both, as `IChatClientWordEdit` does.
