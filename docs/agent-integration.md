# Agent integration

`OfficeAgent.AgentFramework` exposes the OfficeAgent workflow as
Microsoft.Extensions.AI `AIFunction` tools through `OfficeAgentTools`. The tools
address documents by `(connectionId, documentId)` and route every call through
`OfficeAgentClient`, so the language model never sees a file path, cannot register
or delete documents, and cannot leave the storage connection the host configured.
The host pre-registers documents (`OfficeAgentClient.RegisterAsync`) and threads
the resulting opaque id into the agent's system prompt.

The tools use OpenAI / Azure OpenAI strict-mode schemas. Every outcome -
including bad input - is returned as structured JSON, so the model gets an error
it can read and react to instead of an exception.

## Wire up

```csharp
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Word;

var services = new ServiceCollection()
    .AddWordFormat()
    .AddFileSystemDocumentProvider("workspace", "/srv/officeagent/workspace")
    .AddOfficeAgent()
    .BuildServiceProvider();

var client = services.GetRequiredService<OfficeAgentClient>();

// Host registers the file with the connection before the conversation starts.
var seeded = await client.RegisterAsync("workspace", "/srv/officeagent/workspace/contract.docx");

var tools     = new OfficeAgentTools(client).AsAIFunctions();
var prompt    = $"You are editing the document with documentId={seeded.ItemId} on connectionId=workspace.\n\n"
              + OfficeAgentTools.SystemPromptGuidance;
```

The `seeded.ItemId` goes into the system prompt; the LLM threads it through every subsequent tool call.

## Exposed tools

The agent surface is deliberately read-and-edit only. Registration and removal
are host responsibilities - the agent cannot supply file paths.

| Tool | Purpose |
| --- | --- |
| `inspect_document(connectionId, documentId, fidelity?, paragraphOffset?, paragraphLimit?)` | Returns outline, paragraphs (with their containing table when applicable), content controls, nodes (tables, images, document properties, revisions), styles, and a snapshot etag. Pages large documents. |
| `find_in_document(connectionId, documentId, pattern, regex?, wholeWord?, caseSensitive?)` | Returns content-verified anchors usable as plan targets. |
| `preview_plan(connectionId, documentId, planJson)` | Validates a `DocumentPlan` JSON and returns proposed changes and errors without writing. |
| `apply_plan(connectionId, documentId, planJson, saveMode?, newName?)` | Applies the plan atomically and saves through the provider. Returns `outputConnectionId`, `outputDocumentId`, `outputVersion`, `outputName`, and `outputContentType`. |

## Return the final document to the user

`apply_plan` saves the committed document through the provider and returns an opaque `outputDocumentId`; it deliberately does **not** send `.docx` bytes through the model context. The application hosting the agent owns delivery to the user:

1. Capture `outputConnectionId` and `outputDocumentId` from the successful `apply_plan` tool result.
2. Retrieve the canonical content with `OfficeAgentClient.OpenReadAsync`.
3. Send the stream or copied bytes through the channel's native file/attachment API.

For example, an ASP.NET Core download endpoint can copy the provider stream into the HTTP response payload:

```csharp
app.MapGet("/documents/{connectionId}/{documentId}", async (
    string connectionId,
    string documentId,
    OfficeAgentClient client,
    CancellationToken cancellationToken) =>
{
    using var content = await client.OpenReadAsync(
        DocumentReference.ForFileSystem(connectionId, documentId),
        cancellationToken);
    using var buffer = new MemoryStream();
    await content.Stream.CopyToAsync(buffer, cancellationToken);

    return Results.File(
        buffer.ToArray(),
        content.Reference.ContentType ?? "application/octet-stream",
        content.Reference.Name ?? "document.docx");
});
```

A chat, Teams, Slack, or custom UI host follows the same pattern but passes `content.Stream` (or copied bytes) to its attachment API. The assistant's final text should say that the document is ready; the host should render the resulting attachment or download link. Avoid returning the document as base64 in assistant text because that consumes model context, can be truncated, and bypasses the host's normal file-delivery controls.

## Microsoft Agent Framework

Hand the tools straight to a `ChatClientAgent`:

```csharp
using Microsoft.Agents.AI;
using Microsoft.Extensions.AI;

AIAgent agent = new ChatClientAgent(
    chatClient,                       // any Microsoft.Extensions.AI IChatClient
    instructions:   prompt,
    name:           "OfficeAgent",
    description:    "Edits Word documents using the OfficeAgent.NET toolkit.",
    tools:          tools.Cast<AITool>().ToList(),
    services:       services);
```

`UseFunctionInvocation()` on the underlying chat client will execute the tool calls automatically. A complete Azure OpenAI sample is in [`samples/AgentEdit`](../samples/AgentEdit).

## Recommended agent loop

The host pre-registers the document and writes the resulting `(connectionId, documentId)` into the system prompt. The agent then:

1. `inspect_document` → understand the structure and capture the snapshot etag.
2. `find_in_document` → obtain content-verified anchors for any text targets.
3. Draft a `DocumentPlan` referencing those anchors.
4. `preview_plan` → surface any validation errors to the user.
5. `apply_plan` → commit, then use the returned `outputDocumentId` for any follow-up edits.

## Errors the LLM can act on

| Code | Meaning |
| --- | --- |
| `stale-snapshot` | The document drifted since inspection. Call `inspect_document` again before retrying. |
| `expect-mismatch` | A text anchor's expected content is no longer in the live document. Re-find that anchor. |
| `not-found` / `access-denied` | The supplied `documentId` is wrong or outside the connection's reach. |
| `version-conflict` | A `Replace` save lost a race. Re-inspect and re-author the plan. |
| `content-too-large`, `extension-not-allowed` | Provider policy refused the input. |
| `invalid-argument`, `invalid-json` | The plan or arguments were malformed. The error message says what to fix. |
| `configuration-error` | The `connectionId` is not registered on this host. |

Every error also carries `connectionId` and `itemId` (when known) so the agent can correlate the failure to a specific call.

## Prompt guidance

`OfficeAgentTools.SystemPromptGuidance` is a `const string` you concatenate into your agent's instructions. It teaches the model the host-registered `(connectionId, documentId)` contract, the safety loop (re-inspect on stale snapshot, re-find on expect mismatch), the default `Tracked` change mode, and the rule that anchors and node paths come from the engine - never invented.
