# Agent integration

`OfficeAgent.AgentFramework` exposes the OfficeAgent workflow as
Microsoft.Extensions.AI `AIFunction` tools through `OfficeAgentTools`. The tools
address documents by `(connectionId, documentId)` and route every call through
`OfficeAgentClient`, so the language model never sees a file path or credential
and cannot leave the storage connection the host configured. By default the host
pre-registers documents (`OfficeAgentClient.RegisterAsync`) and threads the
resulting opaque id into the agent's system prompt; hosts that want the agent to
stage its own ids opt in to the registration tools (below).

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

The default agent surface is read-and-edit only - registration and removal are
host responsibilities, and the agent cannot supply file paths.

| Tool | Purpose |
| --- | --- |
| `inspect_document(connectionId, documentId, fidelity?, paragraphOffset?, paragraphLimit?)` | Returns outline, paragraphs (with their containing table when applicable), content controls, nodes (tables, images, document properties, revisions), styles, and a snapshot etag. Pages large documents. |
| `find_in_document(connectionId, documentId, pattern, regex?, wholeWord?, caseSensitive?)` | Returns content-verified anchors usable as plan targets. |
| `preview_plan(connectionId, documentId, planJson)` | Validates a `DocumentPlan` JSON without writing. Returns `{ isValid, changes, errors }`. |
| `apply_plan(connectionId, documentId, planJson, saveMode?, newName?)` | Applies the plan atomically and saves through the provider. `saveMode` is `Replace` (default), `NewVersion`, or `NewDocument` (with `newName`); an unrecognised value is refused rather than defaulted. Returns `{ isValid, committed, outputConnectionId, outputDocumentId, outputVersion, outputName, outputContentType, changes, errors }`. |

## Let the agent stage its own documents

When the user names files the host has not staged - "open the contract in the
legal library and fix the payment terms" - or asks for a document that does not
exist yet, the agent needs a way to mint ids itself. `OfficeAgentToolsOptions`
offers separate least-privilege switches for existing documents and new ones:

| Tool | Opt-in | Purpose |
| --- | --- | --- |
| `register_document(connectionId, source)` | `AllowRegistration` | Registers an existing document with a configured connection and returns its opaque `documentId`. `source` is connection-specific: a path under the filesystem connection's root, or - for a SharePoint connection - the document's SharePoint/OneDrive URL or a `driveId/itemId` pair. Filesystem traversal, disallowed extensions, and oversized files are rejected by the provider. |
| `remove_document(connectionId, documentId)` | `AllowRegistration` | Removes the registration only - the underlying file is never deleted. |
| `open_document(connectionId, source, fidelity?, paragraphOffset?, paragraphLimit?)` | `AllowRegistration` | `register_document` + `inspect_document` in one call. Returns `{ connectionId, documentId, name, contentType, version }` followed by the whole `inspect_document` payload. |
| `edit_document(connectionId, source, planJson, saveMode?, newName?)` | `AllowRegistration` | `register_document` + anchor resolution + `apply_plan` in one call. Targets may name text directly instead of a paragraph id. Returns the `apply_plan` shape plus `sourceDocumentId`. |
| `create_document(connectionId, name, planJson)` | `AllowCreation` | Creates a **new** document in the connection, registers it, and optionally applies an initial plan in the same call. Pass `""` for no initial plan. Returns the `apply_plan` shape, so the new id arrives as `outputDocumentId`. `name` is a bare file name with its extension; a name already in use is refused rather than overwritten, and an initial plan that fails validation creates nothing at all. |

A blank document contains one empty paragraph, addressed as paragraph id
`auto-0000`. An initial plan can target
`{ "paraId": "auto-0000", "expect": "" }`; use `"position": "Before"` to keep
the empty anchor as the trailing paragraph.

### Addressing text instead of paragraph ids

The single-purpose loop is `find_in_document` to get an anchor, then `apply_plan`
to use it. `edit_document` folds that into one call: a target may name the text
itself, and the tool resolves it against live content before applying anything.

```jsonc
// Instead of: find_in_document → read paraId → apply_plan with that paraId
[ { "op": "changeText", "target": { "find": "Acme Corp" }, "with": "Globex Inc." } ]
```

**Text that matches more than once is refused**, not guessed at. The error is
`ambiguous-anchor` and lists every candidate with its surrounding context, so the
next call can name the one it meant:

```jsonc
[ { "op": "changeText", "target": { "find": "Acme Corp", "match": 1 }, "with": "Globex Inc." } ]
```

`match` is zero-based over the document-wide match list. Text matching nothing is
`anchor-not-found`. Every unresolvable target in a plan is reported in the same
result, so a plan with two bad targets costs one call to discover both, not two.

`find` is a literal, case-insensitive search. Regex, whole-word, and
case-sensitive matching stay on `find_in_document` - resolve there and pass the
resulting `paraId` targets, which `edit_document` accepts and can mix freely with
`find` targets in one plan.

`planJson` accepts a bare operations array `[ … ]` as well as
`{ "operations": [ … ] }`, on `edit_document`, `create_document`, `preview_plan`,
and `apply_plan` alike.

The switches are independent: a host may expose creation without letting the
agent register or remove arbitrary existing documents. Both remain off by default
for in-process tools:

```csharp
var tools = new OfficeAgentTools(client)
    .AsAIFunctions(new OfficeAgentToolsOptions
    {
        AllowRegistration = true,
        AllowCreation     = true    // off by default; omit for a read-and-edit agent
    });
var prompt = OfficeAgentTools.SystemPromptGuidance
           + OfficeAgentTools.RegistrationPromptGuidance
           + OfficeAgentTools.CreationPromptGuidance;   // only when AllowCreation = true
```

Creation writes under a filesystem connection's root or, for SharePoint, into
the connection's explicitly configured drive and folder. Initial-plan errors
happen before storage is touched. A later provider error can mean storage accepted
the file but registration did not finish, so the agent should not blindly retry
the same name.

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

A chat or custom-UI host follows the same pattern, passing `content.Stream` (or copied bytes) to its attachment API. Have the assistant's final text say the document is ready and let the host render the attachment or download link - don't return the document as base64 in assistant text (it consumes model context, can be truncated, and bypasses the host's file-delivery controls).

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

`UseFunctionInvocation()` on the underlying chat client will execute the tool
calls automatically. See
[`samples/IChatClientWordEdit`](../samples/IChatClientWordEdit) for a minimal
direct `IChatClient` Azure OpenAI host, or
[`samples/AgentEdit`](../samples/AgentEdit) for an interactive Agent Framework
host.

## Recommended agent loop

The host pre-registers the document and writes the resulting `(connectionId, documentId)` into the system prompt. The agent then:

1. `inspect_document` → understand the structure and capture the snapshot etag.
2. `find_in_document` → obtain content-verified anchors for any text targets.
3. Draft a `DocumentPlan` referencing those anchors.
4. `preview_plan` → surface any validation errors to the user.
5. `apply_plan` → commit, then use the returned `outputDocumentId` for any follow-up edits.

When the composite tools are enabled and the user names a file by path, the same
work is two calls or one:

- `open_document` → registers and inspects together; carry on from step 2.
- `edit_document` → registers, resolves `find` targets, and applies, when the
  edit is already known. Reach back for the single-purpose tools when you need a
  preview before writing, a regex or case-sensitive search, or an id you already
  hold.

## Errors the LLM can act on

| Code | Meaning |
| --- | --- |
| `stale-snapshot` | The document drifted since inspection. Call `inspect_document` again before retrying. |
| `expect-mismatch` | A text anchor's expected content is no longer in the live document. Re-find that anchor. |
| `not-found` / `access-denied` | The supplied `documentId` is wrong or outside the connection's reach. |
| `version-conflict` | A `Replace` save lost a race. Re-inspect and re-author the plan. |
| `content-too-large`, `extension-not-allowed` | Provider policy refused the input. |
| `already-exists` | A `create_document` name is taken. Nothing was overwritten; retry with a different name. |
| `ambiguous-anchor` | An `edit_document` `find` target matched several times. The message lists each candidate; re-issue with `"match": <index>` or more surrounding text. Nothing was written. |
| `anchor-not-found` | A `find` target matched nothing, or its `match` index was out of range. Check the wording with `inspect_document` rather than retrying the same text. |
| `invalid-argument`, `invalid-json` | The plan or arguments were malformed. The error message says what to fix. |
| `configuration-error` | The `connectionId` is not registered on this host, or - for `create_document` - that connection cannot create documents. Try another connection rather than retrying. |

Every error also carries `connectionId` and `itemId` (when known) so the agent can correlate the failure to a specific call.

## Prompt guidance

`OfficeAgentTools.SystemPromptGuidance` is a `const string` you concatenate into your agent's instructions. It teaches the model the host-registered `(connectionId, documentId)` contract, the safety loop (re-inspect on stale snapshot, re-find on expect mismatch), the default `Tracked` change mode, and the rule that anchors and node paths come from the engine - never invented. Append `OfficeAgentTools.RegistrationPromptGuidance` when the registration tools are enabled, and `OfficeAgentTools.CreationPromptGuidance` when `create_document` is among them - each block should be present only when its tools are.

## Use OfficeAgent over MCP

The same core tool surface is available to any MCP-capable agent (Claude code, Codex Copilot Studio agents etc.) through the
[`OfficeAgent.Mcp` server](mcp-server.md) - stdio for local hosting, streamable
HTTP for the cloud. When registration or creation is enabled, the MCP server also
adds `list_connections`, which returns
`{connectionId, provider, canCreateDocuments}` entries so the agent can discover
the available connections and creation capability. The MCP server advertises this
prompt guidance as its server instructions, so MCP clients pick up the contract
automatically.
