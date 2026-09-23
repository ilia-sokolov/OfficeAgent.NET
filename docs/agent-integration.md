# Agent integration

This adapter is optional. If your application needs document automation but not
model tool calling, use `OfficeAgent.Core` and a format module directly. See
[Choose OfficeAgent.NET](choose-officeagent.md) for package selection, non-goals,
and a complete direct .NET edit before adding an agent adapter.

`OfficeAgent.AgentFramework` exposes the OfficeAgent workflow as
Microsoft.Extensions.AI `AIFunction` tools through `OfficeAgentTools`. The tools
address documents by `(connectionId, documentId)` and route every call through
`OfficeAgentClient`, so the default tool surface gives the model opaque ids and
never credentials. Opt-in registration and composite tools also accept a path
or SharePoint source from the model. In either case, the provider enforces the
configured connection boundary. By default the host
pre-registers documents (`OfficeAgentClient.RegisterAsync`) and threads the
resulting opaque id into the agent's system prompt; hosts that want the agent to
stage its own ids opt in to the registration tools (below).

The tools use OpenAI / Azure OpenAI strict-mode schemas. Every outcome -
including bad input - is returned as structured JSON, so the model gets an error
it can read and react to instead of an exception.


Structural validation is not native Office acceptance and neither is visual parity. See
[what "it worked" means](getting-started.md#what-it-worked-means-here) before reporting a
document as verified.

## Wire up

Install the integration, format modules, and dependency-injection container:

```bash
dotnet add package OfficeAgent.AgentFramework
dotnet add package OfficeAgent.Word
dotnet add package OfficeAgent.PowerPoint
dotnet add package OfficeAgent.Excel
dotnet add package Microsoft.Extensions.DependencyInjection
```

```csharp
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.PowerPoint;
using OfficeAgent.Excel;
using OfficeAgent.Word;

var services = new ServiceCollection()
    .AddWordFormat()
    .AddPowerPointFormat()          // drop if the agent only handles .docx
    .AddExcelFormat()               // drop if the agent only handles .docx/.pptx
    .AddFileSystemDocumentProvider("workspace", "/srv/officeagent/workspace", o =>
        o.AllowedExtensions = new[] { ".docx", ".pptx", ".xlsx" })
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

The default connection-addressed surface has seven read, plan, comparison, and edit tools.
Registration, creation, inline-content, and session-transfer tools are separate opt-ins; an
agent cannot supply file paths or create outputs unless the host grants the corresponding
surface.

The function schemas use strict mode: every listed argument is required on the
wire even when it has a semantic default. Send the defaults shown below rather
than omitting a property.

### Default connection-addressed tools

| Tool | Purpose |
| --- | --- |
| `describe_capabilities()` | Lists the accepted contract, formats, operations, limits, rendering availability, and only the connections and capabilities visible to the caller. Call it before planning when the surface is not already known. |
| `inspect_document(connectionId, documentId, fidelity, paragraphOffset, paragraphLimit, sheetId, range, maximumCells)` | Returns format-specific structure and a snapshot etag. Send `"content"`, `0`, `200`, `0`, `""`, and `1000` for the defaults. Paragraph paging applies to Word/PowerPoint; `sheetId`, `range`, and `maximumCells` bound Excel inspection. |
| `find_in_document(connectionId, documentId, pattern, regex, wholeWord, caseSensitive, spreadsheetValueView)` | Returns content-verified anchors usable as plan targets. Send `false` for each search flag and `"both"` for the spreadsheet value default. |
| `preview_plan(connectionId, documentId, planJson)` | Validates a `DocumentPlan` JSON without writing. Returns the canonical plan-report envelope below, with `committed: false` and null output fields. |
| `apply_plan(connectionId, documentId, planJson, saveMode, newName)` | Applies the complete plan in memory and then saves through the provider. Send `"Replace"` and `""` for the defaults; other modes are `NewVersion` and `NewDocument`. Always inspect `writeOutcome`; a provider can fail after accepting the bytes. |
| `compare_documents(originalConnectionId, originalDocumentId, revisedConnectionId, revisedDocumentId, revisionAuthor)` | Reads two Word files and returns hashes, differences, coverage diagnostics, and a snapshot-bound tracked plan when coverage is complete. Requires read access to both connections and writes nothing. |
| `preview_document_merge(requestJson)` | Previews ordered whole-document Word assembly, reports compatibility and remapping, and returns a hash-bound merge plan. Requires read access to every source and writes nothing. |

### Opt-in tool groups

| Option | Added tools | Boundary |
| --- | --- | --- |
| `AllowRegistration` | `register_document`, `remove_document`, `open_document`, `edit_document` | Lets the agent name a provider-relative source. The underlying file is never deleted by `remove_document`. |
| `AllowCreation` | `create_document`, `discover_template`, `preview_template_batch`, `populate_template_batch`, `merge_documents` | Lets the agent create new stored outputs. A preflight is available only when its corresponding commit can also be enabled. |
| `AllowInlineContent` | `create_document_content`, `inspect_document_content`, `edit_document_content` | Carries the complete document as base64 with no storage. Use for bounded, self-contained calls. |
| `AllowEphemeralDocuments` | `import_document_content`, `export_document_content` | Transfers bytes into and out of a session connection; ordinary default tools handle the edits between those calls. |

The workflow-specific contracts remain:

| Tool | Purpose |
| --- | --- |
| `discover_template(connectionId, documentId)` | Reports what a template can be bound to: scalar slots, image and native-chart media slots, repeating Word rows, ambiguity diagnostics, and the template hash. Writes nothing and requires read access only. Exposed with `AllowCreation`, because a preflight exists to precede a commit. |
| `preview_template_batch(connectionId, documentId, requestJson)` | Validates a whole batch and writes nothing, reporting every item rather than stopping at the first failure, with per-item operation, row, image and byte counts. Returns a token binding the preview to the exact template bytes and normalized batch. Read access only. |
| `populate_template_batch(connectionId, documentId, requestJson, expectedTokenJson)` | Resolves scalar, image and native-chart values and repeating Word rows and saves bounded, independent outputs with one receipt per item. Pass the token from `preview_template_batch` to refuse a commit whose template or batch changed since it was reviewed; send `""` to commit without that binding. Exposed only with `AllowCreation`; requires read and create access on the connection. |
| `merge_documents(planJson, destinationConnectionId, outputName)` | Revalidates every source and creates one new `.docx`. Exposed with `AllowCreation`; requires source read and destination create access. See [assembly](document-assembly.md). |

Plan reports always contain `isValid`, `committed`, `writeOutcome`, `possibleOutput`, `receipt`, `sourceDocumentId`,
`outputConnectionId`, `outputDocumentId`, `outputVersion`, `outputName`,
`outputContentType`, `changes`, and `errors`. Non-applicable values are `null`.
`writeOutcome` is `committed`, `notWritten`, `unknown`, or `writtenNotRegistered`.
Only `notWritten` proves storage did not change. The two uncertain outcomes include a
`possibleOutput` locator; preserve it and reconcile the destination before retrying.

`inspect_document` returns `snapshot` as an etag string. To detect drift in the
format's covered text hosts, copy it into the plan token:

```json
{
  "snapshot": { "eTag": "<snapshot string from inspect_document>" },
  "operations": []
}
```

For Word, the etag covers body/header/footer/footnote/endnote XML; for PowerPoint,
slide and notes XML. For Excel, it covers workbook XML, shared strings, worksheets,
table definitions, and legacy notes. It does not cover every package part: properties and
media bytes are excluded across formats, and PowerPoint masters/layouts and Excel styles,
drawings, charts, external links, and pivot parts are also excluded. Omit `snapshot` only
when per-anchor and provider-version checks are sufficient. The engine does not insert it
automatically.

## Let the agent stage its own documents

When the user names files the host has not staged - "open the contract in the
legal library and fix the payment terms" - or asks for a document that does not
exist yet, the agent needs a way to mint ids itself. `OfficeAgentToolsOptions`
offers separate least-privilege switches for existing documents and new ones:

| Tool | Opt-in | Purpose |
| --- | --- | --- |
| `register_document(connectionId, source)` | `AllowRegistration` | Registers an existing document with a configured connection and returns its opaque `documentId`. `source` is connection-specific: a path under the filesystem connection's root, or - for a SharePoint connection - the document's SharePoint/OneDrive URL or a `driveId/itemId` pair. Filesystem traversal, disallowed extensions, and oversized files are rejected by the provider. |
| `remove_document(connectionId, documentId)` | `AllowRegistration` | Removes the registration only - the underlying file is never deleted. |
| `open_document(connectionId, source, fidelity, paragraphOffset, paragraphLimit, sheetId, range, maximumCells)` | `AllowRegistration` | `register_document` + `inspect_document` in one call. Send the same explicit inspection defaults as `inspect_document`. |
| `edit_document(connectionId, source, planJson, saveMode, newName)` | `AllowRegistration` | `register_document` + anchor resolution + `apply_plan` in one call. Send `"Replace"` and `""` for the defaults. Targets may name text directly instead of a paragraph id. |
| `create_document(connectionId, name, planJson)` | `AllowCreation` | Creates a **new** document in the connection, registers it, and optionally applies an initial plan in the same call. Pass `""` for no initial plan. Returns the `apply_plan` shape, so the new id arrives as `outputDocumentId`. `name` is a bare file name with its extension; a name already in use is refused rather than overwritten, and an initial plan that fails validation creates nothing at all. |

A new `.docx` contains one empty paragraph, addressed as paragraph id
`auto-0000`, over a style catalogue carrying `Heading1`–`Heading3`,
`ListParagraph` and `TableGrid`, so `styleId` works from the first operation. An
initial plan can fill it with `changeText` targeting
`{ "paraId": "auto-0000", "expect": "" }`. Alternatively, use an `insert` with
`"position": "Before"` to preserve the empty anchor as a trailing paragraph. A
new **deck** uses its initial slide anchor; see
[PowerPoint support](powerpoint.md#creating-a-deck).

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
        connectionId, documentId, cancellationToken);
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

1. `inspect_document` → understand the structure, capture the snapshot etag, and copy it to `plan.snapshot.eTag`.
2. `find_in_document` → obtain content-verified anchors for any text targets.
3. Draft a `DocumentPlan` referencing those anchors.
4. `preview_plan` → surface any validation errors to the user.
5. `apply_plan` → inspect `writeOutcome`. On `committed`, use the returned `outputDocumentId` for follow-up edits. On `unknown` or `writtenNotRegistered`, stop and reconcile `possibleOutput`; never retry blindly.

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
| `stale-snapshot` | Covered text-host or slide/notes XML drifted since inspection. Call `inspect_document` again before retrying. |
| `expect-mismatch` | A text anchor's expected content is no longer in the live document. Re-find that anchor. |
| `not-found` / `access-denied` | The supplied `documentId` is wrong or outside the connection's reach. |
| `version-conflict` | A `Replace` save lost a race. Re-inspect and re-author the plan. |
| `content-too-large`, `extension-not-allowed` | Provider policy refused the input. |
| `already-exists` | A `create_document` name is taken. Nothing was overwritten; retry with a different name. |
| `ambiguous-anchor` | An `edit_document` `find` target matched several times. The message lists each candidate; re-issue with `"match": <index>` or more surrounding text. Nothing was written. |
| `anchor-not-found` | A `find` target matched nothing, or its `match` index was out of range. Check the wording with `inspect_document` rather than retrying the same text. |
| `contract-mismatch` | The edit plan must omit `contractVersion` for legacy `0.2` behavior or set it to `"0.2"`. Null, empty, malformed, and unknown versions are refused before saving. |
| `invalid-argument`, `invalid-json` | The plan or arguments were malformed. Unknown properties, operations, enum names, integer enum values, and non-string versions are refused. The error message says what to fix. |
| `configuration-error` | The `connectionId` is not registered on this host, or - for `create_document` - that connection cannot create documents. Try another connection rather than retrying. |
| `outcome-unknown` | A write began but the provider could not confirm whether it landed. Do not retry or report the document unchanged; reconcile `possibleOutput.expectedSha256`. |
| `registration-failed` | The bytes were written but the output registration failed. Do not write again; recover the named output from `possibleOutput`. |
| `write-rejected` | Storage definitely rejected the write, so `writeOutcome` is `notWritten`. Correct the storage condition before retrying. |
| `cancelled` | The call was cancelled before a write began. A cancellation after writing begins is instead `outcome-unknown`. |
| `connection-forbidden` | Host policy denied the required capability. Do not probe other ids; use an authorized connection or ask the host to grant access. |

Every error also carries `connectionId` and `itemId` (when known) so the agent can correlate the failure to a specific call.

## Prompt guidance

`OfficeAgentTools.SystemPromptGuidance` is a `const string` you concatenate into your agent's instructions. It teaches the model the host-registered `(connectionId, documentId)` contract, the safety loop (re-inspect on stale snapshot, re-find on expect mismatch), uncertain-write recovery, the `Tracked` default for Word and why a deck refuses it, how a deck is addressed, that saving replaces the document in place, and the rule that anchors and node paths come from the engine - never invented. Append `OfficeAgentTools.RegistrationPromptGuidance` when the registration tools are enabled, and `OfficeAgentTools.CreationPromptGuidance` when `create_document` is among them - each block should be present only when its tools are.

## Use OfficeAgent over MCP

The same core tool surface is available to any MCP-capable agent (Claude code, Codex Copilot Studio agents etc.) through the
[`OfficeAgent.Mcp` server](mcp-server.md) - stdio for local hosting, streamable
HTTP for the cloud. When registration or creation is enabled, the MCP server also
adds `list_connections`, which returns
`{connectionId, provider, canCreateDocuments}` entries. The boolean means that
at least one creatable format is configured; it does not enumerate formats or
prove that the provider is reachable and authorized. The MCP server advertises this
prompt guidance as its server instructions, so MCP clients pick up the contract
automatically.
