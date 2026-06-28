# OfficeAgent.NET

A translation layer between AI agents and OOXML - the format behind real
Microsoft Word documents. Agents express intent; the library turns it into
valid Open XML manipulations.

An agent can reason about a document, but it cannot safely produce one: the
things that make a `.docx` a Word document - styles, numbering, tracked
changes, comments, content controls - live in XML parts no language model
should write by hand. OfficeAgent.NET does that translation. The agent expresses
intent as a typed, JSON-serialisable change plan - "replace this clause as a
tracked change", "add a row to that table" - and the library translates the
plan into the Open XML changes that carry it out. Every change is validated
against the live document, previewed, and applied all-or-nothing. The agent
never writes raw `.docx` bytes, and Word itself is not automated.

The Word support is built on the Open XML SDK (`DocumentFormat.OpenXml`).

Use it when you need to:

- locate stable targets in a `.docx` (paragraphs, runs, tables, content
  controls, document properties, tracked revisions);
- ask a language model to return a typed edit plan instead of document bytes;
- preview the result before saving;
- keep Word semantics intact - runs, styles, content controls, comments,
  tracked revisions;
- reject edits that target the wrong place or that need a layout/calculation
  engine.

## Key concepts

### The workflow: inspect → find → preview → commit

Every edit follows the same four steps:

1. **Inspect** - read the document and get a structured map of it: the outline,
   paragraphs (with stable ids), styles, content controls, and nodes (tables,
   images, document properties, revisions).
2. **Find** - search for text and get an address back for each match.
3. **Preview** - check a set of changes against the current document. Nothing is
   written; you get a before/after report and any validation errors.
4. **Commit** - apply the changes and save through storage. The operation is
   all-or-nothing: if any step fails, nothing is written.

### Plan

A **plan** (`DocumentPlan`) is the list of changes you want to make. It is a
typed, JSON-serialisable object, so a language model can produce one directly.
The library validates the whole plan before it touches the document.

### Anchors

An **anchor** is an address inside the document. The library issues anchors from
inspect and find; the caller (or the model) reuses them and never invents one.

Anchors carry the content they expect to find. At commit time the library
re-checks each anchor against the live document. If the content has changed, the
operation fails safely instead of editing the wrong place.

### Operations

Each entry in a plan is one **operation** - a verb such as `changeText`,
`format`, `insert`, or `setProperty`. Every operation targets one anchor. The
Word module ships 17 verbs covering text, tables, images, styles, comments,
document properties, and tracked revisions.

### Documents are registered with a provider, not uploaded

A document is not edited by file path. It lives behind a **document provider**
(storage). The provider is a **registry of references** - it persists only the
path (or URL, drive id, …) the host hands it, not the bytes. You register an
existing document with a connection; the provider returns an **opaque document
id**. Every later call addresses the document by `(connectionId, documentId)`,
and the provider routes reads and saves back to the referenced location.

The agent never sees a file path or credential and cannot leave the storage you
configured. By default it cannot register documents either - the host stages ids
up front; hosts that want the agent to manage its own registrations opt in to
the `register_document` / `remove_document` tools, which stay inside the
connection's boundary and never delete content. Filesystem and SharePoint
providers ship in the box; a database or any other store can implement the same
interface.

## Quick start

Install the packages:

```bash
dotnet add package OfficeAgent.Core
dotnet add package OfficeAgent.Word
dotnet add package OfficeAgent.AgentFramework   # only for the Agent Framework path
dotnet add package OfficeAgent.SharePoint       # only for the SharePoint provider
```

### Over MCP (any agent, any language)

The `OfficeAgent.Mcp` server exposes the same workflow as Model Context Protocol
tools, so any MCP-capable agent can edit real Word documents without taking a
.NET dependency. It runs over stdio for local hosting or streamable HTTP for the
cloud.

```bash
dotnet tool install --global OfficeAgent.Mcp --prerelease
officeagent-mcp --stdio          # local: child process of an MCP client
officeagent-mcp                  # cloud: streamable HTTP + /healthz
```

Point it at your documents through configuration (`appsettings.json` or
`OfficeAgent__`-prefixed environment variables) and the agent gets
`inspect_document`, `find_in_document`, `preview_plan`, `apply_plan`, and -
unless you turn registration off - `register_document` / `remove_document` /
`list_connections`.
See [the MCP server guide](docs/mcp-server.md) for client config snippets,
SharePoint connections, and cloud hosting.

### With Microsoft Agent Framework (MAF)

`OfficeAgent.AgentFramework` exposes the workflow as four core tools a language
model can call: `inspect_document`, `find_in_document`, `preview_plan`, and
`apply_plan`. By default the host registers documents up front and threads the
resulting opaque id into the agent's system prompt; the agent never sees a file
path and cannot register or delete storage on its own. Opt in via
`OfficeAgentToolsOptions.AllowRegistration` to add `register_document` and
`remove_document`, which let the agent stage its own ids inside the configured
connections.

```csharp
using Microsoft.Agents.AI;
using Microsoft.Extensions.AI;
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

// Register the existing document with the connection. The provider stores only
// the path; the host owns the file's lifecycle.
var seeded = await client.RegisterAsync(
    "workspace", "/srv/officeagent/workspace/contract.docx");

var tools  = new OfficeAgentTools(client).AsAIFunctions();
var prompt = $"You are editing documentId={seeded.ItemId} on connectionId=workspace.\n\n"
           + OfficeAgentTools.SystemPromptGuidance;

AIAgent agent = new ChatClientAgent(
    chatClient,                       // any Microsoft.Extensions.AI IChatClient
    instructions: prompt,
    name:         "OfficeAgent",
    description:  "Edits Word documents using OfficeAgent.NET.",
    tools:        tools.Cast<AITool>().ToList(),
    services:     services);
```

`apply_plan` saves the result and returns an `outputDocumentId`. It does not send
`.docx` bytes back through the model. The host reads the id with `OpenReadAsync`
and delivers the file through its own download or attachment API.

A runnable Azure OpenAI example is in [`samples/AgentEdit`](samples/AgentEdit).

### As a library

Drive the same workflow directly from code:

```csharp
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Word;

var services = new ServiceCollection()
    .AddWordFormat()
    .AddFileSystemDocumentProvider("workspace", "/srv/officeagent/workspace")
    .AddOfficeAgent()
    .BuildServiceProvider();

var client = services.GetRequiredService<OfficeAgentClient>();

// Register an existing file with the connection; the provider mints the opaque id.
var doc = await client.RegisterAsync(
    "workspace", "/srv/officeagent/workspace/contract.docx");

// inspect → find → preview → commit, addressing the document by its id.
var inspect = await client.InspectAsync("workspace", doc.ItemId);
var hit     = (await client.FindAsync(
    "workspace", doc.ItemId, new FindQuery("Acme Corp"))).First();

var plan = new DocumentPlan
{
    Snapshot   = inspect.Snapshot,          // opt in to drift detection
    Operations = new PlanOperation[]
    {
        new ChangeTextOp { Target = hit.Anchor, With = "Globex Inc.", Mode = ChangeMode.Tracked }
    }
};

var preview = await client.PreviewAsync("workspace", doc.ItemId, plan);
if (!preview.IsValid) { /* surface preview.Errors */ return; }

// By default a fresh id is minted for the result; the source is preserved.
var result = await client.CommitAsync("workspace", doc.ItemId, plan);
if (result.Committed)
{
    using var saved = await client.OpenReadAsync(result.Document);
    // saved.Stream holds the edited bytes; result.Document.ItemId is the new id.
}
```

A runnable version is in [`samples/QuickEdit`](samples/QuickEdit).

## Documentation

- [Getting started](docs/getting-started.md) - one full edit, step by step.
- [Concepts](docs/concepts.md) - providers, inspect, anchors, snapshots, plans,
  preview/commit, capabilities, transactions.
- [Document providers](docs/document-providers.md) - `IDocumentProvider`,
  registration, save modes, optimistic concurrency, the SharePoint provider.
- [Document plans](docs/document-plans.md) - the JSON contract for every verb.
- [Agent integration](docs/agent-integration.md) - wiring `OfficeAgentTools`
  into Microsoft Agent Framework / MEAI, opt-in registration tools.
- [MCP server](docs/mcp-server.md) - hosting `OfficeAgent.Mcp` locally over
  stdio or in the cloud over streamable HTTP.
- [Deployment & client setup](docs/deployment.md) - connecting the MCP server to
  Claude Code, Codex, Copilot Studio, and Microsoft 365 Copilot, with the
  identity checklist.
- [Operations](docs/operations.md) - thread safety, stream ownership, memory,
  cancellation, telemetry.

## Scope

OfficeAgent.NET ships the Word path today. Excel and PowerPoint can plug in
through the same `IFormatModule` interface but are not implemented yet. Work that
needs a renderer - pagination, field recalculation, table-of-contents rendering,
page-fit checks - is rejected on purpose: the engine can write the OOXML but
cannot compute the displayed value.

## License

MIT. See [LICENSE](LICENSE).
