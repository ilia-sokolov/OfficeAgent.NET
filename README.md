# OfficeAgent.NET
<!-- mcp-name: io.github.ilia-sokolov/officeagent -->

[![build](https://img.shields.io/github/actions/workflow/status/ilia-sokolov/OfficeAgent.NET/build.yml?branch=main)](https://github.com/ilia-sokolov/OfficeAgent.NET/actions/workflows/build.yml)
[![NuGet](https://img.shields.io/nuget/v/OfficeAgent.Core.svg)](https://www.nuget.org/packages/OfficeAgent.Core)
[![downloads](https://img.shields.io/nuget/dt/OfficeAgent.Core.svg)](https://www.nuget.org/packages/OfficeAgent.Core)
[![license](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

OfficeAgent.NET translates an AI agent’s intent into controlled changes to Microsoft Word documents and PowerPoint decks. The agent proposes a typed edit plan; the library validates and applies it while preserving document features such as styles and comments. Word edits can be recorded as tracked changes for human review, while structured document operations can reduce token use compared with processing entire files.


![OfficeAgent.NET finds, previews, and applies a contract edit as a tracked change in Word.](https://raw.githubusercontent.com/ilia-sokolov/OfficeAgent.NET/main/media/demo.gif)

## What this project does

A `.docx` or `.pptx` file is a package of related XML parts. A small text change
can affect runs, styles, numbering, comments, content controls, or revision
markup. OfficeAgent.NET handles that document-specific work. The model works
with structured document data and JSON-serialisable operations such as "replace
this clause as a tracked change" or "add a row to this table."

The same engine is available in three forms:

- an MCP server for agents that support the Model Context Protocol;
- tools for Microsoft Agent Framework and `Microsoft.Extensions.AI`;
- a .NET API for applications that want to control the workflow directly.

It supports Word `.docx` files and PowerPoint `.pptx` decks; one client serves
both, routing each document to the module that handles it. Excel is not
implemented. See [Scope and limitations](#scope-and-limitations) before choosing
it for a workflow that depends on Office's layout or calculation engine.

## Choose a starting point

| I want to... | Start here |
| --- | --- |
| Edit my first Word document | [Your first edit](#your-first-edit) |
| Connect Codex, Claude Code, Copilot Studio, or Microsoft 365 Copilot | [Deployment and client setup](docs/deployment.md) |
| Use OfficeAgent from C# | [Getting started](docs/getting-started.md) |
| Add tools to a Microsoft Agent Framework agent | [Agent integration](docs/agent-integration.md) |
| Host the MCP server or use SharePoint | [MCP server](docs/mcp-server.md) and [document providers](docs/document-providers.md) |
| Edit documents with no storage configured | [Documents with no storage](docs/mcp-server.md#documents-with-no-storage) |
| Teach an agent to review documents properly | [word-document-review skill](skills/word-document-review/SKILL.md) |
| Contribute | [Contributing](#contributing) |

## Your first edit

Install the server:

```bash
dotnet tool install --global OfficeAgent.Mcp
```

Make a folder for the agent to work in and download the
[sample contract](samples/documents/services-agreement.docx) into it — a fictional services
agreement with a clause to change, a table, an open comment, and a pending redline:

```bash
mkdir -p ~/officeagent-documents
curl -Lo ~/officeagent-documents/services-agreement.docx \
  https://raw.githubusercontent.com/ilia-sokolov/OfficeAgent.NET/main/samples/documents/services-agreement.docx
```

Any `.docx` of your own works too — the sample just gives you something with a comment and a
pending revision already in it.

Register the server with Claude Code, pointed at that folder and nothing else:

```bash
claude mcp add \
  --env OfficeAgent__FileSystemConnections__0__ConnectionId=documents \
  --env OfficeAgent__FileSystemConnections__0__RootPath=$HOME/officeagent-documents \
  --transport stdio \
  officeagent -- officeagent-mcp --stdio
```

Then ask:

> In services-agreement.docx, change the payment terms from thirty days to forty-five days.

Open the file in Word. Clause 3 now reads **forty-five days** as a tracked change you can
accept or reject, and everything else — the table, the comment, the redline that was
already there — is exactly as it was. That is the whole idea: the document that comes out
is the one that went in, minus the edit you asked for.

[What else the sample is good for](samples/documents/README.md) — reviewing comments,
accepting revisions, editing the table.

### If it does not work

| | |
| --- | --- |
| `claude mcp list` shows officeagent as failed | Check `RootPath` is an absolute path to a directory that exists. |
| The agent says it cannot find the document | The name must be relative to `RootPath`, not a full path. |
| `io-error` on save | The document is open in Word. Close it. |

## Beyond the first edit

The quick start above is deliberately the smallest thing that works. Four settings extend it:

| Setting | Adds |
| --- | --- |
| `OfficeAgent__AllowCreation=true` | `create_document`, so "draft a project brief in brief.docx" makes a new file instead of failing |
| `OfficeAgent__FileSystemConnections__0__AllowedExtensions__1=.pptx` | PowerPoint decks. Add `__DefaultChangeMode=Direct` with it — a deck has no redline vocabulary and refuses tracked changes |
| `OfficeAgent__EphemeralConnectionId=session` | Names the in-memory session connection explicitly. With no configuration at all the server already falls back to one - this is for running it alongside storage, or under a different id |
| `OfficeAgent__AllowInlineContent=true` | Tools that carry the document as base64, for a single self-contained call |

Past a couple of settings, use a file instead — the same `OfficeAgent` section, where a list
is a list:

```json
{
  "OfficeAgent": {
    "AllowCreation": true,
    "FileSystemConnections": [
      {
        "ConnectionId": "documents",
        "RootPath": "C:\\officeagent-documents",
        "AllowedExtensions": [ ".docx", ".pptx" ],
        "DefaultChangeMode": "Direct"
      }
    ]
  }
}
```

```bash
claude mcp add --transport stdio officeagent -- officeagent-mcp --stdio --config ./officeagent.json
```

Environment variables still override the file. Windows, PowerShell, other MCP clients,
HTTP hosting and SharePoint are in
[Deployment and client setup](docs/deployment.md); every setting is listed in
[MCP server](docs/mcp-server.md).

### Teaching the agent to review, not just replace

[`skills/word-document-review`](skills/word-document-review/SKILL.md) is an
[Agent Skill](https://code.claude.com/docs/en/skills) that teaches the review loop: read the
comments and pending revisions before editing, keep edits as redlines, address documents by
id rather than passing bytes around, and recover from each error code rather than retrying.
Copy it into `.claude/skills/` in your project, or `~/.claude/skills/` for every project:

```bash
cp -r skills/word-document-review ~/.claude/skills/
```

### What reaches the model

The inspect and find tools return document text and structure to the model — that is how it
locates an edit. The `.docx` package itself does not travel that way *unless* you enable
`AllowInlineContent`, whose tools carry the whole file as base64 in both directions by
design. Connect folders and model providers appropriate for the data you are handling.

The server ships no authentication layer for HTTP hosting; put it behind your own. A
filesystem root is a trust boundary: its ACLs must stop untrusted principals creating,
renaming or replacing entries while the server runs.

## .NET quick start

Install the core package and Word module:

```bash
dotnet add package OfficeAgent.Core
dotnet add package OfficeAgent.Word
```

After registering services and a document provider, the edit loop looks like
this:

```csharp
var client = services.GetRequiredService<OfficeAgentClient>();
var doc = await client.RegisterAsync("workspace", "/srv/workspace/contract.docx");

var inspect = await client.InspectAsync("workspace", doc.ItemId);
var hit = (await client.FindAsync(
    "workspace", doc.ItemId, new FindQuery("Acme Corp"))).First();

var plan = new DocumentPlan
{
    Snapshot = inspect.Snapshot,
    Operations = new PlanOperation[]
    {
        new ChangeTextOp
        {
            Target = hit.Anchor,
            With = "Globex Inc.",
            Mode = ChangeMode.Tracked
        }
    }
};

var preview = await client.PreviewAsync("workspace", doc.ItemId, plan);
if (preview.IsValid)
    await client.CommitAsync("workspace", doc.ItemId, plan);
```

The complete example, including service registration and reading the saved
file, is in [Getting started](docs/getting-started.md).
The minimal sample replaces the first `Acme Corp` with `Globex Inc.`. To run it,
copy a Word document containing `Acme Corp` to `contract.docx` in the cloned
repository root, then run:

```bash
dotnet run --project samples/QuickEdit -- ./contract.docx ./contract-edited.docx
```

The repository also contains a
[direct `IChatClient` Word-editing sample](samples/IChatClientWordEdit/) and an
interactive
[Agent Framework sample](samples/AgentEdit/).

## How it works

Every edit follows the same four steps:

1. **Inspect** returns a structured map of the document: its outline,
   paragraphs, styles, content controls, tables, images, and revisions.
2. **Find** searches text and returns a content-verified anchor for each match.
3. **Preview** validates a plan against the current document and reports the
   proposed changes without writing.
4. **Apply** commits the complete plan and saves it through the configured
   provider.

A plan (`DocumentPlan`) is a typed, JSON-serialisable list of operations. An
anchor records both a location and the content expected there. If the content
or optional document snapshot has changed, validation fails instead of silently
targeting a different location. Applying a plan is all-or-nothing.

The Word module supports changes to text, paragraphs, tables, images, styles,
content controls, comment threads, footnotes and endnotes, page geometry and
breaks, document properties, and tracked revisions. Every verb that changes
content records a redline when the connection asks for one - an inserted clause,
a deleted row and a restyled heading all come back as revisions a reviewer
accepts or rejects, not only a replaced phrase. The
PowerPoint module implements a broad, explicitly documented set of deck
operations: text, bullets, run and paragraph formatting, template
slots, style copying, tables, images, text boxes, embedded video and audio,
speaker notes, resolvable comments, footers and slide numbers, sections,
transitions and animations, and the slide lifecycle - adding, removing,
reordering and duplicating. Several slide inserts in one plan author a deck end
to end, so a single call turns nothing into a finished presentation. Any verb it
does not support is named rather than silently skipped. The full operation
schema is documented in [Document plans](docs/document-plans.md), and the deck
specifics in [PowerPoint support](docs/powerpoint.md).

Documents are accessed through configured providers. After registration,
editing calls use a `(connectionId, documentId)` pair instead of a storage path
or credentials. The filesystem provider restricts registrations to its root;
the SharePoint provider uses the permissions of its configured identity.
`CreateAsync` starts a new document inside a connection: the requested `.docx`
or `.pptx` extension selects a registered blank-document factory. The engine
applies an optional initial plan in memory, and then asks
the provider to create and register it without overwriting an existing name.

## Documentation

| Guide | Covers |
| --- | --- |
| [Documentation hub](docs/README.md) | Learning paths, package map, and the complete documentation set |
| [Getting started](docs/getting-started.md) | A complete edit from service registration to reading the result |
| [Concepts](docs/concepts.md) | Anchors, snapshots, plans, providers, transactions, and capabilities |
| [Document plans](docs/document-plans.md) | JSON shapes and validation rules for every operation |
| [Document providers](docs/document-providers.md) | Filesystem, SharePoint, save modes, and custom providers |
| [PowerPoint support](docs/powerpoint.md) | Slide addressing, the verbs the deck module implements, and what it preserves |
| [Agent integration](docs/agent-integration.md) | Microsoft Agent Framework and `Microsoft.Extensions.AI` tools |
| [MCP server](docs/mcp-server.md) | Server configuration, transports, security notes, and tool contracts |
| [Deployment and client setup](docs/deployment.md) | Codex, Claude Code, Microsoft Copilot clients, containers, and Azure |
| [Operations](docs/operations.md) | Concurrency, streams, cancellation, telemetry, and production concerns |
| [Troubleshooting](docs/troubleshooting.md) | Startup, registration, validation, concurrency, and provider failures |
| [Failure modes](docs/operations.md#failure-modes-you-should-handle) | Common plan errors and what to do next |
| [Releasing](docs/releasing.md) | Publishing to NuGet, the MCP Registry, and GitHub |

## Contributing

Bug reports, documentation fixes, new document operations, provider
integrations, and focused test cases are useful contributions. If you found a problem,
[open an issue](https://github.com/ilia-sokolov/OfficeAgent.NET/issues) with the
document feature involved, the operation you attempted, and the error or
unexpected result. Do not attach confidential documents; a small sanitised
reproduction is enough.

To work on the code, install the .NET 8 SDK, fork the repository, and run:

```bash
dotnet build OfficeAgent.NET.sln
dotnet test OfficeAgent.NET.sln
```

Before starting a larger change, especially one that changes public types or
the JSON wire format, [open an issue](https://github.com/ilia-sokolov/OfficeAgent.NET/issues)
so the design can be discussed. See
[CONTRIBUTING.md](CONTRIBUTING.md)
for code style, tests, and pull-request expectations.

## Scope and limitations

OfficeAgent.NET edits Word `.docx` files and PowerPoint `.pptx` decks; it does
not automate the Office desktop applications. An Excel module can be added
through `IFormatModule`, but it does not ship today.

The deck module refuses the verbs a presentation has no vocabulary for -
`setProperty`, `revision`, `pageSetup`, `insertBreak` and `note` - per
operation, rather than applying part of a plan, and refuses an explicit tracked
mode on any verb that carries one. PresentationML has no redline model, so tracked changes are Word-only, and
a slide has no header (that is a notes and handout concept). Animations cover
the effects expressible as a filtered `p:animEffect`; fly-in, zoom and motion
paths are refused rather than approximated. See
[PowerPoint support](docs/powerpoint.md) for what a deck does and does not
accept.

The engine does not render pages or calculate Word fields. Operations that
depend on pagination, table-of-contents rendering, field recalculation, or
page-fit checks are outside its scope. Preview reports structural changes, not
a visual rendering of the final document. Test the workflow on representative
documents and keep human review in the loop for consequential edits.

Two more limits worth knowing before you build on it:

- **Token savings depend on how you connect.** Addressing a document by id keeps
  the package out of the conversation, and inspection can be narrowed with
  `fidelity` and paging - that is where the saving comes from. The inline
  `*_content` tools are the deliberate exception: they carry the whole file as
  base64 in both directions, which costs tokens in proportion to file size. They
  suit a single self-contained call, not a sequence of edits - a model asked to
  pass a document of a few kilobytes back for a second edit reproduces it
  imperfectly and the follow-up fails. Use a connection, or a session connection,
  when more than one edit is coming.
- **A skill helps, and is not automatic.** Nothing here makes an agent read the
  open comments before editing, or keep an edit as a redline. The
  [word-document-review skill](skills/word-document-review/SKILL.md) teaches that;
  without it, behaviour depends on the model and the prompt.

## Commercial support

OfficeAgent.NET is MIT-licensed and can be self-hosted. Managed hosting and
commercial support are available from dotaction:
[contact dotaction](mailto:contact@dotaction.io?subject=OfficeAgent.NET%20commercial%20support).

## License

MIT. See [LICENSE](LICENSE).
