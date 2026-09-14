# OfficeAgent.NET
<!-- mcp-name: io.github.ilia-sokolov/officeagent -->

[![build](https://img.shields.io/github/actions/workflow/status/ilia-sokolov/OfficeAgent.NET/build.yml?branch=main)](https://github.com/ilia-sokolov/OfficeAgent.NET/actions/workflows/build.yml)
[![NuGet](https://img.shields.io/nuget/v/OfficeAgent.Core.svg)](https://www.nuget.org/packages/OfficeAgent.Core)
[![downloads](https://img.shields.io/nuget/dt/OfficeAgent.Core.svg)](https://www.nuget.org/packages/OfficeAgent.Core)
[![license](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

OfficeAgent.NET is a .NET document-automation library built on the
[Open XML SDK](https://learn.microsoft.com/office/open-xml/open-xml-sdk). Its direct
.NET API turns document intent into typed, validated operations for Word `.docx`,
PowerPoint `.pptx`, and Excel `.xlsx` packages.

Use it to generate documents and presentations, make targeted edits, update tables,
styles, and images, or manage comments and review state. An MCP server and adapters
for Microsoft Agent Framework and `Microsoft.Extensions.AI` are optional interfaces
to the same engine. Applications can use the engine directly without MCP.

One example is a targeted Word edit whose result remains reviewable:

![OfficeAgent.NET finds, previews, and applies a contract edit as a tracked change in Word.](https://raw.githubusercontent.com/ilia-sokolov/OfficeAgent.NET/main/media/demo.gif)

## What this project does

An Office Open XML file is a package of related XML parts. A small change
can affect runs, styles, numbering, comments, content controls, or revision
markup. OfficeAgent.NET handles that document-specific work. The model works
with structured document data and JSON-serialisable operations such as "replace
this clause as a tracked change" or "add a row to this table."

The same engine is available in three forms:

- an MCP server for agents that support the Model Context Protocol;
- tools for Microsoft Agent Framework and `Microsoft.Extensions.AI`;
- a .NET API for applications that want to control the workflow directly.

It supports Word `.docx`, PowerPoint `.pptx`, and Excel `.xlsx`; one client routes
each document to the module that handles it. See [Scope and limitations](#scope-and-limitations) before choosing
it for a workflow that depends on Office's layout or calculation engine.

### What you can build

| Area | Supported workflows |
| --- | --- |
| Word creation and editing | Create `.docx` files; inspect and change text, paragraphs, tables, images, styles, content controls, headers, footers, notes, page setup, and document properties |
| Word review | Read and manage comments, preserve or resolve review state, set one revision identity per plan, and record supported edits as tracked revisions |
| PowerPoint creation and editing | Build or update decks with slides, layouts, text, tables, native editable charts, images, media, notes, comments, sections, transitions, and animations |
| Excel inspection and editing | Inspect worksheets, tables, and bounded ranges; find raw or displayed values; set cells and formulas; append table rows; manage cell notes |
| Template generation | Bind unique Word content-control tags or PowerPoint shape names, expand repeating Word table rows, and create bounded batches with one receipt per output |
| Word comparison | Compare supported free-body paragraph text read-only and produce a snapshot-bound native redline plan only when all other package content is unchanged |
| Agent and application integration | Use MCP over stdio or HTTP, Microsoft Agent Framework tools, or the direct .NET API, with SHA-256 apply receipts and host-supplied audit actors |
| Document access | Work with bounded filesystem roots, SharePoint, in-memory sessions, or self-contained inline content |

## Choose a starting point

| I want to... | Start here |
| --- | --- |
| Decide whether OfficeAgent fits my application | [Library selection guide](docs/choose-officeagent.md) |
| Try a targeted Word edit | [Try a Word edit](#try-a-word-edit) |
| Create a Word document from scratch | [Create a document](docs/getting-started.md#create-a-document-instead) |
| Create or edit a PowerPoint deck | [PowerPoint support](docs/powerpoint.md) |
| Inspect or edit an Excel workbook | [Excel support](docs/excel.md) |
| Connect Codex, Claude Code, Copilot Studio, or Microsoft 365 Copilot | [Deployment and client setup](docs/deployment.md) |
| Use OfficeAgent from C# | [Getting started](docs/getting-started.md) |
| Add tools to a Microsoft Agent Framework agent | [Agent integration](docs/agent-integration.md) |
| Host the MCP server or use SharePoint | [MCP server](docs/mcp-server.md) and [document providers](docs/document-providers.md) |
| Add per-user hosted connection authorization | [Hosted gateway reference](samples/HostedGateway/) |
| Add optional PDF/page-image rendering | [Visual rendering](docs/rendering.md) |
| Edit documents with no storage configured | [Documents with no storage](docs/mcp-server.md#documents-with-no-storage) |
| Run a tracked-review workflow | [Optional word-document-review skill](skills/word-document-review/SKILL.md) |
| Build a contract-review agent | [ContractReview sample](samples/ContractReview/) |
| Populate quote templates or compare Word documents | [Template and comparison workflows](docs/document-workflows.md) |
| Check support, compatibility, or security policy | [Support](SUPPORT.md) and [security](SECURITY.md) |
| Contribute | [Contributing](#contributing) |

## Try a Word edit

This small workflow demonstrates that OfficeAgent can change an existing OOXML file
without flattening its structure. It uses tracked changes because the result is easy to
verify in Word; review is one part of the broader document operation set.

Install the server. The published package command is:

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

PowerShell:

```powershell
$officeAgentDocuments = Join-Path $env:USERPROFILE "officeagent-documents"
New-Item -ItemType Directory -Force $officeAgentDocuments | Out-Null
Invoke-WebRequest `
  https://raw.githubusercontent.com/ilia-sokolov/OfficeAgent.NET/main/samples/documents/services-agreement.docx `
  -OutFile (Join-Path $officeAgentDocuments "services-agreement.docx")
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

PowerShell:

```powershell
claude mcp add `
  --env OfficeAgent__FileSystemConnections__0__ConnectionId=documents `
  --env "OfficeAgent__FileSystemConnections__0__RootPath=$officeAgentDocuments" `
  --transport stdio `
  officeagent -- officeagent-mcp --stdio
```

For this review-specific workflow, you can optionally install the
[word-document-review skill](docs/skill-installation.md) before starting the client.

Then ask:

> In services-agreement.docx, change the payment terms from thirty days to forty-five days.

Open the file in Word. Clause 3 now reads **forty-five days** as a tracked change you can
accept or reject, and everything else — the table, the comment, the redline that was
already there — is exactly as it was. This demonstrates a key engine property: apply the
requested operation while preserving unrelated package content.

[What else the sample is good for](samples/documents/README.md) — reviewing comments,
accepting revisions, editing the table.

Next, try [creating a Word document](docs/getting-started.md#create-a-document-instead),
[generating a PowerPoint deck](docs/powerpoint.md#generating-a-deck), or using the
[direct .NET workflow](docs/getting-started.md).

### If it does not work

| | |
| --- | --- |
| `claude mcp list` shows officeagent as failed | Check `RootPath` is an absolute path to a directory that exists. |
| The agent says it cannot find the document | Use a relative name, or an absolute path that still resolves inside `RootPath`. |
| `io-error` on save | Close the file in Word, then check filesystem permissions and the available disk space. |

## Configure broader workflows

The quick start above is deliberately the smallest thing that works. Four settings extend it:

| Setting | Adds |
| --- | --- |
| `OfficeAgent__AllowCreation=true` | `create_document`, so "draft a project brief in brief.docx" makes a new file instead of failing |
| `OfficeAgent__FileSystemConnections__0__AllowedExtensions__0=.docx` plus `OfficeAgent__FileSystemConnections__0__AllowedExtensions__1=.pptx` | Word and PowerPoint on one connection. Declaring this list replaces the `.docx` default. Set `OfficeAgent__FileSystemConnections__0__DefaultChangeMode=Direct` for decks, and send `"mode": "Tracked"` explicitly for reviewable Word edits on that mixed connection. |
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

The same configuration is available as
[`samples/config/word-and-powerpoint.json`](samples/config/word-and-powerpoint.json). Change
`RootPath` before using it.

```bash
claude mcp add --transport stdio officeagent -- officeagent-mcp --stdio --config ./officeagent.json
```

Environment variables still override the file. Windows, PowerShell, other MCP clients,
HTTP hosting and SharePoint are in
[Deployment and client setup](docs/deployment.md); every setting is listed in
[MCP server](docs/mcp-server.md).

### Optional guidance for Word review

[`skills/word-document-review`](skills/word-document-review/SKILL.md) teaches the review
loop: read comments and pending revisions before editing, keep reviewable Word edits as
redlines, use document ids for multi-step work, and recover from stable error codes. The
[installation guide](docs/skill-installation.md) gives complete Bash and PowerShell steps
for Claude Code and Codex, including installation from a fresh machine and verification.
The skill is only needed when the task requires that review discipline; document creation,
ordinary direct edits, and PowerPoint workflows use the server without it.

### What reaches the model

The inspect and find tools return document text and structure to the model — that is how it
locates an edit. Filesystem and SharePoint operations keep the package behind an opaque id.
Inline tools carry the whole file as base64 on every call. Session import/export also carries
the package as base64 if the agent performs those calls; a host integration can instead move
the bytes outside model context. Connect storage and model providers appropriate for the data.

The standalone server ships no authentication layer for HTTP hosting; put it behind your own,
or start from the authenticated [HostedGateway reference](samples/HostedGateway/). A
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
file, is in [Getting started](docs/getting-started.md). The minimal direct-.NET
sample runs against the bundled fictional contract, so it needs no MCP client,
language model, or document of your own:

```bash
dotnet run --project samples/QuickEdit -- \
  samples/documents/services-agreement.docx quickedit-output.docx
```

Open `quickedit-output.docx` in Word and verify that the payment term is a tracked
change while the existing revision, comment, table, and headings remain intact.
[QuickEdit](samples/QuickEdit/) also accepts an exact source and replacement text
for your own document.

The repository also contains a
[direct `IChatClient` Word-editing sample](samples/IChatClientWordEdit/) and an
interactive
[Agent Framework sample](samples/AgentEdit/), plus a complete
[contract-review agent](samples/ContractReview/) that separates model judgement from
validated document writes. The [TemplateBatch](samples/TemplateBatch/) sample generates two
quotes from one tagged template, while [DocumentComparison](samples/DocumentComparison/)
turns covered body-paragraph differences into a reviewable Word redline.
The [DocumentAssembly](samples/DocumentAssembly/) sample combines a proposal, statement of
work, and appendix into one editable package with a multi-source audit receipt. See
[Word document assembly](docs/document-assembly.md) for its formatting and compatibility scope.

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
breaks, document properties, and tracked revisions. Operations with a Word revision
representation record a redline when the connection asks for one - an inserted clause,
a deleted row and a restyled heading all come back as revisions a reviewer
accepts or rejects, not only a replaced phrase. Image resizing is applied directly because
WordprocessingML has no revision representation for drawing dimensions. The
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
| [Library selection guide](docs/choose-officeagent.md) | Supported jobs, non-goals, package choices, alternatives, and a verified direct .NET recipe |
| [Getting started](docs/getting-started.md) | A complete edit from service registration to reading the result |
| [Concepts](docs/concepts.md) | Anchors, snapshots, plans, providers, transactions, and capabilities |
| [Document plans](docs/document-plans.md) | JSON shapes and validation rules for every operation |
| [Document providers](docs/document-providers.md) | Filesystem, SharePoint, save modes, and custom providers |
| [PowerPoint support](docs/powerpoint.md) | Slide addressing, the verbs the deck module implements, and what it preserves |
| [Template population and comparison](docs/document-workflows.md) | Batch binding, repeating Word rows, comparison limits, and redline generation |
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

OfficeAgent.NET edits Word `.docx`, PowerPoint `.pptx`, and Excel `.xlsx` files;
it does not automate the Office desktop applications.

The deck module refuses the verbs a presentation has no vocabulary for -
`setProperty`, `revision`, `pageSetup`, `insertBreak` and `note` - per
operation, rather than applying part of a plan, and refuses an explicit tracked
mode on any verb that carries one. PresentationML has no redline model, so tracked changes are Word-only, and
a slide has no header (that is a notes and handout concept). Animations cover
the effects expressible as a filtered `p:animEffect`; fly-in, zoom and motion
paths are refused rather than approximated. See
[PowerPoint support](docs/powerpoint.md) for what a deck does and does not
accept.

The core engine does not render pages, calculate Word fields, or evaluate Excel formulas.
Formula edits set the workbook to recalculate when Excel opens it. Operations that depend on
pagination, table-of-contents rendering, or field recalculation are outside the core scope.
Preview reports structural changes. The optional [rendering package](docs/rendering.md) can
produce PDF-derived page images through external processes, but it does not yet detect overflow
or page-fit problems. Test the workflow on representative documents and keep human review in the
loop for consequential edits.

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
- **Review guidance is optional.** For review tasks, the server alone does not make an
  agent read open comments before editing or choose a redline. The
  [word-document-review skill](skills/word-document-review/SKILL.md) teaches that workflow;
  without it, review behaviour depends on the model and the prompt.

## Commercial support

OfficeAgent.NET is MIT-licensed and can be self-hosted. Commercial support and
deployment assistance are available from dotaction:
[contact dotaction](mailto:contact@dotaction.io?subject=OfficeAgent.NET%20commercial%20support).

## License

MIT. See [LICENSE](LICENSE).
