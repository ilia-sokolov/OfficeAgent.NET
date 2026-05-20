# OfficeAgent.NET

Agent-friendly .NET building blocks for producing real Microsoft Office documents from structured AI output.

OfficeAgent.NET is an early-stage open-source toolkit for .NET developers building AI agents that need to create, fill, or revise Word documents without dropping down into hundreds of lines of low-level Open XML code.

The goal is simple: let an agent describe what should change in a document, and let a reliable .NET layer handle the Office file-format details.

## Why This Exists

AI agents are good at producing text, Markdown, JSON, and tool calls. Professional workflows often need something more specific: a `.docx` file that opens cleanly in Word, uses the organization's real template, preserves styles and numbering, fills named fields, and survives review with comments and tracked changes.

Today, .NET developers usually have to choose between:

- writing verbose `DocumentFormat.OpenXml` code directly;
- converting Markdown or HTML into a rough Word approximation;
- buying a general-purpose commercial document engine;
- building a private helper library that becomes its own mini-framework.

OfficeAgent.NET sits in the gap between structured agent output and Office documents. The agent says what it wants. The library translates that plan into valid Office document changes.

The intended workflow is **inspect -> plan -> apply**: the library describes a Word document to the agent in structured, anchored terms; the agent returns a typed change plan; the library validates and applies that plan while preserving the underlying `.docx` package.

## Intended v0.1 Scope

The first target is Word document generation and revision through a .NET package.

Planned primitives:

| Primitive | What it does |
|---|---|
| **Inspect** | Return an anchored map of a `.docx` document that an agent can reason about |
| **Template fill** | Populate named slots in an existing `.docx` template |
| **Content controls** | Write into Word content controls without disturbing surrounding styles |
| **Structured tables** | Turn typed data into clean Word tables |
| **Text changes** | Change anchored text spans without blind global replacement |
| **Tracked changes** | Express agent-proposed edits as Word redlines |
| **Comments** | Attach review comments to specific document locations |

Explicitly out of v0.1:

- MCP tools for document-editing agents.
- Excel workbook primitives.
- PowerPoint presentation primitives.
- Higher-level recipes for reports, proposals, review packs, and contract drafts.

## Current Status

OfficeAgent.NET is in problem-validation and design.

There is not yet:

- a NuGet package;
- a stable public API;
- production-ready code.

## Roadmap

v0.1 target:

- open and inspect an existing `.docx`;
- fill content controls;
- insert a structured table;
- produce at least one tracked-change suggestion;
- save a Word document that opens cleanly and preserves the template.

v0.2 target:

- expose the Word workflow through MCP tools for agents that revise documents directly.

Future versions:

- Excel workbook primitives for agent-produced business artifacts;
- PowerPoint primitives for slide and placeholder workflows;
- higher-level document recipes once the core Word workflow is proven.

## Contributing

The project is too early for broad code contributions, but feedback is very welcome.

Good first ways to help:

- star the repo if the problem statement matches your experience;
- open an issue describing a real Word automation problem you hit with an AI agent;
- share a minimal `.docx` scenario that broke your pipeline;
- point to existing .NET libraries or examples that solve part of this well;
- comment on the planned v0.1 primitives.
