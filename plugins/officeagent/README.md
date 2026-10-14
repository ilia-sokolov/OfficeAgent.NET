# OfficeAgent

Inspect and edit existing Word, PowerPoint and Excel files from Claude. OfficeAgent works on
the real `.docx`, `.pptx` and `.xlsx` package rather than rewriting the file: Claude finds the
text it wants to change, previews a typed edit plan, and OfficeAgent applies it all or nothing.
Word edits are written as real tracked changes by default, so you accept or reject them in
Word, and everything the edit did not touch stays as it was.

## What's included

- **The OfficeAgent MCP server**, which gives Claude tools to inspect documents, find text,
  preview and apply edit plans, compare two Word documents, and fill templates.
- **The `word-document-review` skill**, which teaches Claude to review a Word document the way
  a person does: read the open comments first, redline with tracked changes, and resolve
  review state on request.
- **The `officeagent-integration` skill**, for developers who want to use OfficeAgent.NET from
  their own .NET code or agent framework.

## Use it

Ask Claude about a document in your project, for example:

- "Replace Acme Corp with Globex Inc. in contract.docx as tracked changes."
- "Add a row for March to the table in budget.xlsx."
- "Read the open comments in the proposal and address each one."

Claude previews each change before saving it. The server can reach only the `.docx`, `.pptx`
and `.xlsx` files inside your project folder; nothing outside it is reachable.

Two options are off until you turn them on in the plugin's settings:

- **Allow creating new documents** lets Claude create new files in the project folder.
- **Calculate formulas in workbooks** saves the results of formulas Claude writes, for the
  functions OfficeAgent can calculate.

## Requirements

The server is the `officeagent-mcp` package from PyPI, started with
[uv](https://docs.astral.sh/uv/)'s `uvx` command, so install uv first. The package carries a
self-contained server executable; neither .NET nor Microsoft Office is needed. It runs on
Windows (x64, Arm64), macOS 15 or later, and Linux with glibc 2.27 or later.

The server is a local process, so it runs in Claude Code and in Cowork sessions on your
computer. In claude.ai chat the skills load without it.

## What it runs and what data it sends

- On first use, `uvx` downloads `officeagent-mcp` version 1.2.0 from PyPI, pinned to that exact
  version, and runs it on your computer.
- The server reads and writes documents only inside your project folder. To give each
  document an id, it keeps a small index at `.officeagent/index.json` in that folder.
- It sends nothing anywhere. It has no telemetry, analytics or update checks, and it never
  contacts the OfficeAgent project.
- The document text and change previews it returns go to Claude, as part of your conversation.

## Privacy Policy

OfficeAgent collects no data and contacts no third party. The full policy, including what is
stored on your computer and for how long, is at
https://github.com/ilia-sokolov/OfficeAgent.NET/blob/main/PRIVACY.md.

## Support and source

- Source, documentation and releases: https://github.com/ilia-sokolov/OfficeAgent.NET
- Issues: https://github.com/ilia-sokolov/OfficeAgent.NET/issues

MIT licensed; see LICENSE.
