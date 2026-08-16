# OfficeAgent.NET documentation

OfficeAgent.NET gives agents a structured, validated way to inspect and edit
Word documents and PowerPoint decks. Use this page to choose the shortest path
for your scenario.

> [!IMPORTANT]
> OfficeAgent.NET is pre-1.0. Pin package and container versions in production,
> review release notes before upgrading, and test representative documents.

## Choose a learning path

| Goal | Start here | Then read |
| --- | --- | --- |
| Edit a Word document from C# | [Getting started](getting-started.md) | [Concepts](concepts.md), [document plans](document-plans.md) |
| Add document tools to an agent | [Agent integration](agent-integration.md) | [Document providers](document-providers.md) |
| Connect an MCP client locally | [MCP server](mcp-server.md#local-hosting-stdio) | [Deployment](deployment.md#option-a---local-stdio-for-claude-code-and-codex) |
| Host OfficeAgent for a team | [Deployment](deployment.md#option-b---hosted-http-for-all-four-clients) | [Operations](operations.md), [troubleshooting](troubleshooting.md) |
| Read or write SharePoint files | [SharePoint provider](document-providers.md#the-sharepoint-provider) | [Authentication and identity](deployment.md#authentication--identity) |
| Work with PowerPoint | [PowerPoint support](powerpoint.md) | [Document plans](document-plans.md) |
| Look up an operation's JSON | [Document plans](document-plans.md) | [Validation errors](document-plans.md#validation-errors) |

## Packages

| Package | Use it for | Target frameworks |
| --- | --- | --- |
| `OfficeAgent.Core` | Engine, direct .NET API, filesystem provider | `netstandard2.0`, `net8.0` |
| `OfficeAgent.Word` | Word `.docx` inspection, creation, and editing | `netstandard2.0`, `net8.0` |
| `OfficeAgent.PowerPoint` | PowerPoint `.pptx` inspection, creation, and editing | `netstandard2.0`, `net8.0` |
| `OfficeAgent.AgentFramework` | `Microsoft.Extensions.AI` / Microsoft Agent Framework tools | `netstandard2.0`, `net8.0` |
| `OfficeAgent.SharePoint` | Microsoft Graph document provider | `netstandard2.0`, `net8.0` |
| `OfficeAgent.Mcp` | Standalone MCP server and .NET global tool | `net8.0` |

Applications add `OfficeAgent.Core` plus at least one format module. Add a
provider or agent adapter only when that hosting model needs it. The standalone
MCP tool already includes both format modules and both built-in providers.

## Core workflow

Every integration uses the same safety loop:

1. **Register or create** a document and keep its opaque `(connectionId, documentId)`.
2. **Inspect** structure and capture a snapshot.
3. **Find** text to obtain content-verified anchors.
4. **Preview** the complete plan without writing.
5. **Commit** atomically through the provider.
6. **Open the result** by its returned id and deliver it outside model context.

The engine refuses stale snapshots, mismatched anchors, unsupported operations,
and optimistic-concurrency conflicts instead of guessing.

## Documentation map

### Tutorials

- [Getting started](getting-started.md) — first successful Word edit from C#.
- [Deployment and client setup](deployment.md) — local and hosted MCP recipes.

### Concepts

- [Concepts](concepts.md) — providers, anchors, snapshots, plans, and transactions.
- [PowerPoint support](powerpoint.md) — slide-specific addressing and behavior.

### How-to guides

- [Agent integration](agent-integration.md) — expose bounded tools and deliver output.
- [Document providers](document-providers.md) — filesystem and SharePoint storage.
- [Operations](operations.md) — concurrency, memory, telemetry, and production operation.
- [Troubleshooting](troubleshooting.md) — diagnose setup, identity, and edit failures.

### Reference

- [Document plans](document-plans.md) — operation JSON and validation codes.
- [MCP server](mcp-server.md) — transports, settings, tools, and response contracts.

## Scope

OfficeAgent edits OOXML packages; it does not automate desktop Office or render
pages and slides. Excel, Word pagination, field calculation, and visual page-fit
validation are outside the current scope. See the
[project limitations](../README.md#scope-and-limitations) before production use.

## Get help

Start with [Troubleshooting](troubleshooting.md). If the problem remains, open a
[GitHub issue](https://github.com/ilia-sokolov/OfficeAgent.NET/issues) with the
package version, provider, document format, operation, stable error code, and a
small sanitized reproduction. Do not attach confidential documents, access
tokens, registration indexes, or tenant identifiers.
