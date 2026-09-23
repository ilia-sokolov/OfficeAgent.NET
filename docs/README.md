# OfficeAgent.NET documentation

OfficeAgent.NET gives agents a structured, validated way to create, inspect, and edit
Word documents, PowerPoint decks, and Excel workbooks. It supports document generation, targeted changes,
rich content operations, review workflows, and several hosting and storage models. Use
this page to choose the shortest path for your scenario.

> [!IMPORTANT]
> This documentation targets OfficeAgent.NET 1.0.0. Pin package and container versions in
> production, review release notes before upgrading, and test representative documents.

## Choose by role

| You are... | Start with | Use as the contract |
| --- | --- | --- |
| Evaluating fit, evidence, and limits | [Library selection guide](choose-officeagent.md) | [Scope](#scope), [support](../SUPPORT.md), and the named evidence pages—not marketing examples |
| A .NET application developer | [Getting started](getting-started.md) | Public C# API plus [document plans](document-plans.md) |
| An agent or MCP integrator | [Agent integration](agent-integration.md) | Runtime tool schemas, tool descriptions, and the matching versioned skill |
| A production operator | [Deployment](deployment.md) | Provider, security, ingestion-limit, operation, and [recovery](recovery.md) contracts |
| A maintainer or release verifier | [Releasing](releasing.md) | Generated baselines, candidate verifiers, and explicitly versioned evidence corpora |

Generated [wire-contract](wire-contract.md) schemas define request and response structure but
deliberately exclude model-facing descriptions. Agents must use the runtime schemas and current
skills together; humans should use the task guides and public API reference. Historical `v0.9.0`
corpora and evaluations remain evidence for that release, not instructions for a 1.0 integration.

## Choose a task

| Goal | Start here | Then read |
| --- | --- | --- |
| Decide whether OfficeAgent fits | [Library selection guide](choose-officeagent.md) | [Support and compatibility](../SUPPORT.md) |
| Combine OfficeAgent with direct Open XML SDK edits | [SDK interoperability](sdk-interoperability.md) | [Concepts](concepts.md), [operations](operations.md) |
| Inspect Word preservation evidence | [Word preservation matrix](word-preservation-evidence.md) | [Corpus and rerun recipe](../tests/OfficeAgent.Tests/Corpus/v0.9.0/word-preservation/README.md) |
| Check what Office itself sees in OfficeAgent output | [Native Office compatibility](native-compatibility.md) | [Native corpus](../tests/OfficeAgent.Tests/Corpus/v1.0.0/native/README.md) |
| Edit a Word document from C# | [Getting started](getting-started.md) | [Concepts](concepts.md), [document plans](document-plans.md) |
| Look up a public C# type or member | [Generated C# API reference](csharp-api.md) | [Getting started](getting-started.md), [document plans](document-plans.md) |
| Create a Word document from scratch | [Create a document](getting-started.md#create-a-document-instead) | [Document plans](document-plans.md) |
| Create or update a PowerPoint deck | [PowerPoint support](powerpoint.md#creating-a-deck) | [Generating a deck](powerpoint.md#generating-a-deck) |
| Inspect or update an Excel workbook | [Excel support](excel.md) | [Document plans](document-plans.md) |
| Generate many documents from a template | [Template population](document-workflows.md#populate-a-template-batch) | [TemplateBatch sample](../samples/TemplateBatch/) |
| Check a template batch before it writes | [Template preflight](document-workflows.md#preflight-a-template-batch) | [Capability discovery](capability-discovery.md) |
| Compare two Word documents and create a redline | [Document comparison](document-workflows.md#compare-two-word-documents) | [DocumentComparison sample](../samples/DocumentComparison/) |
| Assemble a proposal and appendices | [Word document assembly](document-assembly.md) | [DocumentAssembly sample](../samples/DocumentAssembly/) |
| Add document tools to an agent | [Agent integration](agent-integration.md) | [Document providers](document-providers.md) |
| Connect an MCP client locally | [MCP server](mcp-server.md#local-hosting-stdio) | [Deployment](deployment.md#option-a---local-stdio-for-claude-code-and-codex) |
| Host OfficeAgent for a team | [Deployment](deployment.md#option-b---hosted-http-for-all-four-clients) | [Operations](operations.md), [troubleshooting](troubleshooting.md) |
| Enforce per-user connection access | [Hosted gateway reference](../samples/HostedGateway/) | [Deployment](deployment.md#authentication--identity) |
| Render pages for visual verification | [Visual rendering](rendering.md) | [Operations](operations.md) |
| Read or write SharePoint files | [SharePoint provider](document-providers.md#the-sharepoint-provider) | [Authentication and identity](deployment.md#authentication--identity) |
| Look up an operation's JSON | [Document plans](document-plans.md) | [Validation errors](document-plans.md#validation-errors) |
| Run a tracked Word review | [Sample review workflows](../samples/documents/README.md#reproducible-review-workflows) | [Optional review skill](skill-installation.md) |
| Edit a document that already has tracked changes | [Pending revisions](pending-revisions.md) | [Document plans](document-plans.md#validation-errors) |
| Bound what an untrusted document can spend | [Ingestion limits](ingestion-limits.md) | [Operations](operations.md) |
| Recover from a failed or uncertain save | [Storage outcomes and recovery](recovery.md) | [Operations](operations.md#what-a-failed-commit-tells-you-about-storage) |
| Find out what the server supports before planning | [Capability discovery](capability-discovery.md) | [Document plans](document-plans.md) |
| Measure a candidate, or check for a regression | [Benchmarks](benchmarks.md) | [Word preservation evidence](word-preservation-evidence.md) |
| Move an existing integration from 0.9 to 1.0 | [Upgrading to 1.0](upgrading-to-1.0.md) | [Wire contract](wire-contract.md) |

## Packages

| Package | Use it for | Target frameworks |
| --- | --- | --- |
| `OfficeAgent.Core` | Engine, direct .NET API, filesystem provider | `netstandard2.0`, `net8.0` |
| `OfficeAgent.Word` | Word `.docx` inspection, creation, and editing | `netstandard2.0`, `net8.0` |
| `OfficeAgent.PowerPoint` | PowerPoint `.pptx` inspection, creation, and editing | `netstandard2.0`, `net8.0` |
| `OfficeAgent.Excel` | Excel `.xlsx` inspection, creation, and editing | `netstandard2.0`, `net8.0` |
| `OfficeAgent.AgentFramework` | `Microsoft.Extensions.AI` / Microsoft Agent Framework tools | `netstandard2.0`, `net8.0` |
| `OfficeAgent.SharePoint` | Microsoft Graph document provider | `netstandard2.0`, `net8.0` |
| `OfficeAgent.Mcp` | Standalone MCP server and .NET global tool | `net8.0` |
| `OfficeAgent.Rendering` | Optional bounded LibreOffice and Poppler page-image rendering | `net8.0` |

Applications add `OfficeAgent.Core` plus at least one format module. Add a
provider or agent adapter only when that hosting model needs it. The standalone
MCP tool already includes all three format modules and both built-in providers.

## Core workflow

Every integration uses the same safety loop:

1. **Register or create** a document and keep its opaque `(connectionId, documentId)`.
2. **Inspect** structure and capture a snapshot.
3. **Find** text to obtain content-verified anchors.
4. **Preview** the complete plan without writing.
5. **Commit** the complete in-memory transaction through the provider and inspect `writeOutcome`.
6. **Open the result** only after `committed`; reconcile `possibleOutput` before retrying an uncertain write.
7. **Verify in native Office** when layout, recalculation, or consequential review matters.

The engine refuses stale snapshots, mismatched anchors, unsupported operations,
and optimistic-concurrency conflicts instead of guessing.

## Documentation map

### End-user tutorials and evidence

- [Library selection guide](choose-officeagent.md) - supported jobs, explicit non-goals, package choices, alternatives, and a verified direct .NET recipe.
- [Word preservation evidence](word-preservation-evidence.md) - operation-specific changed parts, protected bytes, semantic checks, refusals, and native Office limits.
- [Native Office compatibility](native-compatibility.md) - every advertised operation opened in Word, PowerPoint and Excel, with the harness, its controls, and known unsupported cases.
- [Getting started](getting-started.md) — first successful Word edit from C#.
- [Deployment and client setup](deployment.md) — local and hosted MCP recipes.
- [Skill installation](skill-installation.md) — optionally teach Claude Code or Codex a repeatable Word review workflow.

### Concepts

- [Concepts](concepts.md) — providers, anchors, snapshots, plans, and transactions.
- [Pending revisions](pending-revisions.md) - the text view of a redlined document, which successive edits are allowed, and what accept and reject guarantee.
- [PowerPoint support](powerpoint.md) — slide-specific addressing and behavior.
- [Excel support](excel.md) — worksheet, range, table, formula, and cell-note behavior.

### How-to guides

- [Agent integration](agent-integration.md) — expose bounded tools and deliver output.
- [Document providers](document-providers.md) — filesystem and SharePoint storage.
- [Open XML SDK interoperability](sdk-interoperability.md) - the tested copy, edit, reinspect workflow, per-stage guarantees, and the two distinct staleness guards.
- [Operations](operations.md) — concurrency, memory, telemetry, and production operation.
- [Ingestion limits](ingestion-limits.md) - host ceilings for package size, expansion, parts, and XML, plus what is refused and how.
- [Capability discovery](capability-discovery.md) - the contract version, per-format verbs and change modes, effective ceilings, and authorized connections.
- [Visual rendering](rendering.md) — optional out-of-process PDF and page-image conversion.
- [Template population and comparison](document-workflows.md) — batch binding, repeating Word rows, comparison coverage, and redline generation.
- [Troubleshooting](troubleshooting.md) — diagnose setup, identity, and edit failures.
- [Support and compatibility](../SUPPORT.md) — supported runtimes, version policy, release assurances, and help routes.
- [Compatibility policy](compatibility.md) — what 1.x promises for the API, JSON, configuration and defaults, and what it excludes.
- [Security policy](../SECURITY.md) — supported versions, private reporting, and the hosted security boundary.

### Reference

- [Generated C# API reference](csharp-api.md) - public types and members from every library package.
- [Document plans](document-plans.md) — operation JSON and validation codes.
- [MCP server](mcp-server.md) — transports, settings, tools, and response contracts.

### Maintainer and historical evidence

- [Maintenance and continuity](maintenance.md) - ownership, release accounts, runtime lifecycle, and repository continuity.
- [Releasing](releasing.md) — the maintainer steps for NuGet, the MCP Registry, GitHub releases, and repository metadata.
- [Adoption validation](adoption-validation.md) — fresh-user trials required before broad promotion; blank rows are not evidence.
- [v0.9.0 agent selection evaluation](../evaluations/agent-selection/v0.9.0/README.md) - historical controlled protocol with explicit live-evidence limits.
- [v0.9.0 benchmarks](../evaluations/benchmarks/v0.9.0/README.md) - historical measurements, not 1.0 adoption proof.

## Scope

OfficeAgent edits OOXML packages and does not automate desktop Office. Core structural preview
does not render pages or slides. The optional rendering package can produce page images through
external LibreOffice and Poppler processes, but it does not evaluate Excel formulas or make
Office's own layout engine available. Automated page-fit and overflow analysis remain outside
the current scope. See the
[project limitations](../README.md#scope-and-limitations) before production use.

## Get help

Start with [Troubleshooting](troubleshooting.md). If the problem remains, open a
[GitHub issue](https://github.com/ilia-sokolov/OfficeAgent.NET/issues) with the
package version, provider, document format, operation, stable error code, and a
small sanitized reproduction. Do not attach confidential documents, access
tokens, registration indexes, or tenant identifiers.
