# Compatibility

What OfficeAgent.NET 1.x promises, what it does not, and how each promise is enforced. For the
changes between 0.9 and 1.0, see [upgrading to 1.0](upgrading-to-1.0.md).

## The promise

Within 1.x, an application built against 1.0 keeps compiling, keeps running, and keeps
receiving the JSON it was written for. Breaking changes wait for 2.0.

"Breaking" is measured against two generated baselines, not against intent:

| Baseline | Records | Covers |
| --- | --- | --- |
| [C# API reference](csharp-api.md) | Every public type and member, with constant values, enum values, default arguments, `init` versus `set`, nullability, base types, protected members and stability class | Compiled and recompiled .NET callers |
| [Wire contract](wire-contract.md) | Plan operations and their members, anchors, every tool's input schema, the response shape of all 21 tools and the error envelope, configuration keys, and the default value of every member a caller constructs | Agents, MCP clients, stored plans, configuration files |

CI regenerates both and fails when either differs from the committed file. A change that shows
up in either diff is a contract change and needs a recorded decision. Each gate has been checked
by deliberately introducing a breaking change against a passing control, and it caught every
one.

## Four version numbers

The package version is not the version of any wire format. Each format has its own version, and
a new package major version does not change them.

| Identity | 1.0 value | Where it appears |
| --- | --- | --- |
| Package | `1.0.0` | NuGet, the MCP server's `serverInfo` |
| Edit plan | `0.2` | `DocumentPlan.ContractVersion`, `contractVersion` in a plan |
| Merge plan | `1` | `DocumentMergePlan.Version` |
| Apply receipt, merge receipt | `1`, `1` | `receiptVersion` on each receipt |

`describe_capabilities` reports the edit-plan and receipt versions under `contracts`. A
receipt that omits `receiptVersion` was written by a build that emitted schema `1`, which every
release since 0.8 has done.

## Requests: strict

What a caller sends is checked, and anything the contract does not define is refused before
any document is read or written.

- **Unknown properties, operations and enum names** are refused with `invalid-json` on the agent
  and MCP surfaces. An integer where an enum name is expected is refused the same way.
- **An edit plan's `contractVersion`** must be omitted or exactly `"0.2"`. Anything else,
  including `null`, an empty string or a future version, is refused with `contract-mismatch`.
  A version is never silently upgraded or reinterpreted.
- **A merge plan's `version`** must be `"1"`; anything else makes the plan invalid before a
  source is read.
- **Names are read case-insensitively.** `paraId`, `ParaId` and `PARAID` are the same member.
  Values are not: enum names and verbs follow the documented spelling.

Full table: [document plans, contract compatibility](document-plans.md#contract-compatibility).

## Responses: tolerant readers

What OfficeAgent returns may grow within 1.x. A reader written for 1.0 must:

- **ignore properties it does not recognise.** New fields may be added to any response.
- **treat an unrecognised enum value or code as a generic case.** New change kinds, formats,
  diagnostic codes and error codes may appear. Code *values* never change; a code in one of the
  public catalogues keeps its string for all of 1.x.
- **not depend on property order, whitespace or array order** unless a document says the order
  is meaningful, such as operations in a plan or sources in a merge.

In exchange, within 1.x no response property is removed, renamed or given a different JSON
kind, and every property name stays camelCase. Enum values are written as their names.

**Nulls.** A property that can be absent is written as `null` rather than omitted. The
[wire contract](wire-contract.md#tool-responses) records which properties can be `null`. A
property that is never `null` in 1.0 does not become nullable within 1.x.

**Deserialising into the library's own types.** Use `JsonSerializerDefaults.Web` (or
`PropertyNameCaseInsensitive = true`) with a `JsonStringEnumConverter`. System.Text.Json's
plain defaults match names case-sensitively in PascalCase, and every member of these types has a
default, so a mismatch produces plausible defaults rather than an error.

## Defaults and security-relevant behaviour

These are part of the contract and are recorded in the [wire contract's defaults](wire-contract.md#defaults):

- A Word edit that names no `mode` is tracked. Connections default to `Tracked`, and so does
  `ChangeTextOp.Mode`. PowerPoint refuses an explicit `Tracked` rather than writing an untracked
  edit.
- Creation, registration and inline content are opt-in on the in-process tools. On the MCP server,
  registration defaults to on and creation to off.
- Every ingestion ceiling, and every merge, comparison and template-batch limit, keeps its default.
  A request can lower a host ceiling but can never raise it.
- Connection access is checked before any read or write, and a denied call fails with
  `connection-forbidden` without touching the document. Discovery (`describe_capabilities`,
  `list_connections`) omits connections the caller cannot use rather than listing them as denied.
- Content that no longer matches is refused, never edited in its place. A missing anchor, a
  changed `expect`, a changed snapshot, a changed template batch or a changed merge source fails
  with `anchor-not-found`, `expect-mismatch`, `stale-snapshot`, `stale-batch-preview` or
  `stale-merge-source`.

A 1.x release may **tighten** one of these to fix a vulnerability, for example by refusing an
input that was accepted before. It will never loosen one. The changelog calls out every such
tightening, with the stable code the refusal carries.

## Not part of the contract

- **Tool descriptions and prompt text.** They are guidance for a model and are tuned between
  releases. Tool names and input schemas are the contract; the wire contract strips
  descriptions for this reason.
- **Human-readable messages.** Match on `code`, never on `message`.
- **Output bytes.** The same plan produces the same document semantics across 1.x, but not
  necessarily the same bytes: part ordering, relationship ids and XML serialisation may change.
  Receipts hash the exact bytes a particular run produced.
- **Rendering.** `OfficeAgent.Rendering` drives an external LibreOffice install, and its output
  depends on that install.
- **Experimental types.** See below.

## Engine extensibility

A format module and its operation handlers can live in their own assembly, so the types they are
built from are public. They are also the engine's internal architecture, and on netstandard2.0
an interface cannot gain a member without breaking every implementer. Freezing them would freeze
the engine, so they are excluded from the 1.x promise and marked
`[Experimental("OFFICEAGENT001")]`.

| Area | Types |
| --- | --- |
| Engine and pipeline | `IDocumentService`, `ICapabilityReportingService`, `IHandleResolver`, `IOpenXmlPackage`, `ApplyContext`, `OperationPreview`, `CapabilityDiscovery` |
| Format modules | `IFormatModule`, `IOperationHandler`, `IApplyTimeProvider`, `IBlankDocumentFactory`, `ICapabilityDeclaringModule`, `IPlanValidatingModule`, `IDocumentAssembler`, `DocumentAssemblyCandidate` |
| Text engine | `ITextDialect`, `WordmlDialect`, `PresentationmlDialect` |
| Word and PowerPoint internals | `IWordNodeProvider`, `WordObjectMap`, `OfficeAgent.Word.ResolvedNode`, `IPowerPointNodeProvider`, `IPowerPointOperationHandler`, `PowerPointObjectMap`, `OfficeAgent.PowerPoint.ResolvedNode` |

Naming one of these types in your code produces error `OFFICEAGENT001`. To build against them
anyway, accept that a 1.x minor release may change them and suppress the diagnostic:

```xml
<PropertyGroup>
  <NoWarn>$(NoWarn);OFFICEAGENT001</NoWarn>
</PropertyGroup>
```

Hosting the engine never needs them. `new OfficeAgentClient(new WordModule())`,
`AddWordFormat()` and `AddOfficeAgent()` are stable, and passing a module to a constructor is
not a use of `IFormatModule`, because the compiler reports only types the caller names. The
diagnostic works for netstandard2.0 consumers too. Samples are built without the suppression, so
a sample that reached for a seam would fail the build.

These are the host extension points, and they are stable. Hosts implement them, and 1.x will not
add members to them:

| Interface | Implemented to |
| --- | --- |
| `IDocumentProvider`, `IDocumentCreatingProvider`, `IConnectionEditingDefaults` | Add a storage backend |
| `IConnectionAccessPolicy`, `ITrustedPrincipalAccessor` | Authorise tool calls |
| `IAuditActorProvider` | Stamp receipts with the acting user |
| `IAccessTokenProvider`, `ISharePointRegistrationStore` | Customise SharePoint authentication and registration |
| `IDocumentRenderer` | Plug in a renderer |

`ITextFormat` and `ITrackedOperation` are stable to read but are implemented only by the
library's own plan records, so 1.x may add members to them. A test fails whenever a new public
interface appears without being placed in one of these groups.

## Deprecation

A member or wire feature to be removed in 2.0 is first marked `[Obsolete]` with the replacement
named in the message, and noted in the changelog, at least one minor release before 2.0. It keeps
working until then. Nothing is removed within 1.x.

## Platforms

The libraries target `netstandard2.0` and `net8.0`; `OfficeAgent.Rendering` and the MCP server
target `net8.0`. Both library targets expose the same public surface, and the gate checks this on
every build. It loads the netstandard2.0 binaries, renders them with the same renderer as the
baseline, and fails on any type or member that exists in only one target.
[SUPPORT.md](../SUPPORT.md#runtime-and-platform-compatibility) states which targets CI executes as
well as compiles.
