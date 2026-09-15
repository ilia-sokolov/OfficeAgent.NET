# Capability discovery

Before planning an edit, ask the server what it supports. `describe_capabilities` (and
`OfficeAgentClient.DescribeCapabilities` in .NET) reports the accepted contract version,
every registered format with the verbs and change modes it takes, the host's ingestion
ceilings, and the connections you may use.

Everything in the result is derived from what is actually registered. A module that is
not registered contributes nothing; a verb no handler accepts is not listed. **If a verb
is absent for a format, sending it there is refused.** Advertising and refusing are the
same decision, and a test reconciles them for every verb in the vocabulary.

## What comes back

```json
{
  "Contracts": { "EditPlan": "0.2", "ApplyReceipt": "1", "MergeReceipt": "1" },
  "Formats": [
    {
      "Format": "Word",
      "Operations": ["changeText", "comment", "format", "..."],
      "ChangeModes": ["Direct", "Tracked"],
      "NodeKinds": ["comment", "docProperty", "image", "note", "revision", "table"]
    }
  ],
  "Limits": { "MaximumCompressedBytes": 134217728, "MaximumParts": 4000, "...": 0 },
  "RenderingAvailable": false,
  "Connections": [
    { "ConnectionId": "workspace", "Provider": "filesystem", "Allowed": ["Read", "Edit"] }
  ],
  "RequiresInspection": ["Anchor identifiers. …"]
}
```

| Field | Meaning |
| --- | --- |
| `Contracts` | The only edit-plan version accepted, and the receipt schema versions emitted. A plan with any other `contractVersion` fails with `contract-mismatch` |
| `Formats[].Operations` | The plan verbs at least one registered handler for that format accepts |
| `Formats[].ChangeModes` | What the format can honour. PowerPoint and Excel are `Direct` only; a tracked request there is refused, not silently downgraded |
| `Formats[].NodeKinds` | The node kinds that format's inspection surfaces, which are the `kind` values a `NodeAnchor` may use |
| `Limits` | The host [ingestion ceilings](ingestion-limits.md) in force |
| `RenderingAvailable` | Whether an optional page renderer is wired up. Rendering is not part of the core engine |
| `Connections` | Only the connections the caller may use, each with only the capabilities they hold |
| `RequiresInspection` | What discovery deliberately cannot answer |

## Access is not bypassed

Through the agent and MCP surface, the connection list is filtered by the host's
`IConnectionAccessPolicy` for the calling principal. A connection the caller cannot touch
at all is **absent** rather than listed as denied, so the result does not disclose that it
exists. A caller with partial access sees the connection with exactly the capabilities
they hold.

The direct .NET client has no policy of its own. It reports every registered connection
and leaves `Allowed` empty, because it does not know who is asking. Filtering belongs to
the layer that holds the policy.

## What discovery cannot tell you

`RequiresInspection` names these explicitly so silence is not read as support:

- **Anchor identifiers.** Paragraph, table, slide and shape ids come from
  `inspect_document` for the document in hand. They are not predictable and are never
  fabricated here.
- **Content-dependent refusals.** Whether a particular edit is legal can depend on the
  document, for example an edit that
  [spans a pending tracked revision](pending-revisions.md).
- **Template bindings.** Which tags and repeating rows a template exposes is a property of
  that template.

## Extending it

Adding a verb to a module's handlers makes it appear automatically. Adding a format
module makes that format appear. Nothing needs to be listed twice.

Where a rule cannot be discovered by asking a handler - change modes live in plan-wide
validation, which needs an open package - the module declares it through
`ICapabilityDeclaringModule`, and a reconciliation test proves the declaration matches
what validation actually enforces. A declaration that drifts from behavior fails that
test rather than misleading an agent.
