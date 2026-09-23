# Word preservation evidence

This matrix preserves the historical `v0.9.0` corpus evidence for a fixed fictional
Word package across two supported operations and two unsupported operations. It is
operation-specific evidence, not a claim that every Word feature or document is
preserved. For the broader 1.0 corpus and desktop-application results, see
[native Office compatibility](native-compatibility.md).

The recorded run is based on commit
`a687bc15a1d7af87790333984be4721055bdb90e`. The input fixture SHA-256 is
`68b8217443aaf8bd381ec8164e2d0f2237a6db5aafdc6deb8f9568d9b4626c1c`.
See the tracked [manifest](../tests/OfficeAgent.Tests/Corpus/v0.9.0/word-preservation/manifest.json),
[generated report](../tests/OfficeAgent.Tests/Corpus/v0.9.0/word-preservation/report.json),
and [inspectable result packages](../tests/OfficeAgent.Tests/Corpus/v0.9.0/word-preservation/README.md).

## Evidence levels

| Level | Result | What it establishes |
| --- | --- | --- |
| Generated fixture identity | Passed | The fictional package, operation outputs, recipes, and hashes are fixed and rerunnable. |
| Office 2019 Open XML schema | Passed | The input and two supported-operation outputs have no validator errors. |
| Protected-part bytes | Passed | Every declared protected ZIP part has the same SHA-256 before and after its operation. |
| Changed-part semantics | Passed | Deliberate assertions validate revisions, comments, identifiers, controls, numbering, fields, and sections inside changed XML. |
| Refusal state | Passed | Unsupported Word mutations create no direct output and do not change provider bytes, version, or count. |
| Native Microsoft Word | Partial | Word 16.0.20326.20144 opened the tracked result. Accept and reject are `NOT_RUN`. |

## Operation matrix

| Category | Operation | Current boundary | Evidence |
| --- | --- | --- | --- |
| Mixed runs | Tracked `changeText`, `target` to `objective` across bold and italic runs | Supported edit. Only `word/document.xml` changes. | Old text is a deletion, new text is an insertion, with fixed author and time. |
| Existing insertions and deletions | Same tracked replacement in another paragraph | Preservation only. | Existing revision IDs 100 and 101 retain author, time, and text. |
| Classic comments | Same tracked replacement | Preservation only. `word/comments.xml` is byte-protected. | Comment ID, author, body, paragraph identity, and range markers remain. |
| `commentsExtended` | Tracked replacement, then separate resolve case | Preserved during text edit. Changed only for resolve. | Resolve changes only the `done` flag for the matching paragraph identity. |
| `commentsIds` | Text edit and comment resolve | Preservation only. `word/commentsIds.xml` is byte-protected. | Paragraph identity `20000001` remains mapped to durable identity `D0000001`. |
| `people` | Text edit and comment resolve | Preservation only. `word/people.xml` is byte-protected. | `Reviewer One` remains associated with the comment author. |
| Paragraph identifiers | Tracked replacement | Preservation only outside the targeted content. | All seven body paragraph IDs remain unchanged. |
| Nested controls | Tracked replacement outside the controls | Preservation only. | The `OuterClause` block and `InnerParty` run control XML subtrees remain equal. |
| Numbering | Tracked replacement outside the numbered paragraph | Preservation only. `word/numbering.xml` is byte-protected. | Numbering instance 42 and its paragraph reference remain. |
| Field | Tracked replacement outside the field | Preservation only. | The `PAGE` instruction and cached result `7` remain. |
| Sections | Tracked replacement outside the section boundaries | Preservation only. | The landscape next-page section and final portrait section remain. |
| Excel-only mutation | `setCell` against the Word fixture | Refused as `unsupported-operation`. | No output or provider mutation. |
| PowerPoint-only mutation | `insertSlide` against the Word fixture | Refused as `unsupported-operation`. | No output or provider mutation. |

Relationship parts `_rels/.rels` and `word/_rels/document.xml.rels` are
byte-protected in both supported operations. Corruption controls deliberately
damage a classic comment part, a relationship, an existing revision author, and
a durable comment identity. Each must make the verifier fail.

## Reproduce the report

Run the focused suite:

```powershell
dotnet test tests/OfficeAgent.Tests/OfficeAgent.Tests.csproj `
  --configuration Release `
  --filter "FullyQualifiedName~WordPreservationEvidenceTests"
```

To regenerate the committed fixture, operation outputs, and JSON report from
fixed inputs, follow the explicit-output recipe in the
[corpus README](../tests/OfficeAgent.Tests/Corpus/v0.9.0/word-preservation/README.md).
The ordinary test compares fresh in-memory evidence with the committed report
and fails on drift.

## Limits

The fixtures are generated, fictional, and repository-owned. Schema validity
does not establish layout, pagination, visual parity, or Word behavior. Exact
part preservation applies only to the declared operations and protected parts.
Microsoft Word 16.0.20326.20144 opened the tracked result in Compatibility Mode.
The tracked replacement, prior revisions, classic comment, nested controls,
numbering, field, and two-page section layout were visible. This is one native
opening observation, not a visual-parity guarantee. Native accept and reject
were not performed for this historical corpus. The separate 1.0 compatibility
corpus includes those checks.
