# Ingestion and resource limits

Every OOXML package OfficeAgent opens passes through one code path, and the host
ceilings below are applied there. A malformed or hostile document fails predictably
instead of spending unbounded memory, CPU, or disk.

> [!IMPORTANT]
> These are resource ceilings inside your process, not an operating-system sandbox.
> They bound what one package can spend; they do not isolate the process, cap total
> memory across concurrent calls, or replace OS limits. Renderer isolation is a
> separate concern documented in [visual rendering](rendering.md).

## The ceilings

Configured with `OpenXmlIngestionLimits`.

| Limit | Default | What it bounds |
| --- | --- | --- |
| `MaximumCompressedBytes` | 128 MiB | The package as read. A stream is refused at this point rather than drained first |
| `MaximumExpandedBytes` | 512 MiB | Total expanded size of all parts |
| `MaximumPartBytes` | 128 MiB | Expanded size of any single part |
| `MaximumParts` | 4000 | Entries in the package |
| `MaximumXmlDepth` | 256 | Element nesting in package XML |
| `MaximumXmlCharacters` | 64 MiB | Characters in a single XML part |
| `MaximumExpansionRatio` | 200 | Expanded-to-compressed ratio before a package is treated as a decompression bomb |

A ceiling of zero or less is a configuration error and throws at construction. There is
no value that disables a limit.

## Setting them

Register the limits before `AddOfficeAgent()`:

```csharp
services.AddSingleton(new OpenXmlIngestionLimits
{
    MaximumCompressedBytes = 32L * 1024 * 1024,
    MaximumParts = 1000
});
services.AddWordFormat();
services.AddOfficeAgent();
```

The engine reads the registered `OpenXmlIngestionLimits`; a client built without the container,
such as `new OfficeAgentClient(new WordModule())`, uses the defaults.

A caller may ask for something stricter with `limits.Restrict(requested)`, which takes
the smaller of each pair. **A request can lower an effective limit; it can never raise a
host ceiling.** Host policy is the maximum, always.

## Which entry points are covered

| Entry point | Covered by |
| --- | --- |
| `Inspect`/`Find`/`Preview`/`Commit` on a byte array | Compressed ceiling, then every package ceiling |
| The same on a `StreamHandle` | Bounded read at the handle, before the engine copies anything |
| The same on a `FileHandle` | File length checked before the file is read at all |
| Provider reads (filesystem, SharePoint, iManage, Google Drive, Box, Azure Blob, Amazon S3) | The provider's own `MaximumBytes`, then every package ceiling |
| `create_document`, `edit_document`, `edit_documents`, `refresh_document`, `open_document` | Every package ceiling, for each document, and for the template `create_document` starts from |
| Inline base64 tools (`inspect_document_content` and friends) | Refused from the encoded length before decoding, then every package ceiling |
| Session import and export | Every package ceiling |
| Template population, comparison | Every package ceiling, per document read |
| Word assembly | `DocumentMergeLimits` in addition, which bounds source count and combined size |

The filesystem provider keeps its own `MaximumBytes` (default 100 MiB), applied on
register, open, and save. It is a storage policy and is independent of these ceilings;
a document must satisfy both.

## What is refused, and how

| Input | Result |
| --- | --- |
| Package larger than the compressed ceiling | `input-too-large`, naming the limit and the observed size |
| Endless or non-seekable stream | `input-too-large` at the ceiling; reading stops there |
| Decompression bomb | `input-too-large`, refused from the declared ratio before any part is expanded |
| File protected by a sensitivity label or a password | `encrypted-document`; OfficeAgent does not decrypt |
| Truncated or corrupt archive | `malformed-package` |
| Bytes that are not an archive | `malformed-package` |
| Duplicate part name | `malformed-package`; a duplicated name makes the package ambiguous |
| Traversing (`../`) or absolute entry name | `malformed-package` |
| Document type definition or external entity in the manifest | `malformed-package`; DTDs are prohibited and no entity is resolved |
| Missing `[Content_Types].xml` | `malformed-package` |
| Relationship with no target or an invalid target, part reached under two spellings, unreadable core properties part, empty main or slide part, part with no root element, or attribute that does not parse as its type | `malformed-package`, whether found at open or when the part is first read |
| XML deeper than the depth ceiling | `input-too-large` |
| Macro-enabled template a new document would start from (`.dotm`, `.potm`, `.xltm`, `.docm`, `.pptm`, `.xlsm`, or any package carrying a VBA project) | `macro-enabled-document`; OfficeAgent does not copy macros into the documents it creates |

Directly, these surface as `OpenXmlIngestionLimitException`,
`OpenXmlPackageRejectedException`, `EncryptedDocumentException` and
`MacroEnabledDocumentException`. Through agent and MCP tools they surface as the stable wire
codes `input-too-large`, `malformed-package`, `encrypted-document` and
`macro-enabled-document`. The last applies only where a document is created from a template
(`create_document` with `templateSource`, `CreateFromTemplateAsync`, and template discovery,
preview and population).

The inline base64 tools are the one deliberate exception. Content that decodes but is
not a readable package is reported as `invalid-argument` there, because on that path the
useful advice concerns the copy rather than the package: re-sending the same string
fails identically. A genuine ceiling breach still reports `input-too-large`.

## Guarantees on refusal

- Nothing is read past the limit.
- No output document is produced.
- Stored bytes are unchanged; a refused provider read leaves the file byte-identical.
- No network access is made on account of package content. Legitimate external
  hyperlinks are preserved in the document and never fetched.
- Diagnostics name the limit, the ceiling, and the observed value, and stay short. A
  refusal never echoes document content back to the caller.

## How the limits are tested

`BoundedIngestionTests` checks each ceiling at, below and above its boundary with hand-built
packages, and runs a fixed set of malformed seeds through every entry point.

Beyond the unit tests, a fuzz target feeds the ingestion path inputs nobody wrote by hand:
[`tests/OfficeAgent.Fuzz`](../tests/OfficeAgent.Fuzz/README.md). It wraps arbitrary bytes in a
`StreamHandle` and inspects them through a client with the Word, PowerPoint and Excel modules,
under the default ceilings, using SharpFuzz and libFuzzer, seeded from `samples/documents` and the
test corpus. An input must produce a result or one of the refusals `input-too-large`,
`malformed-package` or `encrypted-document`. Any other
exception, a hang past 10 seconds, more than 2 GiB of live managed memory, more than 1 GiB
allocated by one inspection, or a refusal message of 600 characters or more is a finding. It runs
locally from its README; it is not part of CI.

## Tuning

Raise `MaximumCompressedBytes` and `MaximumPartBytes` together when your documents carry
large embedded media. Raise `MaximumParts` for documents with very many slides, images,
or embedded objects. `MaximumExpansionRatio` is the one to leave alone unless a
legitimate document is being refused: ordinary Office files sit far below 200, and
lowering it is a cheap way to reject bombs earlier.

Set ceilings from what your own corpus actually needs, not from what a single large
document happens to be. A limit that admits every document you have ever seen bounds
nothing.
