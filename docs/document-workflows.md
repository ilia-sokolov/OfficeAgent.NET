# Template population and document comparison

To combine whole documents in a specified order, see [Word document assembly](document-assembly.md).

OfficeAgent 0.8 adds two higher-level workflows built on the same inspect, plan, preview,
commit, provider, and receipt contracts as ordinary edits. Template population creates
independent documents from tagged values and repeating Word rows. Comparison reads two Word
documents and proposes a native tracked-change plan when it can account for every change in
its supported scope and verify that unsupported package content is unchanged.

## Populate a template batch

Tag scalar Word content controls with stable names such as `CustomerName`. In a row that
should repeat, write placeholders such as `{{Description}}` and `{{Amount}}`, then inspect the
template and copy its table path. Tags must be unique and the row is selected explicitly by
its zero-based index.

```csharp
var request = new TemplateBatchRequest
{
    Items = new[]
    {
        new TemplateBatchItem
        {
            OutputName = "quote-acme.docx",
            Binding = new TemplateBinding
            {
                Values = new Dictionary<string, string?> { ["CustomerName"] = "Acme BV" },
                RepeatingTables = new[]
                {
                    new RepeatingTableBinding
                    {
                        TablePath = "table#0",
                        TemplateRowIndex = 1,
                        Records = new IReadOnlyDictionary<string, string?>[]
                        {
                            new Dictionary<string, string?>
                            {
                                ["Description"] = "Consulting", ["Amount"] = "1200.00"
                            }
                        }
                    }
                }
            }
        }
    }
};
var result = await client.PopulateTemplateBatchAsync(template, request);
```

The engine inspects and binds the source again for every item, creates a snapshot-bound plan,
and saves with `NewDocument`. It never changes the template. Each attempted item returns its
own provider reference, report, receipt, and stable diagnostics. `ContinueOnError=false`
stops after the first failed item; `MaximumDocuments` defaults to 100.

The batch is a sequence of independent provider creates, not one storage transaction. Outputs
saved before a later failure remain saved. With `ContinueOnError=false`, results cover only the
items attempted through the first failure. A provider may also accept a create before its
response is lost; use the provider's recovery process before retrying the same output name.

`MissingValueBehavior` is `Fail`, `Ignore`, or `Empty`; it applies to scalar slots and row
placeholders. `RejectUnknownValues` defaults to true. Population defaults to `Direct`, which
suits generated documents; select `Tracked` when a reviewer needs to accept or reject the
population in Word. Repeating rows are Word-only. Scalar binding also supports unique
PowerPoint shape names.

The equivalent MCP or Agent Framework call is `populate_template_batch(connectionId,
documentId, requestJson)`. The host must opt in with `AllowCreation`; its connection access
policy must also grant read and create. For example:

```json
{
  "items": [{
    "outputName": "quote-acme.docx",
    "binding": {
      "values": { "CustomerName": "Acme BV" },
      "repeatingTables": [{
        "tablePath": "table#0",
        "templateRowIndex": 1,
        "records": [{ "Description": "Consulting", "Amount": "1200.00" }]
      }],
      "missingValueBehavior": "Fail",
      "rejectUnknownValues": true,
      "mode": "Direct"
    }
  }],
  "continueOnError": true,
  "maximumDocuments": 100
}
```

Run the complete offline [quote-generation sample](../samples/TemplateBatch/).

## Preflight a template batch

Three steps, in order: find out what the template offers, check the whole batch without
writing anything, then commit the exact thing you checked.

### 1. Discover what the template offers

```csharp
var discovered = await client.DiscoverTemplateAsync(template);

foreach (var slot in discovered.Slots.Where(slot => slot.Bindable))
    Console.WriteLine($"{slot.Name} ({slot.Kind})");
```

`Slots` are confirmed: a content control or a named shape really is in the file. A slot
whose tag occurs more than once has `Bindable` false and a matching
`ambiguous-template-slot` diagnostic, because there is no way to tell which one a value
was meant for.

`RepeatingRows` are **candidates**, with `Confirmed` false. Nothing in WordprocessingML
declares a row repeatable; these are inferred from the `{{Field}}` placeholder
convention. Treat them as a starting point for a binding you then preview, not as a
promise.

### 2. Preview the whole batch

```csharp
var preview = await client.PreviewTemplateBatchAsync(template, request);

if (!preview.IsValid)
    foreach (var item in preview.Items.Where(item => !item.IsValid))
        Report(item.OutputName, item.Diagnostics);
```

Preview writes nothing. It validates **every** item, not just up to the first failure, so
one call surfaces everything you have to fix. Per-item diagnostics are capped by the
effective budget, and `DiagnosticsTruncated` says when a list is partial rather than
complete.

Batch-level problems, such as a duplicate output name or a batch larger than the host
allows, arrive in `preview.Diagnostics`.

### 3. Commit the intent you reviewed

```csharp
var result = await client.PopulateTemplateBatchAsync(template, request, preview.Token);
```

The token binds the preview to the exact template bytes and the exact batch. Pass it and
the commit is refused with `stale-batch-preview` unless both still match. Anything a
reviewer would care about changes it:

| Change | Effect |
| --- | --- |
| A bound value | `BatchSha256` changes |
| An output name | `BatchSha256` changes |
| Item order | `BatchSha256` changes; a preview of one order does not authorize another |
| Adding or removing an item | `BatchSha256` changes |
| Editing the template | `TemplateSha256` changes |

Omitting the token keeps the older one-call behavior. The batch is still preflighted
either way: a batch that cannot validate is refused before any output exists, rather than
leaving a prefix of it in storage.

### Host budgets

`TemplateBatchLimits` is owned by the host:

| Budget | Default |
| --- | --- |
| `MaximumDocuments` | 100 |
| `MaximumRowsPerDocument` | 2,000 |
| `MaximumFieldsPerDocument` | 500 |
| `MaximumTotalRows` | 20,000 |
| `MaximumValueLength` | 100,000 |
| `MaximumDiagnosticsPerItem` | 50 |

```csharp
client.WithTemplateLimits(new TemplateBatchLimits { MaximumDocuments = 25 });
```

A request may ask to be held to something stricter through `TemplateBatchRequest.Limits`.
The effective budget is the **stricter** of the two, so a request cannot raise a host
ceiling. `TemplateBatchRequest.MaximumDocuments` is treated the same way: a request, not
a ceiling. Call `client.EffectiveLimits(request)` to see what a given request will
actually run under.

### What each item outcome means

| Outcome | What happened | What to do |
| --- | --- | --- |
| `Committed` | The output exists and is registered | Use `Document` |
| `Failed` | Refused before any write. This output does not exist | Fix the diagnostics and retry |
| `Skipped` | Never attempted, because an earlier item failed and the batch stops on error | Retry once the earlier failure is fixed |
| `Uncertain` | The provider was given the bytes and then failed to confirm | Do not retry the name blindly. Inspect the destination and reconcile |
| `Previewed` | Validation only | Nothing was written |

`Committed` on the whole result still means every item committed.

**There is no multi-output atomicity.** Each output is assembled fully and then saved in
one call, so a single output is not written half-formed, but a batch that fails partway
leaves the outputs that already committed in place. The per-item outcomes are how you
find out which.

## Compare two Word documents

```csharp
var comparison = await client.CompareDocumentsAsync(original, revised,
    new DocumentComparisonOptions
    {
        Revision = new RevisionMetadata { Author = "Contract Comparison" }
    });

if (comparison.IsComplete && comparison.Plan is not null)
{
    var preview = await client.PreviewAsync(original, comparison.Plan);
    if (preview.IsValid)
        await client.CommitAsync(original, comparison.Plan,
            new SaveDocumentOptions { Mode = SaveMode.NewDocument, NewName = "redline.docx" });
}
```

Comparison itself is read-only. Its result includes exact SHA-256 hashes of both inputs,
structured added/removed/changed paragraph records, diagnostics, and a plan bound to the
original snapshot. Applying that plan creates real Word insertions and deletions: accepting
all revisions produces the revised body text and rejecting all revisions restores the
original body text. Preview the plan against the original immediately before commit so drift
fails as `stale-snapshot`.

The v0.8 comparison scope is deliberately bounded to free Word body-paragraph text. Existing
revisions and changes in tables, images or image bytes, headers, footers, footnotes, endnotes,
run structure, direct formatting, paragraph styles, fields, properties, comments, sections,
relationships, or other package parts make the result incomplete and suppress the plan. The
comparison uses structural checks plus an exact fingerprint of every package part after
normalizing the supported free-paragraph text and volatile paragraph identifiers. This is a
strict coverage gate: package metadata changes can suppress a plan even when they do not affect
the visible document. Always require `IsComplete` and a non-null `Plan`; the exact input hashes
identify the compared files but do not by themselves prove semantic coverage. Default limits
are 64 MiB per input, 2,000 body paragraphs per document, and 1,000 reported differences.
Cancellation is checked while reading and while computing alignment.

The equivalent MCP or Agent Framework call is `compare_documents` with the original and
revised `(connectionId, documentId)` pairs plus the displayed `revisionAuthor`. A hosted
access policy must grant read access to both connections. The tool never writes; pass its
plan to normal preview and apply tools against the original.

Run the [document-comparison sample](../samples/DocumentComparison/) with two `.docx` files.

## Evidence boundary

The samples and automated checks prove contract behavior, round trips, stale-snapshot
rejection, resource bounds, and access-policy enforcement. They do not measure whether a
specific organization can adopt the workflows or whether the output semantics fit its
documents; that requires trials with representative templates and document pairs.
