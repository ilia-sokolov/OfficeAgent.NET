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
documentId, requestJson, expectedTokenJson)`. The host must opt in with `AllowCreation`; its
connection access policy must also grant read and create. For example:

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

### Binding images and charts

Scalar values are strings. An image or a native chart is bound through `TypedValues`,
which sits beside `Values` rather than replacing it: a scalar-only request behaves
exactly as it always did.

```csharp
var binding = new TemplateBinding
{
    Values = new Dictionary<string, string?> { ["CustomerName"] = "Fabrikam Services" },
    TypedValues = new Dictionary<string, TemplateValue>
    {
        ["Photo"] = new TemplateImageValue
        {
            Base64Bytes = photoBase64,
            ImageType = "png",
            WidthPx = 120,
            HeightPx = 120,
            AltText = "A product photograph"
        }
    }
};
```

Ask discovery which slots can hold what before binding:

```csharp
foreach (var slot in (await client.DiscoverTemplateAsync(template)).MediaSlots)
    Console.WriteLine($"{slot.Name} holds {slot.MediaKind}");
```

#### Where the bytes may come from

Exactly one of two places, and never the open internet:

- **Inline**, as base64 the caller already holds.
- **A provider document** the caller previously registered, named by `ImageConnectionId`
  and `ImageDocumentId`. It is read through the provider, so that connection's access
  rules and size limits apply to the image exactly as they would to any other document.

There is no URL form. Fetching one would make the engine a client of whatever address a
plan happened to carry, which is not a decision a template binding should be able to
make.

#### What is checked before anything is written

| Check | Refusal |
| --- | --- |
| The slot can hold this kind of value | `wrong-slot-kind` |
| The slot is not inside a repeating table row | `media-in-repeating-row` |
| The image carries alt text | `missing-alt-text` |
| The bytes match the declared image type | `image-type-mismatch` |
| Exactly one byte source is given | `invalid-image-binding` |
| The image fits the budget | `image-too-large` |
| The same slot is not bound twice | `duplicate-template-binding` |
| A chart binding targets a native chart in a deck | `wrong-slot-kind` |
| A chart binding is not aimed at a Word template | `unsupported-template-feature` |

The type check compares the declared type against the actual magic bytes, so arbitrary
content cannot be labelled a PNG and embedded under that name.

#### Charts are PowerPoint only

A chart binding updates a native chart that is **already in the template**; it does not
create one. This engine has no native Word chart handling, so a Word template reports no
chart slots and a chart binding against one is refused with
`unsupported-template-feature`. That absence is the honest answer rather than an
omission.

#### Media in repeating rows

Not supported. A slot inside a table row is reported with `InRepeatingRow` true, and a
binding that targets one is refused: every generated row would otherwise share a single
image. Move the media slot outside the table.

#### The token covers the resolved bytes

`TemplateBatchToken.MediaSha256` hashes the payloads the bindings actually resolved to.
The batch hash covers the binding as written, which for a provider-held image is only its
id, so replacing the image behind an unchanged reference would not move it. The media
hash does, and a commit carrying the older token is refused.

Media is resolved during preview, not at commit, because the exact bytes have to be in
hand before the token can bind to them.

#### Media budgets

| Budget | Default |
| --- | --- |
| `MaximumImageBytes` | 8 MiB |
| `MaximumImagesPerDocument` | 20 |
| `MaximumTotalImageBytes` | 64 MiB |
| `MaximumChartPoints` | 1,000 |

Set through `TemplateBatchLimits.Media`. As with every other budget here, a request may
ask for something stricter and can never raise a host ceiling.

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

### Preflighting a batch from an agent

An agent does not have to guess a template's tag names or discover a bad batch one output at
a time. Three tools cover the workflow, and the same contract is on the direct client as
`DiscoverTemplateAsync`, `PreviewTemplateBatchAsync` and `PopulateTemplateBatchAsync`.

1. `discover_template(connectionId, documentId)` reports the scalar slots, the image and
   native-chart media slots, and the repeating Word rows the template exposes, plus
   diagnostics for anything ambiguous such as one tag used twice. It writes nothing.
2. `preview_template_batch(connectionId, documentId, requestJson)` validates the entire
   batch and writes nothing. Every item is validated even after an earlier one fails, so one
   call reports all the problems rather than the first, and each item reports the operation,
   row, image and byte counts it would produce.
3. `populate_template_batch(connectionId, documentId, requestJson, expectedTokenJson)`
   commits. Passing the preview's `token` refuses the commit if the template bytes or the
   normalized batch changed after the review, including the bytes a provider-held image id
   resolved to. Send `""` to commit without that binding.

The batch is preflighted internally whether or not step 2 was called, so a batch that cannot
validate writes nothing rather than leaving part of itself in storage. The token adds the
separate guarantee that what commits is what was reviewed.

A stale token is reported as `stale-batch-preview` and writes nothing. The repair is to
preview the batch you actually intend to commit and use the token that preview returns.

### What the comparison understands

The comparison covers free body paragraphs and, since it learned about table cells, the
words inside a table whose shape has not moved.

| Difference | Result |
| --- | --- |
| Body paragraph text added, removed or changed | Tracked revision in the plan |
| The same text split differently across equally formatted runs | **Not a difference.** No finding, no refusal |
| Cell text changed with the table's geometry unchanged | Tracked revision in the plan |
| Cell text changed together with its formatting | Refused, `unsupported-table-markup-change` |
| Table rows, columns, spans or nesting changed | Refused, `unsupported-table-change` |
| Paragraph formatting or style changed | Refused, `unsupported-paragraph-markup-change` or `unsupported-style-change` |
| Headers, footers, notes, images, numbering, styles, other parts changed | Refused, with that area's own code |

#### Run segmentation is not content

Word re-segments runs constantly. Typing in the middle of a sentence, a spell-check pass,
or a round trip through another editor can turn one run into three carrying identical
formatting and identical words. The comparison merges adjacent runs that hold only text
and share the same run properties before comparing structure, so that stops being a
difference.

Only text-carrying runs merge. A run holding a break, a tab, a field, a drawing or a note
reference keeps its boundary, because those are content: merging them would change the
document rather than normalise it. Whitespace is likewise text - a paragraph whose spacing
changed is a real difference and is reported as one.

#### Cells align by position, only when the shape matches

Cell alignment is positional, which is sound **only** because it runs after the table's
geometry has been confirmed identical. The geometry check hashes the tables with cell
words stripped, so it answers "did the shape move" rather than "did anything change". If
it moved, the cells in position three are not the same cell in both documents, and the
plan is withheld before any cell text is trusted.

A cell whose formatting changed alongside its words is refused rather than flattened:
reproducing it as plain text would lose the formatting the author was editing.

### Reading comparison coverage

A comparison answers two separate questions: what changed in the body, and whether it can
offer a redline for it. `Coverage` reports the second one area by area.

| State | Meaning |
| --- | --- |
| `Unchanged` | Compared, and the same in both documents |
| `Changed` | Compared, differs, and represented in the plan |
| `Blocked` | Differs, and this comparison cannot represent the difference |
| `NotCompared` | Not examined. **Nothing is known about it** |

`NotCompared` is not a softer `Unchanged`. A comparison that stopped at a precondition -
no anchor paragraph, too many paragraphs, existing tracked revisions - reports every area
`NotCompared`, because it never looked. Treating that as "matches" is the inference this
result refuses to make for you.

#### Findings survive a blocked area

An unsupported change no longer discards the supported ones. A batch with an edited
paragraph and a replaced image reports the paragraph difference **and** names the image
area as blocked:

```csharp
var comparison = await client.CompareDocumentsAsync(original, revised);

foreach (var difference in comparison.Differences)
    Console.WriteLine($"{difference.Before} -> {difference.After}");   // still reported

foreach (var area in comparison.Coverage.Where(a => a.State == ComparisonAreaState.Blocked))
    Console.WriteLine($"{area.Name}: {area.Code}");                     // images: unsupported-image-change

Console.WriteLine(comparison.Plan is null);                             // True
```

You can act on the findings by hand. What you cannot do is apply a plan, because there
isn't one.

#### Which area blocked it

Each area has its own code, so a refusal says what to look at:

| Area | Code |
| --- | --- |
| `tables` | `unsupported-table-change` |
| `images` | `unsupported-image-change` |
| `notes` | `unsupported-note-change` |
| `headersAndFooters` | `unsupported-header-footer-change` |
| `styles` | `unsupported-style-definition-change` |
| `numbering` | `unsupported-numbering-change` |
| `otherParts` | `unsupported-package-change` |

Body-level reasons keep their existing codes: `unsupported-style-change`,
`unsupported-paragraph-markup-change`, `unsupported-existing-revisions`,
`comparison-anchor-unavailable`, `comparison-paragraph-limit-exceeded` and
`comparison-difference-limit-exceeded`.

#### There is no partial plan

`IsComplete` is false whenever any diagnostic exists, and `Plan` is null whenever
`IsComplete` is false. Hitting a resource ceiling cannot be traded for a plan either: a
truncated finding list is visible in the diagnostics and the body area is reported
`Blocked`, so a shortened list is never mistaken for a complete one.

The agent and MCP tools report the same thing. An incomplete comparison serializes with
no plan at all, so there is no adapter-shaped way around this.

## Evidence boundary

The samples and automated checks prove contract behavior, round trips, stale-snapshot
rejection, resource bounds, and access-policy enforcement. They do not measure whether a
specific organization can adopt the workflows or whether the output semantics fit its
documents; that requires trials with representative templates and document pairs.
