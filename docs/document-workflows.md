# Template population and document comparison

To combine whole documents in a specified order, see [Word document assembly](document-assembly.md).

OfficeAgent 0.8 adds two higher-level workflows built on the same inspect, plan, preview,
commit, provider, and receipt contracts as ordinary edits. Template population creates
independent documents from tagged values and repeating Word rows. Comparison reads two Word
documents and proposes a native tracked-change plan when it can account for every change in
its supported scope.

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

The v0.8 comparison scope is deliberately bounded to free Word body paragraphs. Existing
revisions and inspected changes in tables, images, headers, footers, footnotes, endnotes,
paragraph styles, fields, properties, comments, or sections make the result incomplete and
suppress the plan. A byte change with no covered body-text difference is also reported as an
unsupported package change. Opaque package metadata changed alongside covered paragraph text
cannot always be distinguished by structural inspection; compare the returned exact hashes
when that distinction matters. Default limits are 64 MiB per input, 2,000 body paragraphs per
document, and 1,000 reported differences. Cancellation is checked while reading and while
computing alignment.

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
