# Open XML SDK interoperability

OfficeAgent.NET uses `DocumentFormat.OpenXml` internally, so a .NET application can
combine OfficeAgent operations with direct Open XML SDK edits. This page defines the
boundary that makes that safe: which guarantees apply at each stage, which do not
carry across an SDK edit, and which workflow is tested.

The short rule: OfficeAgent guarantees apply to OfficeAgent operations. An SDK edit
is an external mutation. It gets a fresh inspection and a fresh plan, never a reused
anchor, snapshot, or approved preview.

## When to reach for the SDK

Use an OfficeAgent operation whenever one exists. Reach for the SDK only for
low-level part or schema work OfficeAgent does not implement, and own the result:
the SDK does not guarantee package validity, and it does not provide Word layout or
Excel recalculation. See the [library selection guide](choose-officeagent.md#officeagentnet-or-the-open-xml-sdk-directly)
for the abstraction-level comparison.

## The supported workflow

Copy, edit the copy, validate it, reinspect it, then plan again. The registered
source is never opened for writing by the SDK.

1. **Snapshot.** Open the registered document through its provider and copy the bytes
   out. `DocumentContent.Reference.Version` is the provider version those bytes were
   read at.
2. **Edit a separate output.** Write the copied bytes to a new file and edit that file
   with the SDK. Nothing hands the SDK a live OfficeAgent package.
3. **Validate the output.** Run `OpenXmlValidator` against the edited file before
   OfficeAgent is asked to work with it.
4. **Reinspect and replan.** Register the output, inspect it, and build a new plan from
   the new inspection's snapshot and anchors.

```csharp
var source = await client.RegisterAsync("workspace", sourcePath);

// 1. Authorized snapshot.
byte[] snapshotBytes;
using (var content = await client.OpenReadAsync(source))
using (var buffer = new MemoryStream())
{
    await content.Stream.CopyToAsync(buffer);
    snapshotBytes = buffer.ToArray();
}

// 2. SDK edit, on a separate output only.
await File.WriteAllBytesAsync(outputPath, snapshotBytes);
using (var document = WordprocessingDocument.Open(outputPath, isEditable: true))
{
    document.MainDocumentPart!.Document!.Body!
        .AppendChild(new Paragraph(new Run(new Text("Appendix A is incorporated by reference."))));
    document.MainDocumentPart.Document.Save();
}

// 3. Validate the SDK output.
using (var document = WordprocessingDocument.Open(outputPath, isEditable: false))
{
    var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(document).ToList();
    if (errors.Count > 0) throw new InvalidOperationException("SDK output failed schema validation.");
}

// 4. Reinspect, then author a fresh plan from the fresh inspection.
var output = await client.RegisterAsync("workspace", outputPath);
var inspection = await client.InspectAsync(output);
var hit = (await client.FindAsync(output, new FindQuery("Invoices are due in 30 days."))).Single();

var plan = new DocumentPlan
{
    Snapshot = inspection.Snapshot,
    Operations = new PlanOperation[]
    {
        new ChangeTextOp { Target = hit.Anchor, With = "Invoices are due in 45 days.", Mode = ChangeMode.Tracked }
    }
};
```

## What is guaranteed at each stage

| Stage | Validation | Authorization | Concurrency | Receipt |
| --- | --- | --- | --- | --- |
| 1. Provider snapshot | The package is read as bytes; no schema check is implied | The provider connection's configured root and access rules apply | The returned `DocumentReference.Version` pins the bytes that were read | None; reading produces no receipt |
| 2. SDK edit of the output | None from OfficeAgent; the SDK does not guarantee validity | None; this is your application writing a file it owns | None; OfficeAgent is not involved and cannot detect an interleaved writer | None; an SDK edit has no OfficeAgent receipt |
| 3. `OpenXmlValidator` on the output | Office schema conformance only | Not applicable | Not applicable | Not applicable |
| 4. Register, inspect, preview, commit | Full plan validation: anchors, snapshot, operation support, format | The provider connection's access rules apply to the output document | Plan `Snapshot` guards text-host drift; `SaveDocumentOptions.ExpectedVersion` and the reference version guard the write | `ApplyResult.Receipt` covers the OfficeAgent operations only |

## What does not carry across an SDK edit

- **Anchors and snapshots.** Anchor ids and `SnapshotToken` values from before the SDK
  edit are not valid after it. Reinspect and rebuild.
- **An approved preview.** A `ChangeReport` describes the document it was previewed
  against. It does not authorize a commit against different bytes.
- **Receipts and audit coverage.** `ApplyReceipt` records OfficeAgent operations. An SDK
  edit produces no receipt and appears in no OfficeAgent audit trail. If an SDK mutation
  must be auditable, express it as a supported OfficeAgent operation instead.
- **Preservation evidence.** The [Word preservation matrix](word-preservation-evidence.md)
  measures OfficeAgent operations. It says nothing about parts your SDK code rewrote.

Schema validation is also not a semantic guarantee. A package with zero
`OpenXmlValidator` errors can still lay out wrongly in Word or carry tracked-change
markup Word will not accept or reject as you intended. Open the result in the Office
application your reviewers use before adopting the workflow.

## Two different staleness guards

External change detection and plan drift detection are not the same mechanism, and
only one of them sees every SDK edit.

- **Provider version** (`DocumentReference.Version`, `SaveDocumentOptions.ExpectedVersion`)
  is derived from the complete bytes. Any SDK edit changes it. A reference pinned to an
  older version cannot even be opened, and a save with a stale `ExpectedVersion` fails
  with a version conflict before anything is written. This is the guard to build on.
- **Plan snapshot** (`DocumentPlan.Snapshot`) is derived from the document's text hosts.
  For Word that is the body, headers, footers, footnotes, and endnotes. An SDK edit
  confined to another part, for example package properties or the styles part, leaves
  the snapshot etag unchanged, so the plan is not refused as stale.

A plan snapshot that still matches is therefore evidence that the text hosts have not
drifted. It is not evidence that no external mutation occurred. Keep the provider
version on the reference, or pass `ExpectedVersion` on the save, whenever another
writer can reach the document.

```csharp
// After an external SDK write, the pinned reference is refused outright.
await Assert.ThrowsAsync<DocumentVersionConflictException>(
    () => client.PreviewAsync(registeredReference, planAuthoredEarlier));
```

Dropping the version from a reference discards that guard. Do it only when your
application has deliberately accepted the external change, and reinspect immediately
afterwards.

## What is not supported

- **No live-package callback.** There is no public API that hands your code the open
  `OpenXmlPackage` mid-operation. `IOpenXmlPackage` is the engine's internal handle.
  Use the separate-output workflow above.
- **No in-place SDK replacement of a registered document through OfficeAgent.** To
  publish SDK-edited bytes into the source document, write them through the provider
  yourself with the current version in hand, and treat the result as a new document
  to register and inspect. OfficeAgent does not offer a "commit these external bytes"
  path, because it could not validate or receipt them.
- **No arbitrary execution tool.** The agent and MCP surface is a fixed list of
  document operations. No tool accepts SDK code, a package callback, or an arbitrary
  expression, so agent-driven use cannot bypass provider authorization or host limits.
  See [agent integration](agent-integration.md) for the exposed tool set.

## Run it yourself

The workflow above is recipe 4 in the installable integration kit and runs against
installed packages in a disposable directory. See
[installed-package recipes](../skills/officeagent-integration/references/recipes.md#recipe-4-open-xml-sdk-interoperability).
The repository additionally proves the boundaries in
`tests/OfficeAgent.Tests/SdkInteroperabilityTests.cs`: the source stays byte-identical,
a plan held across an SDK text edit is refused as stale, a stale provider version cannot
overwrite a changed document, the snapshot's text-host limit is explicit, and the agent
tool surface exposes no execution tool.
