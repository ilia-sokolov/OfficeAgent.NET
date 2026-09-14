# Operations guide

Practical answers to "how do I run OfficeAgent.NET in production?"

## Thread safety

`OfficeAgentClient`, the engine behind it, the format modules, and registered `IDocumentProvider` implementations are safe for concurrent use. Inspect, find, and preview are pure reads; apply opens a fresh in-memory editable package per call, so concurrent edits on *different* documents never share mutable state. **Register `OfficeAgentClient` as a singleton** in your DI container:

```csharp
services
    .AddWordFormat()
    .AddPowerPointFormat()
    .AddFileSystemDocumentProvider("workspace", "/srv/officeagent/workspace")
    .AddOfficeAgent();
```

Two concurrent `Commit` calls on the *same* document id are safe at the engine level - each opens its own in-memory copy. Under the default `Replace` save mode they are also checked at the provider boundary: the commit carries the version read when it opened the document, and a save whose source changed in between is rejected with `DocumentVersionConflictException` rather than overwriting the other writer. Handle that exception by re-inspecting and re-authoring the plan.

The filesystem provider holds a gate per document id across the version check and the publish, so that check and the write it guards are one step. Without it two callers that read the same version would both pass the check and both write, and the first caller's edit would disappear while it was told the save succeeded. The gate is keyed by item, so contested writes to one document serialise while writes to other documents continue in parallel.

**That gate is in-process.** It orders callers inside one `OfficeAgentClient`; it does nothing about another process, another machine, or someone editing the file in Word. Those are still caught by the version check, which compares content, but they are caught rather than ordered - the loser receives a conflict. For ordering across processes, use storage that provides it, or coordinate above OfficeAgent.

The version check does nothing for `NewVersion`/`NewDocument`, which write to a fresh name. If you want to fail before doing the work instead of after, pass an explicit `SaveDocumentOptions.ExpectedVersion`.

## Stream and lifetime ownership

- `StreamHandle` - the engine **copies** your stream into an internal `MemoryStream` and never disposes the source. It seeks to position 0 on a seekable source before reading, so the caller's stream position may be modified.
- `DocumentContent` (returned from `OpenReadAsync`) implements `IDisposable`. The stream is owned by the engine-allocated buffer; dispose it (or `using var content = ...`) to release.
- `ApplyResult` (returned from `Commit` with a `DocumentHandle`) implements `IDisposable` and exposes `ToBytes()` / `Save(path)` / `SaveAsync(path)`.
- `ProviderApplyResult` (returned from `CommitAsync` with a `DocumentReference`) carries the saved `Document` reference; the bytes live in storage and are reached via `OpenReadAsync` when needed.
- Agent tool `apply_plan` results follow the same model: they return `outputDocumentId`, `outputName`, and `outputContentType`. The host resolves the id with `OpenReadAsync` and delivers the stream through its native file/attachment API; the LLM should not echo base64 document bytes.

## Memory model

During `Apply` the engine holds, at peak, ~3 copies of the document in memory: the source bytes, the validation package's parsed Open XML DOM, and the commit package's parsed Open XML DOM. For a 10 MB document that's roughly 30–40 MB of allocations. The serialised output is a fresh `MemoryStream` returned through `ApplyResult.Output`.

Practical limits:
- Documents up to a few tens of MB: fine on a typical server.
- Documents larger than 100 MB: expect noticeable per-call memory pressure and consider hosting the engine in a dedicated process or batching reads with backpressure.

`Inspect`, `Find`, and `Preview` hold one parsed DOM. They are noticeably cheaper than `Apply`.

## Async surface and cancellation granularity

Every public method has an async overload that accepts a `CancellationToken`:

```csharp
await client.CommitAsync(handle, plan, ct);
```

File IO uses `File.ReadAllBytesAsync` / `File.WriteAllBytesAsync` on .NET 8+ and falls back to `Task.Run` on `netstandard2.0` so the token is observed at the boundary either way. The engine checks the token between operations, so cancelling a long multi-op plan stops at the next operation. **Cancellation is not propagated inside a single handler invocation** - inspecting a 100k-paragraph document or applying a single very large `ChangeTextOp` will run to completion regardless. If you need finer granularity, slice the work into smaller plans.

## Logging and telemetry

OfficeAgent uses standard `Microsoft.Extensions.Logging` and `System.Diagnostics.ActivitySource` primitives. Both are quiet by default.

- **Logger category:** `OfficeAgent` (see `OfficeAgentTelemetry.LogCategory`).
- **Activity source:** `OfficeAgent` (see `OfficeAgentTelemetry.ActivitySourceName`).

Wire them up at startup:

```csharp
services.AddLogging(b => b.AddConsole());

// OpenTelemetry consumers:
builder.Services.AddOpenTelemetry()
    .WithTracing(t => t.AddSource(OfficeAgentTelemetry.ActivitySourceName));
```

What you get:
- `OfficeAgent.Inspect` / `Find` / `Apply` activities with byte counts and durations.
- `Information` log lines for committed applies (op count, output size, elapsed).
- `Warning` log lines when apply aborts mid-plan (with the failing code and message).
- `Debug` log lines for inspect, find, and dry-run apply.

Treat logs as application data. Information events contain connection ids,
opaque item ids, and document names; debug events can contain search patterns or
validation messages derived from document content. Registration logging does not
record the raw source path or SharePoint URL. Restrict log access and retention,
and avoid debug logging in production unless it is needed for a bounded incident.

## Registration-state durability

Filesystem connections persist registrations in
`{root}/.officeagent/index.json`. SharePoint uses memory unless
`RegistrationIndexPath` is configured. A JSON registration index is sensitive
application state: restrict its ACL, back it up with the referenced documents,
and let only one process write it. For a multi-instance deployment, implement
`ISharePointRegistrationStore` over shared storage with appropriate concurrency
control.

## Filesystem root trust boundary

The filesystem provider rejects traversal and checks every existing component
for a symlink or reparse point at registration, open, and save. These checks
protect against caller-supplied link paths and links introduced between calls.
They cannot make a pathname check and the later read/write one atomic operation
against a separate process racing a directory rename.

Enforce the remaining boundary with operating-system ACLs: only the OfficeAgent
service identity and trusted administrators may create, delete, rename, or
replace entries below a connection root. Grant a human who must edit an existing
document file-write permission without directory mutation rights. Do not mount
the same container volume into an untrusted workload. For untrusted uploads,
copy validated bytes into the root through the trusted host.

## Provider retries

OfficeAgent does not automatically retry Microsoft Graph requests or interpret
`Retry-After`. Reads can be retried with normal bounded backoff. A timed-out or
throttled write is ambiguous: first reopen the existing id, or look up the
requested new name, to determine whether Graph committed it. If the source
changed, re-inspect and rebuild the plan; do not replay a stale plan blindly.

## Output paths

`ApplyResult` exposes the committed bytes three ways:

```csharp
byte[] bytes = result.ToBytes();          // for in-memory consumers
result.Save("contract.updated.docx");      // synchronous file write
await result.SaveAsync("contract.updated.docx", ct);   // async file write
```

All three throw `InvalidOperationException` when the plan was a dry run or did not commit. Always check `result.Committed` and inspect `result.Report.Errors` first.

## Revision identity and audit receipts

Every result produced by the built-in engine's `Apply` or `PreviewWithReceiptAsync` carries
an `ApplyReceipt`; provider-backed results carry the same receipt and add the saved
`OutputDocument` reference after storage accepts the write. The receipt records:

- `Outcome`: `Previewed`, `Rejected`, or `Committed`;
- one trusted engine `TimestampUtc` for the attempt;
- the resolved revision author and UTC timestamp;
- lowercase SHA-256 hashes of the effective plan JSON and exact input bytes;
- the output-byte SHA-256 only when the engine committed;
- an optional authenticated `Actor` supplied by the host;
- the provider output reference after a successful provider save.

The hashes make accidental or later changes detectable; the receipt is not signed. Persist
it in an append-only or otherwise protected host audit store if it must serve as durable
evidence. A provider exception can happen after storage accepted a write, so the call may
throw before a final provider receipt is returned; reconcile the destination before retrying.

Keep the two identities separate. `DocumentPlan.Revision.Author` controls what a reviewer
sees in Word and is untrusted plan content. `AuditActor` must come from authentication state:

```csharp
public sealed class RequestActorProvider : IAuditActorProvider
{
    public AuditActor? GetCurrentActor() => new()
    {
        Subject = CurrentPrincipal.Subject,
        Issuer = CurrentPrincipal.Issuer,
        DisplayName = CurrentPrincipal.DisplayName
    };
}

services.AddSingleton<IAuditActorProvider, RequestActorProvider>();
```

For a direct in-memory call, pass the actor with
`new ApplyOptions { DryRun = false, Actor = actor }`. For one provider save, an actor on
`SaveDocumentOptions` takes precedence over the registered provider.

## Failure modes you should handle

| Symptom | Likely cause | What to do |
|---|---|---|
| `ApplyResult.Committed == false`, `Errors` contains `stale-snapshot` | Document was edited after the inspect that produced the plan | Re-inspect and rebuild the plan against the fresh snapshot |
| `Errors` contains `expect-mismatch` for a `changeText` | Paragraph text drifted from the anchor's `Expect` | Re-find or re-inspect the paragraph; rebuild that operation |
| `Errors` contains `requires-renderer` | The plan asked for a field-recalc or pagination value the OOXML engine cannot compute | Use `setProperty` with `updateOnOpen` to defer to Word (only on a document that has fields), or move that work to a renderer |
| `Errors` contains `invalid-operation` for `updateOnOpen` | The document has no fields, so the setting would only make Word prompt about updating fields that do not exist | Insert the field first, or drop the operation |
| Empty `Find` result before building a plan | Target text not present at all | Surface to the user; do not build an operation against a missing anchor - the source document is never modified |
| `DocumentVersionConflictException` from a `Replace` save | Another writer changed the document between this commit's open and its save | Re-inspect and re-author the plan against the current bytes; nothing was overwritten |
| `Errors` contains `unsupported-operation` on a deck | The verb is Word-only - the PowerPoint module implements a subset | Use a verb the deck supports, or record the intent as a comment. See [PowerPoint support](powerpoint.md) |
| `Errors` contains `invalid-operation` naming mode `Tracked` on a deck | PresentationML has no redline vocabulary | Send `"mode": "Direct"`, or set the connection's `DefaultChangeMode` so deck plans need not restate it |
| MCP server exits during startup | An invalid `AuthMode`, unreadable config file, or incomplete provider configuration | Read the named configuration error and restart. No storage configuration by itself starts the `session` connection; invalid authentication modes never fall back to app-only. |
| Graph returns 429 or a transient 5xx | Throttling or a service interruption | Respect `Retry-After` in the host; reconcile an attempted write before retrying |
| A create call reports an I/O error but the name is now occupied | Storage accepted the file before registration or the response failed | Inspect the destination and register the surviving item; do not overwrite or blindly retry the name |

## What a failed commit tells you about storage

Three outcomes are distinguishable, and only the first lets you conclude that nothing was written.

| Outcome | How you know | What to do |
|---|---|---|
| **Nothing was written** | `ApplyResult.Committed` is `false` with validation `Errors`, or a `DocumentVersionConflictException` from the pre-save check | The plan never reached storage. Fix the plan or re-inspect, then retry safely |
| **The write landed** | The call returned and `ProviderApplyResult.Committed` is `true` with a `Document` reference | Use the returned reference. The receipt names the output |
| **Unknown** | A provider or cancellation error raised *after* the content was handed to storage | Do not retry blindly and do not report the document as unchanged. Inspect the destination, reconcile what is actually there, and report the possibly-written name to the operator |

The uncertainty in the third row is real and cannot be engineered away at this layer: a provider that accepts bytes and then fails to confirm leaves a write this library cannot see. OfficeAgent does not claim exactly-once delivery. What it does guarantee is that plan validation failures are always the first row - they are decided entirely in memory, before any provider write.

## Versioning

`DocumentPlan.CurrentContractVersion` is the enforced edit-plan wire version. In OfficeAgent.NET
0.9, omit `contractVersion` for legacy `0.2` behavior or set it explicitly to `"0.2"`.
Null, empty, malformed, and unknown values fail with `contract-mismatch` before document
processing or saving. Package version `0.9.0`, merge-plan version `1`, and receipt version
`1` are separate contracts. Pin the package version in production and re-test on upgrade.
