# Operations guide

Practical answers to "how do I run OfficeAgent.NET in production?"

## Thread safety

`OfficeAgentClient`, the engine behind it, the format modules, and registered `IDocumentProvider` implementations are safe for concurrent use. Inspect, find, and preview are pure reads; apply opens a fresh in-memory editable package per call, so concurrent edits on *different* documents never share mutable state. **Register `OfficeAgentClient` as a singleton** in your DI container:

```csharp
services
    .AddWordFormat()
    .AddFileSystemDocumentProvider("workspace", "/srv/officeagent/workspace")
    .AddOfficeAgent();
```

Two concurrent `Commit` calls on the *same* document id are safe at the engine level - each opens its own in-memory copy - but they are not coordinated at the provider boundary, so the last successful save wins. If write ordering matters, either pass each commit's source `Version` via `SaveDocumentOptions.ExpectedVersion` (the optimistic-concurrency check rejects stale saves with `DocumentVersionConflictException`) or serialise the commits in your application layer (`SemaphoreSlim` keyed by id).

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

## Output paths

`ApplyResult` exposes the committed bytes three ways:

```csharp
byte[] bytes = result.ToBytes();          // for in-memory consumers
result.Save("contract.updated.docx");      // synchronous file write
await result.SaveAsync("contract.updated.docx", ct);   // async file write
```

All three throw `InvalidOperationException` when the plan was a dry run or did not commit. Always check `result.Committed` and inspect `result.Report.Errors` first.

## Failure modes you should handle

| Symptom | Likely cause | What to do |
|---|---|---|
| `ApplyResult.Committed == false`, `Errors` contains `stale-snapshot` | Document was edited after the inspect that produced the plan | Re-inspect and rebuild the plan against the fresh snapshot |
| `Errors` contains `expect-mismatch` for a `changeText` | Paragraph text drifted from the anchor's `Expect` | Re-find or re-inspect the paragraph; rebuild that operation |
| `Errors` contains `requires-renderer` | The plan asked for a field-recalc or pagination value the OOXML engine cannot compute | Use `setProperty` with `updateOnOpen` to defer to Word, or move that work to a renderer |
| Empty `Find` result before building a plan | Target text not present at all | Surface to the user; do not build an operation against a missing anchor - the source document is never modified |

## Versioning

`DocumentPlan.CurrentContractVersion` advertises the plan-contract version the engine speaks. Pre-1.0 the field is informational; a mismatch does not fail the plan. Pin the package version in production and re-test on upgrade.
