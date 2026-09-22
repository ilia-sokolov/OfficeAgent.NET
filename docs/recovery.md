# Storage outcomes and recovery

What a failed call did to storage, what each workflow protects, and how to recover without
losing or duplicating a document. Every row below is exercised by an automated test; the tests
are named in [the fault matrix](#the-fault-matrix).

## Four outcomes

Every storage-writing call ends in one of four outcomes. Only the first lets a caller conclude
that nothing was written.

| Outcome | Tool `writeOutcome` | .NET | Meaning |
| --- | --- | --- | --- |
| Not written | `notWritten` | a validation report, `DocumentProviderException` with a pre-write code, or `OperationCanceledException` | Nothing reached storage, or storage proved it kept nothing |
| Committed | `committed` | `ProviderApplyResult.Committed` | Storage confirmed the write; the receipt names the output |
| Written, not registered | `writtenNotRegistered` | code `RegistrationFailed` | The document exists but has no document id |
| Unknown | `unknown` | `DocumentWriteOutcomeUnknownException`, code `OutcomeUnknown` | The write may or may not have been stored |

A `DocumentProviderException` raised inside a provider's save or create proves nothing was
written only with one of these codes: `NotFound`, `AccessDenied`, `ContentTooLarge`,
`ExtensionNotAllowed`, `VersionConflict`, `InvalidArgument`, `ConfigurationError`,
`AlreadyExists` or `WriteRejected`. The engine reports anything else raised once a write call
has begun as unknown, because it cannot rule the write out. That includes `IO`, a transport
error, a timeout and a cancellation. A custom provider that can guarantee nothing was stored
should say so with `WriteRejected`.

`committed: false` in a tool response means "not confirmed", not "nothing was written". Read
`writeOutcome`.

## Every boundary, and what to do

| Boundary | Reported as | Safe next action |
| --- | --- | --- |
| Validation fails | `notWritten`, plan `errors` | Fix the plan, or re-inspect and rebuild it |
| Access revoked before commit | `notWritten`, `connection-forbidden` | Stop; the caller no longer holds the capability |
| Source changed since preview | `notWritten`, `version-conflict`, `stale-snapshot`, `expect-mismatch`, `stale-batch-preview` or `stale-merge-source` | Re-inspect and re-preview; never replay the old plan |
| Storage refused the write | `notWritten`, `write-rejected`, `already-exists` or another pre-write code | Fix the cause (a lock, a taken name, a size) and retry |
| Storage committed the write | `committed`, with a receipt | Persist the receipt |
| Written, but registration failed | `writtenNotRegistered`, `registration-failed` | Do not write it again; register the named output |
| Receipt persistence failed after storage | `committed` (the engine returned the receipt; persisting it is the host's) | Keep the write; retry persisting the receipt from the result |
| Cancellation after the write began | `unknown`, `outcome-unknown` | Reconcile before retrying |
| Unknown remote outcome (timeout, lost response, server error) | `unknown`, `outcome-unknown` | Reconcile before retrying |

A cancellation that arrives before any write begins is plain cancellation, reported as
`cancelled` by the tools, and nothing was written. A template batch cancelled after its first
item returns its partial result instead of throwing: committed items keep their receipts, and
every item not yet attempted is `Skipped`.

## Reconciling an unknown outcome

An unknown outcome carries a locator, `possibleOutput` in tool responses and the exception's
members in .NET. It holds only what the caller supplied: the connection, the source document id,
the output name the caller asked for, and the SHA-256 of the bytes the write carried. It never
contains a storage path, and it does not imply the output was registered.

1. Find the destination. A `Replace` wrote to `sourceDocumentId`. A new document or version was
   written under `outputName`, or under a versioned name beside the source when no name was given.
2. Read it and hash its bytes: `export_document_content`, or `OpenReadAsync` in .NET.
3. Compare with `expectedSha256`:
   - equal: the write landed. Treat it as committed and register the output if it has no id.
   - the source's prior content: it did not land. Re-preview, then retry.
   - anything else: another writer intervened. Re-inspect before doing anything.

## Retries and idempotency

No built-in provider stores a durable request identity, so **no workflow is idempotent** and
none is retried automatically. What the providers do enforce makes a careful retry safe:

- **New documents and new versions are conditional creates.** The filesystem publishes with a
  no-overwrite rename, the memory store checks names under its lock, and SharePoint uploads with
  `conflictBehavior=fail`. Retrying with the *same name* after an unknown outcome cannot
  duplicate or overwrite: if the first attempt landed, the retry fails `already-exists`.
- **Replace is conditional on the version you read.** Set `SaveDocumentOptions.ExpectedVersion`
  to the version from your inspection. A retry then fails `version-conflict` if the first attempt
  landed, instead of applying the plan twice. Without it, a retry reopens the document and uses
  its current version.

## What each workflow protects

"Version" is the provider's version token: the content SHA-256 for filesystem and memory, and
the eTag for SharePoint.

| Workflow | Package-byte hash | Semantic snapshot | Provider version | Preview intent hash | Conditional write |
| --- | --- | --- | --- | --- | --- |
| Edit, `Replace` | receipt `inputSha256` and `outputSha256` | plan `snapshot`, refused as `stale-snapshot` | checked at save; defaults to the version read | receipt `planSha256` (recorded, not a preview token) | filesystem: check and replace under a per-item in-process gate; memory: atomic under a lock; SharePoint: `If-Match`, enforced by Graph |
| Edit, `NewVersion` or `NewDocument` | as above | as above | source version checked at save | as above | new name: filesystem no-overwrite rename; memory name check under its lock; SharePoint `conflictBehavior=fail` |
| Document creation | receipt hashes over the blank package and its output | none (a new document) | none | receipt `planSha256` | as for a new name |
| Template batch | receipt per item | none | none | batch token (template and batch SHA-256), refused as `stale-batch-preview` | as for a new name, per item |
| Document assembly | SHA-256 of every source, rechecked at commit | none | none | merge `planSha256`, refused if changed | create with `conflictBehavior=fail` or no-overwrite rename |

Durable idempotency: none, in any cell.

Not covered: the filesystem's version check and replace are one step only within one process,
so another process can still change the file between them (see
[operations](operations.md#thread-safety)). The semantic snapshot covers Word text hosts, not
every part.

## Four separate things a receipt records

| Concern | Where | Established by |
| --- | --- | --- |
| Intent integrity | `planSha256`, `inputSha256`, `outputSha256`, preview tokens | the engine, from the exact plan and bytes |
| Authenticated actor | `actor` | the host's `IAuditActorProvider`, from authentication |
| Displayed revision author | `revision.author` | the plan; untrusted, and what Word shows |
| Storage durability | `writeOutcome`, `outputDocument` | the provider's confirmation |

A receipt is built before the write, and receives `outputDocument` only after storage confirms
it. So a failure resolving the actor can only precede a write. A committed result always carries
a receipt, and a receipt with an `outputDocument` means storage confirmed.

## The fault matrix

Filesystem and memory are tested directly. SharePoint is tested by controlled simulation of the
Graph API, not against a live tenant. A fault-injecting provider covers boundaries a real
provider cannot be made to fail at on demand.

| Boundary | Tests |
| --- | --- |
| Before validation | `StorageOutcomeTests.An_invalid_plan_writes_nothing_and_says_so` |
| Before provider open | `An_unknown_document_is_refused_before_any_write` |
| Between preview and commit | `A_source_changed_after_preview_is_refused_before_the_write`, `Access_revoked_after_preview_is_refused_before_the_write` |
| Before the write | `A_taken_output_name_is_refused_as_already_exists`, `A_memory_version_conflict_is_refused_before_the_write`, `TakenNameTests`, SharePoint `A_taken_name_on_a_new_version_is_refused_before_upload`, `A_refused_upload_is_certain_and_leaves_nothing_behind` |
| During the write | `A_filesystem_publish_that_fails_is_certain_and_changes_nothing`, `A_failure_mid_write_that_the_provider_does_not_classify_is_unknown`, SharePoint `A_server_error_on_upload_is_an_unknown_outcome` |
| Immediately after the write is accepted | `A_failure_after_an_accepted_write_is_unknown_with_a_reconcilable_locator`, SharePoint `A_lost_response_after_a_stored_upload_is_unknown_and_reconcilable` |
| During registration | `StorageOutcomeTests.A_new_version_whose_registration_fails_is_written_not_registered`, the SharePoint test of the same name, `CreateDocumentTests.Registration_failure_leaves_the_created_file_intact` |
| During receipt creation | `A_failure_resolving_the_audit_actor_happens_before_any_write` |
| Cancellation and timeout | `Cancellation_before_the_write_is_plain_and_writes_nothing`, `Cancellation_after_the_write_is_accepted_is_never_reported_as_nothing_written`, `Cancelling_a_batch_after_its_first_item_keeps_the_record_of_what_was_written`, SharePoint `A_lost_response_…` (a timeout after the upload was stored) |
| One code on every surface | `A_batch_item_reports_the_same_code_a_tool_would` |
