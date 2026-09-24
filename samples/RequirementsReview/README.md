# Requirements review: TypeSafe Jev, Microsoft Agent Framework, OfficeAgent.NET

A customer receives a supplier proposal in Word. Its data-privacy section says that logs
containing names and mobile numbers will be shared with all project members and kept after
handover without a deletion date. The default sample checks that passage against three
registered requirements and proposes Word comments in a new copy. A reviewer approves the
feedback before anything is written. No agent decides the final route:

| Component | What it does here | What it cannot do |
| --- | --- | --- |
| **TypeSafe Jev** (REST, `POST /v1/systemone`) | Answers three fixed questions per requirement about one section of the document | Choose questions, see other sections, route, approve, or write |
| **Microsoft Agent Framework review agent** (`ChatClientAgent`) | Calls the Jev-backed `evaluate_requirement` function tool once for each eligible requirement | Change the requirement text, choose evidence, set thresholds, route, approve, or write |
| **Microsoft Agent Framework drafting agent** (`ChatClientAgent`) | Drafts a comment for a gap the host has confirmed | Read other sections, change the route, approve, or write |
| **OfficeAgent.NET** | Reads the document, anchors evidence to paragraphs, previews and commits one `DocumentPlan` | Decide anything about requirements |
| **Application code** | Owns requirements, evidence selection, deterministic checks, thresholds, routing, the plan, permissions, the audit record, and the final write | |
| **A reviewer** | Approves or rejects the previewed plan | |

The proposal and all offline answers are fictional. The review agent receives requirement IDs;
the tool looks up the registered text and anchored passage, calls Jev, and records the typed
decision. The host checks that every eligible requirement was called exactly once. A missing,
duplicate, or failed call goes to human review.

| Requirement | Checked by | Result | In Word |
| --- | --- | --- | --- |
| REQ-PRIV-ACCESS Named access roles | Jev tool | Gap | Comment on sharing with all project members |
| REQ-PRIV-STORAGE Restricted location | Jev tool | Gap | Comment on the shared project folder |
| REQ-PRIV-RETENTION Deletion date | Jev tool | Gap | Comment on indefinite retention |

In the offline demo, the review agent's tool calls, Jev-wire answers, and drafted comments
are scripted. They show the mechanism; they are not live Jev or live-model results. Use
`--legacy` to run the earlier installation-method-statement scenario.

## Prerequisites

- .NET 8 SDK. The repository's `global.json` pins SDK 8.0.100 with `latestFeature`. On a
  machine with only SDK 9 or 10, run the commands from outside the repository with the
  project path, for example from its parent folder.
- Nothing else for the offline run: no network, no keys.
- In your own application, reference the packages instead of this repository's projects:
  `OfficeAgent.Core` and `OfficeAgent.Word` 1.0.0, `Microsoft.Agents.AI` 1.22.0,
  `DocumentFormat.OpenXml` 3.5.1, and for a live drafting model
  `Microsoft.Extensions.AI.OpenAI` 10.10.0. The sample was built and run against these
  published packages.

## Run it offline

```bash
dotnet run --project samples/RequirementsReview -- --out ./review-out
dotnet run --project samples/RequirementsReview -- --out ./review-out-approved --approve-as ops-lead@example.com
```

The first command stops after the preview and writes no document. The second approves the
preview as `ops-lead@example.com`, the one reviewer allowed by `review-policy.json`, and
writes the reviewed copy. `--approve-as` stands in for a signed-in reviewer pressing
"approve"; a real application takes the reviewer id from its authentication. A second run
into the same folder reviews the same document again and writes a time-stamped copy, report
and audit record; nothing from the earlier run is overwritten.

| Output | Contents |
| --- | --- |
| `supplier-proposal.docx` | The source document. Never modified. |
| `supplier-proposal.reviewed.docx` | The reviewed copy, only after approval. |
| `report.md` | Evidence, answers, route and Word action per requirement. |
| `audit.json` | The machine-readable record described below. |

## Code map

| File | Role | Reuse |
| --- | --- | --- |
| `JevRequirementEvaluator.cs` | `HttpClient` adapter for `POST /v1/systemone`: builds the request, retries 429/529, validates every answer | Reuse |
| `Evaluation.cs` | `IRequirementEvaluator`, `RequirementDecision`, and `QuestionSet`, the fixed questions | Reuse; adapt the questions to your domain |
| `AgentRequirementReviewer.cs` | `ChatClientAgent` with the Jev-backed `evaluate_requirement` tool; detects missing and duplicate calls | Reuse |
| `ReviewPolicy.cs`, `review-policy.json` | Thresholds, allowed reviewers, and `Router` (rules R1-R10) | Reuse; calibrate the thresholds |
| `Requirements.cs`, `privacy-requirements.json` | The versioned privacy requirement registry | Reuse; write your own requirements |
| `Evidence.cs` | Reads the document once, hashes it, inspects it with OfficeAgent.NET, cuts one section per requirement, flags pending tracked changes | Reuse |
| `HardRules.cs` | Deterministic pattern checks | Reuse |
| `ProposalDrafter.cs` | `ChatClientAgent` with one tool, `submit_proposal`, whose arguments the host validates | Reuse |
| `PlanBuilder.cs` | Builds the `DocumentPlan`: comments first, tracked changes only where a requirement allows them | Reuse |
| `AuthorizedCommitter.cs` | Checks the approval, previews again, writes a new copy | Reuse; plug in your authentication |
| `RequirementsWorkflow.cs` | Runs the steps in order | Reuse |
| `AuditRecord.cs`, `Report.cs` | `audit.json` and `report.md` | Adapt |
| `Program.cs` | Console host and flags | Demo |
| `PrivacyProposalFixture.cs` | Generates the default supplier proposal | Demo only |
| `Scripted.cs`, `ScriptedReviewChatClient.cs` | Offline Jev transport and scripted agents | Demo and tests only |
| `verify-in-word.ps1` | Opens a result in desktop Word and lists what Word sees | Optional check |

## Review your own document

```bash
dotnet run --project samples/RequirementsReview -- --out ./my-review --document ./path/to/submission.docx
```

Edit `privacy-requirements.json` first. Each requirement names the heading of the section
that holds its evidence (`section`), and headings are matched exactly using Word's built-in
Heading 1-9 styles. Missing or ambiguous sections go to human review without a Jev call.
Set `"proposal": "trackedChange"` on a requirement to allow replacement wording as a tracked
change; the default proposal uses comments only because a supplier owns its text. The
`--legacy` scenario uses `requirements.json` and includes a hard rule that stops its run
before any external call if its mandatory programme information is absent.

## Live paths

Both are opt-in and read credentials only from environment variables. Nothing in this folder
contains a key, and the audit record and logs never include one.

```bash
TYPESAFE_API_KEY=... dotnet run --project samples/RequirementsReview -- --out ./live --live-jev
AZURE_OPENAI_ENDPOINT=https://<resource>.openai.azure.com AZURE_OPENAI_DEPLOYMENT=<deployment> AZURE_OPENAI_API_KEY=... \
  dotnet run --project samples/RequirementsReview -- --out ./live --live-jev --live-agents
```

- Jev is in early access. The adapter follows the published contract, API version 0.2.0
  ([openapi.json](https://api.typesafe.ai/openapi.json), [API reference](https://docs.typesafe.ai/api),
  [docs index for agents](https://docs.typesafe.ai/llms.txt)), checked in September 2026.
- `review-policy.json` requests `jev-latest`, an alias that moves with each release. Pin a
  version such as `jev-1.13.0` once you have calibrated the thresholds. The audit records the
  model the API reports and the `x-typesafe-request-id` response header.
- 429 and 529 responses are retried up to twice with exponential backoff (1 s, then 2 s)
  inside the 30-second call timeout, as TypeSafe recommends. A `Retry-After` value is
  honoured, never shortened; if it asks for more than 8 seconds, the requirement goes to a
  person instead. So does anything still throttled after the retries.
- `--live-agents` uses the OpenAI client against the Azure OpenAI v1 endpoint
  (`<endpoint>/openai/v1/`) for both the review and drafting agents. Without this flag,
  their tool calls and drafts remain scripted even if `--live-jev` is set. The older
  `--live-drafter` flag remains an alias.

## Data boundary

| Destination | Receives | Never receives |
| --- | --- | --- |
| Jev | Requirement id, version and text; the paragraphs of that requirement's section, at most 4,000 characters | Document name, paragraph ids, other sections, comments, authors, policy, thresholds |
| Review agent | Eligible requirement IDs and tool result labels | Document passages, Jev probabilities, thresholds, approval, Word write tools |
| Drafting model | The same, only for requirements judged not met | Probabilities, thresholds, policy, other requirements |
| Nowhere | Deterministic checks, routing, plan building, preview, commit | |

A section over the budget goes to a person rather than being truncated. Evidence is not
redacted; add redaction before the adapter if sections can contain personal or confidential
data. Check the providers' terms against your data classification. In September 2026, TypeSafe
stated that Jev is not trained on customer data and that the service is hosted in the United
States. API use falls under its Master Customer Agreement and Data Processing Addendum, which
uses the EU standard contractual clauses. Zero data retention is available to enterprise
customers ([docs.typesafe.ai/legal](https://docs.typesafe.ai/legal)). English is the language
Jev handles best.

## Fails closed

| Condition | Result | Rule |
| --- | --- | --- |
| A hard rule marked `onFail: stop` fails | Run stops before anything is sent | H2 |
| Section missing, duplicated, or empty | Human review; nothing sent | C1-C3 |
| Review state of the section unreadable | Human review; nothing sent | C4 |
| Section contains a pending tracked change | Human review; nothing sent | C5 |
| Section over the budget | Human review; nothing sent | C6 |
| Timeout, outage, auth error, rejected request, malformed or incomplete answer, throttling after retries | Human review | R1-R2 |
| Review agent skips or repeats a requirement tool call, or calls an unknown ID | Human review; unknown ID is never sent to Jev | R1 |
| `evidence_status` not `sufficient`, or its confidence below 0.80 | Human review | R3-R4 |
| `requirement_result` is `not_determined`, or its confidence below 0.80 | Human review | R5-R6 |
| Jev says `satisfied` but a registered keyword check finds nothing | Human review | R7 |
| Document changed after it was read | Proposal stale; nothing written | preview |
| Document changed after approval | Commit refused; nothing written | revalidation |
| No approval, unknown reviewer, approval for another plan, or plan changed after approval | Commit refused | committer |

Confidence is TypeSafe's value, derived from the probabilities; for a choice it is
`(options × top probability − 1) / (options − 1)`. The 0.80 thresholds are placeholders to
calibrate on documents your reviewers have already assessed.

## Audit record

`audit.json` holds, per run:

- the document id, SHA-256 and snapshot;
- the requirement set and policy versions and hashes, and the thresholds;
- the exact questions and allowed answers;
- per requirement: evidence ids, paragraph ids and excerpts, deterministic results, Jev
  answers with probabilities, model, request id and attempts, route, rule, reason, and the
  drafter's model and proposal;
- the plan and its hash, the preview receipt, the reviewer's decision with the hashes it
  covers, and the commit receipt.

A test recomputes every route from the record alone.

## Tests

```bash
dotnet test tests/RequirementsReview.Tests
```

125 offline tests cover the privacy tool path, missing/duplicate/unknown tool calls and fail-closed routing, the legacy three outcomes, preservation of its existing table, comment and
tracked change, the tracked-change option, the Jev request and response contract, every HTTP
failure class, retries and `Retry-After`, timeouts, malformed answers, thresholds, disagreement, stale documents
before and after approval, unauthorized and tampered commits, existing reviewed copies, the
drafter's single tool and argument checks, and the audit record.

The tests read the saved package (`word/document.xml`, `word/comments.xml`). That proves the
markup is present, not how Word shows it. For that:

```powershell
./samples/RequirementsReview/verify-in-word.ps1 -Path ./review-out-approved/supplier-proposal.reviewed.docx
```

It opens the copy read-only in desktop Word, lists what Word sees, and closes without saving.
The default privacy copy was checked in Word 16.0.20326: it showed three comments on the
storage, access, and retention phrases, with no revisions. The older `--legacy` scenario was
also checked in Word, preserving its supplier insertion, comments, and three-row table.

## Limitations

- Each requirement is judged on its own section. A requirement that depends on several
  sections needs its own explicit check.
- Supplier text can argue for its own classification, which TypeSafe lists as a known
  weakness of Jev 1.13. The questions state that a claim of compliance is not evidence; test
  with adversarial examples before relying on a pass.
- Headings are matched exactly; a renamed heading goes to a person.
- Deterministic checks are simple regular expressions with a 250 ms timeout.
- The proposal is held in memory between preview and approval. A real application stores the
  plan and its hashes and reloads them for the commit.
- If the document changes between the final check and the save, the copy is still written, and
  the committer reports `CommittedAgainstChangedInput`. The source is never modified.
- A typed answer can still be wrong, a probability is not proof, and a pass is not a
  compliance sign-off.
