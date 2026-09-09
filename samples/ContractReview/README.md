# ContractReview - a contract review agent (MAF + OfficeAgent.NET)

Reviews a Word contract against a playbook of house positions and produces a
**redlined copy**: tracked changes a lawyer can accept or reject, each with an
anchored comment saying why, plus a markdown report of everything that was
checked.

Uses [Microsoft Agent Framework](https://www.nuget.org/packages/Microsoft.Agents.AI/)
for the judging, Azure OpenAI as the model backend, and OfficeAgent.NET for the
document.

## The design in one line

**The model judges; it never writes.**

The model is good at deciding whether a clause departs from a house position and
at drafting replacement wording. It is unreliable at emitting a valid
multi-operation edit plan. So the work is split:

| Stage | Who does it | What it produces |
| --- | --- | --- |
| Screen | Code (`Screener`) | Candidate anchors, sourced from `find_in_document` |
| Judge | Model (`ReviewAgent`) | Structured findings, one per candidate |
| Build | Code (`PlanBuilder`) | One `DocumentPlan`: comment + tracked change |
| Apply | Code (`ReviewRunner`) | Preview, then a reviewed copy |

The agent is given `inspect_document`, `find_in_document`, and one custom tool,
`report_finding`. It is **not** given `preview_plan` or `apply_plan`. The only
model output that reaches the document is the words inside a tracked insertion
and the text of a comment, and both are length-bounded.

The model never retypes an anchor either: it refers to a candidate by an id the
host minted during screening, so a model slip cannot produce an unresolvable
anchor.

## Run

```bash
export AZURE_OPENAI_ENDPOINT='https://<your-resource>.openai.azure.com'
export AZURE_OPENAI_DEPLOYMENT='<your-chat-deployment>'
export AZURE_OPENAI_API_KEY='<key>'          # optional; omit for DefaultAzureCredential

dotnet run --project samples/ContractReview
```

On first run it generates a small supply agreement with 60-day payment terms,
uncapped liability, automatic renewal, and New York governing law - wording the
bundled `playbook.json` is written to catch.

Optional environment variables:

| Variable | Default |
| --- | --- |
| `REVIEW_CONTRACT` | a generated `contract.docx` |
| `REVIEW_PLAYBOOK` | `playbook.json` |
| `REVIEW_STORAGE` | a per-run temp directory |

## The playbook

Rules are data, not prompt text:

```json
{
  "id": "PAY-01",
  "title": "Payment terms no longer than 30 days",
  "severity": "high",
  "patterns": ["\\b(?:net\\s+)?(\\d{2,3})\\s+days\\b"],
  "regex": true,
  "rationale": "House position is payment within 30 days of invoice.",
  "preferredWording": "30 days",
  "action": "redline"
}
```

`action` decides what the reviewer may do: `redline` allows a tracked change,
`commentOnly` allows only a comment. `rationale` is the test the model applies -
it is the house position, not an instruction about tone.

## Three things the tests pin down

**The plan is bound to a snapshot.** Anchors verify only the text being edited.
A change to an unrelated clause while the model judges leaves every anchor
resolvable, so without a snapshot a stale review applies cleanly to a contract
that moved. The plan carries the etag from an `InspectAsync` taken before
screening; drift becomes `stale-snapshot` and nothing is written.


**Comment before redline, on the same span.** A tracked replacement wraps the
original wording in `w:del`, and struck-through text is no longer part of the
paragraph's visible text. A comment anchored to that wording *after* the
replacement cannot resolve. `PlanBuilder` therefore always emits the comment
first.

The sharp edge: `preview_plan` accepts **both** orders, because it validates each
operation against the pre-apply document. Only the commit fails, with
`expect-mismatch`. `OrderingProbeTests` records this.

**One tracked change per paragraph.** Two playbook patterns can match
overlapping but textually different wording in one clause. The engine keys a
conflict on exact anchor identity, so it would not stop them, and the second
replacement would fail on text the first had already struck through. The builder
allows one redline per paragraph and reports the others as `SupersededOnSpan`.

## Tests

```bash
dotnet test tests/ContractReview.Tests
```

42 tests, no network and no API spend: the model is replaced by a scripted judge,
so the whole pipeline runs offline. They cover playbook validation, screening,
plan assembly, the ordering rule, the tool allowlist, and the failure paths -
a document that changes mid-review, a candidate judged twice, a candidate never
judged, and a redline with no wording proposed.
