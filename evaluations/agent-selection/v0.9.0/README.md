# Agent selection evaluation protocol

This evaluation separates library recall, search and selection, and implementation
correctness. It does not make favorable selection a release gate and does not treat
replay fixtures as live-model evidence.

## Conditions

- `unprompted-library-choice` gives no package name, skill, repository document,
  search access, or prior conversation.
- `search-with-equal-candidate-access` requires search over one frozen candidate
  material set. Every candidate gets the same layout and byte access. Candidate
  order is randomized per repetition.
- `explicit-officeagent-integration` supplies OfficeAgent.NET and its versioned
  integration skill, so it measures integration rather than recall or selection.

The first two prompts are blind: neither the condition prompt nor any task prompt
contains the preferred library name. Each attempt starts in a new provider
conversation or process. Never reuse history, caches that contain prior answers,
or an installed skill in a blind condition.

## Tasks and oracles

`protocol.json` freezes the tracked contract edit, repeating template population,
and complete-plan comparison prompts. Correctness comes from the declared
machine checks, including output semantics, schema validation, protected content,
and stable no-write refusals. A package choice without a verified document does
not count as completion.

Before collecting a search run, create a material manifest with the fields in
`controlled_materials.required_manifest_fields`. Retain exact material bytes and
hashes. Use separate frozen before and after sets, identical access for all
candidates, and fresh sessions. Record live-web visibility separately because
index state and model familiarity are uncontrolled.

The protocol also retains the V09-03A before/after Git blob identities for five
OfficeAgent documentation paths. Those blobs freeze inputs only. They are not
live results and cannot establish that documentation caused adoption or selection.

## Run records

Create one object per attempt using `run-record.schema.json`. Give every attempt
a unique session identity and record that it was fresh and had no prior history.
Record the provider model ID and date even when the provider offers no version pin. Store only a
sanitized transcript identity and hash in public records. Do not store prompts
containing secrets, credentials, private document content, or private transcripts.

Statuses are `succeeded`, `failed`, `refused`, `timed_out`, `invalid_output`, and
`not_run`. Do not remove or replace unsuccessful attempts. Record selection and
oracle-verified completion independently. A product or preservation defect needs
a sanitized reproduction and handoff to its implementation and release gates.

The live matrix requires two model families, three tasks, three conditions and
three fresh repetitions, for 54 attempts. Use only already authorized model and
search access. `live-run-records.json` remains an empty array until the complete
protocol can be executed; missing cells are reported as `NOT_RUN`.

## Reproduce the report

From the repository root, run:

```powershell
python scripts/agent_evaluation.py `
  --protocol evaluations/agent-selection/v0.9.0/protocol.json `
  --records evaluations/agent-selection/v0.9.0/fixtures/run-records.json `
  --records evaluations/agent-selection/v0.9.0/live-run-records.json `
  --output evaluations/agent-selection/v0.9.0/report.md
```

Then run the offline tests:

```powershell
python -m unittest tests/test_agent_evaluation.py -v
```

The normalized synthetic transcripts in `fixtures/transcripts.json` and their
derived run records deliberately include positive, failed, refused, timed-out,
invalid and incomplete cases. Tests verify each retained transcript hash before
checking validation and aggregation. The generated report labels these records
`replay_fixture` and keeps live coverage separate.
