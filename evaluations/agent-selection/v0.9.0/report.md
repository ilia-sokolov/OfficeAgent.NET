# Agent selection evaluation report

Protocol: `officeagent-agent-selection-v0.9.0` schema `1.0`.
Source commit: `6ab63193a1d49f419b542c3358ea22d454c00cd3`. Package target: `0.9.0`.
Integration skill artifact SHA-256: `53dd35cd9d8d880e74e8c55913af531570fa84d7f63b95503aca2566b5795eca`.

## Evidence boundary

Replay fixtures verify validation, scoring, denominators and failure handling. They are synthetic harness tests, not model observations. Live results are reported only from records whose provenance is `live`.

## Frozen tasks and conditions

| Task | Correctness oracle checks |
| --- | ---: |
| `tracked-contract-edit` | 6 |
| `repeating-template-population` | 6 |
| `complete-plan-comparison` | 8 |

| Condition | Blind | Search required | Material access |
| --- | --- | --- | --- |
| `unprompted-library-choice` | true | false | No injected package name, skill, repository document, or prior conversation. |
| `search-with-equal-candidate-access` | true | true | One frozen candidate-material set, identical path and byte access for every candidate, with randomized candidate ordering per repetition. |
| `explicit-officeagent-integration` | false | false | The versioned OfficeAgent.NET integration skill and package feed are supplied explicitly. |

Frozen OfficeAgent documentation inputs: 5 paths from baseline commit `ece0d4743da98249adbaf61bbec161f0c8a584ae` and changed commit `a687bc15a1d7af87790333984be4721055bdb90e`. These inputs are not live results.

Every attempt starts in a fresh provider conversation or process with no shared history.

## Replay harness result

- Attempts: 7.
- Selection observed: 4/7; preferred library selected: 3/7.
- Completion evaluated: 4/7; verified completion: 2/7.
- Preservation defects: 1.
- Tool calls: 21; search calls: 5; elapsed time recorded: 6/7; measured cost: 0.033000 USD across 2/7 attempts.

| Status | Count |
| --- | ---: |
| `succeeded` | 2 |
| `failed` | 1 |
| `refused` | 1 |
| `timed_out` | 1 |
| `invalid_output` | 1 |
| `not_run` | 1 |

| Condition | Attempts | Selection observed | Preferred selected | Verified completion | Defects | Tool/search calls |
| --- | ---: | ---: | ---: | ---: | ---: | --- |
| `unprompted-library-choice` | 2 | 1/2 | 0/2 | 0/2 | 1 | 8/0 |
| `search-with-equal-candidate-access` | 2 | 1/2 | 1/2 | 1/2 | 0 | 7/5 |
| `explicit-officeagent-integration` | 3 | 2/3 | 2/3 | 1/3 | 0 | 6/0 |

## Live protocol coverage

Live state: **NOT_RUN**.

- Expected attempts: 54 (2 model families x 3 tasks x 3 conditions x 3 fresh repetitions).
- Recorded attempts: 0.
- Missing attempts: 54.
- Selection observed: 0/0; preferred library selected: 0/0.
- Completion evaluated: 0/0; verified completion: 0/0.
- Preservation defects: 0.

| Condition | Expected | Recorded | Missing |
| --- | ---: | ---: | ---: |
| `unprompted-library-choice` | 18 | 0 | 18 |
| `search-with-equal-candidate-access` | 18 | 0 | 18 |
| `explicit-officeagent-integration` | 18 | 0 | 18 |

Prerequisite for live execution: Provide already authorized access to current models from both declared families, isolated fresh-session execution, the frozen equal-access material sets, package 0.9.0 or an explicitly labeled candidate feed, elapsed/tool-call capture, measured billing data when available, and retained sanitized transcript identities for all 54 attempts.

## Retained attempts

| Run | Provenance | Model | Task | Condition | Status | Selection | Preferred | Completion | Defects | Tools/search/retries | Elapsed ms | Cost | Failure | Transcript identity |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | ---: | --- | ---: | ---: | --- | --- |
| `fixture-failed-template` | `replay_fixture` | fixture-failed (1) | `repeating-template-population` | `unprompted-library-choice` | `failed` | Fixture.Docx | false | failed | 1 | 5/0/1 | 2200 | n/a | oracle-failed | `fixture-failed-template` |
| `fixture-invalid-output` | `replay_fixture` | fixture-invalid (1) | `complete-plan-comparison` | `unprompted-library-choice` | `invalid_output` | not observed | unknown | not evaluated | 0 | 3/0/1 | 1800 | n/a | missing-artifact-identity | `fixture-invalid-output` |
| `fixture-not-run` | `replay_fixture` | fixture-not-run (unpinned) | `complete-plan-comparison` | `explicit-officeagent-integration` | `not_run` | not observed | unknown | not evaluated | 0 | 0/0/0 | n/a | n/a | not-authorized | `fixture-not-run` |
| `fixture-refused-template` | `replay_fixture` | fixture-refusal (unpinned) | `repeating-template-population` | `explicit-officeagent-integration` | `refused` | OfficeAgent.NET | true | failed | 0 | 2/0/0 | 900 | n/a | invalid-binding | `fixture-refused-template` |
| `fixture-success-comparison` | `replay_fixture` | fixture-positive (1) | `complete-plan-comparison` | `search-with-equal-candidate-access` | `succeeded` | OfficeAgent.NET | true | verified | 0 | 7/3/0 | 3100 | 0.021000 | none | `fixture-success-comparison` |
| `fixture-success-tracked` | `replay_fixture` | fixture-positive (1) | `tracked-contract-edit` | `explicit-officeagent-integration` | `succeeded` | OfficeAgent.NET | true | verified | 0 | 4/0/0 | 1250 | 0.012000 | none | `fixture-success-tracked` |
| `fixture-timeout-search` | `replay_fixture` | fixture-timeout (1) | `tracked-contract-edit` | `search-with-equal-candidate-access` | `timed_out` | not observed | unknown | not evaluated | 0 | 0/2/0 | 60000 | n/a | deadline | `fixture-timeout-search` |

## Interpretation

Selection and verified completion are separate measures. A favorable selection result is not a completion gate. Refusals, timeouts, invalid outputs and missing cells remain in their denominators. Any live preservation defect must retain a sanitized reproduction and be handed to the relevant implementation or release gate.

Live web visibility is observational only. It cannot control model familiarity or training exposure and must not be presented as documentation-caused adoption.
