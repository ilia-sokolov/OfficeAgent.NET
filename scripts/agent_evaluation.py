#!/usr/bin/env python3
"""Validate and aggregate controlled agent selection evaluation records."""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
from collections import Counter
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[1]
STATUS_VALUES = (
    "succeeded",
    "failed",
    "refused",
    "timed_out",
    "invalid_output",
    "not_run",
)
PROVENANCE_VALUES = ("replay_fixture", "live")


class EvaluationError(ValueError):
    """Raised when a protocol or run record is invalid."""


def load_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def _require(mapping: dict[str, Any], key: str, expected: type, context: str) -> Any:
    value = mapping.get(key)
    if not isinstance(value, expected):
        raise EvaluationError(f"{context}.{key} must be {expected.__name__}")
    return value


def _is_sha256(value: Any) -> bool:
    return isinstance(value, str) and len(value) == 64 and all(
        character in "0123456789abcdef" for character in value
    )


def _is_git_object_id(value: Any) -> bool:
    return isinstance(value, str) and len(value) == 40 and all(
        character in "0123456789abcdef" for character in value
    )


def render_prompt(protocol: dict[str, Any], task_id: str, condition_id: str) -> str:
    tasks = {item["id"]: item for item in protocol["tasks"]}
    conditions = {item["id"]: item for item in protocol["conditions"]}
    try:
        task = tasks[task_id]
        condition = conditions[condition_id]
    except KeyError as error:
        raise EvaluationError(f"unknown task or condition: {error.args[0]}") from error
    return f"{condition['prompt'].strip()}\n\n{task['prompt'].strip()}\n"


def prompt_sha256(protocol: dict[str, Any], task_id: str, condition_id: str) -> str:
    return hashlib.sha256(render_prompt(protocol, task_id, condition_id).encode("utf-8")).hexdigest()


def validate_protocol(protocol: dict[str, Any]) -> None:
    _require(protocol, "protocol_id", str, "protocol")
    _require(protocol, "schema_version", str, "protocol")
    preferred = _require(protocol, "preferred_library", str, "protocol").casefold()
    if not _is_git_object_id(protocol.get("source_commit")):
        raise EvaluationError("protocol.source_commit must be a lowercase Git object id")
    _require(protocol, "package_version", str, "protocol")
    if not _is_sha256(protocol.get("skill_artifact_sha256")):
        raise EvaluationError("protocol.skill_artifact_sha256 must be a lowercase SHA-256")

    tasks = _require(protocol, "tasks", list, "protocol")
    conditions = _require(protocol, "conditions", list, "protocol")
    if len(tasks) < 3:
        raise EvaluationError("protocol must define at least three tasks")
    if len(conditions) != 3:
        raise EvaluationError("protocol must define exactly three conditions")

    task_ids = [item.get("id") for item in tasks]
    condition_ids = [item.get("id") for item in conditions]
    if len(set(task_ids)) != len(task_ids) or not all(isinstance(value, str) for value in task_ids):
        raise EvaluationError("task ids must be unique strings")
    if len(set(condition_ids)) != len(condition_ids) or not all(
        isinstance(value, str) for value in condition_ids
    ):
        raise EvaluationError("condition ids must be unique strings")

    for task in tasks:
        _require(task, "prompt", str, f"task {task.get('id')}")
        oracle = _require(task, "oracle", dict, f"task {task.get('id')}")
        checks = _require(oracle, "required_checks", list, f"task {task.get('id')}.oracle")
        if not checks or not all(isinstance(value, str) and value for value in checks):
            raise EvaluationError(f"task {task.get('id')} requires non-empty oracle checks")

    for condition in conditions:
        context = f"condition {condition.get('id')}"
        prompt = _require(condition, "prompt", str, context)
        blind = _require(condition, "blind", bool, context)
        _require(condition, "search_required", bool, context)
        _require(condition, "material_access", str, context)
        if blind:
            for task_id in task_ids:
                rendered = render_prompt(protocol, task_id, condition["id"])
                if preferred in rendered.casefold():
                    raise EvaluationError(
                        f"blind prompt {condition['id']}/{task_id} leaks preferred library"
                    )
        elif preferred not in prompt.casefold():
            raise EvaluationError("explicit condition must name the preferred library")

    live = _require(protocol, "live_protocol", dict, "protocol")
    families = _require(live, "model_families", list, "protocol.live_protocol")
    repetitions = _require(live, "fresh_repetitions_per_cell", int, "protocol.live_protocol")
    if len(families) < 2 or len(set(families)) != len(families):
        raise EvaluationError("live protocol requires at least two distinct model families")
    if repetitions < 3:
        raise EvaluationError("live protocol requires at least three fresh repetitions per cell")
    if live.get("session_policy") != "fresh_session_per_attempt_no_shared_history":
        raise EvaluationError("live protocol must prohibit shared conversation history")
    _require(live, "not_run_prerequisite", str, "protocol.live_protocol")

    controlled = _require(protocol, "controlled_materials", dict, "protocol")
    _require(controlled, "freeze_policy", str, "protocol.controlled_materials")
    _require(controlled, "comparison_policy", str, "protocol.controlled_materials")
    snapshots = _require(
        controlled,
        "officeagent_documentation_snapshots",
        dict,
        "protocol.controlled_materials",
    )
    for key in ("historical_baseline_commit", "changed_commit"):
        if not _is_git_object_id(snapshots.get(key)):
            raise EvaluationError(f"protocol.controlled_materials snapshots.{key} is invalid")
    snapshot_files = _require(
        snapshots, "files", list, "protocol.controlled_materials snapshots"
    )
    if not snapshot_files:
        raise EvaluationError("protocol must retain documentation snapshot identities")
    for item in snapshot_files:
        if not isinstance(item, dict) or not isinstance(item.get("path"), str):
            raise EvaluationError("documentation snapshot entry is invalid")
        if item.get("before_blob") is not None and not _is_git_object_id(item.get("before_blob")):
            raise EvaluationError(f"documentation snapshot before blob is invalid: {item.get('path')}")
        if not _is_git_object_id(item.get("after_blob")):
            raise EvaluationError(f"documentation snapshot after blob is invalid: {item.get('path')}")


def validate_record(record: dict[str, Any], protocol: dict[str, Any]) -> None:
    context = f"record {record.get('run_id', '<missing>')}"
    _require(record, "run_id", str, context)
    provenance = _require(record, "provenance", str, context)
    if provenance not in PROVENANCE_VALUES:
        raise EvaluationError(f"{context}.provenance is invalid")
    status = _require(record, "status", str, context)
    if status not in STATUS_VALUES:
        raise EvaluationError(f"{context}.status is invalid")
    task_id = _require(record, "task_id", str, context)
    condition_id = _require(record, "condition_id", str, context)
    if task_id not in {item["id"] for item in protocol["tasks"]}:
        raise EvaluationError(f"{context}.task_id is unknown")
    if condition_id not in {item["id"] for item in protocol["conditions"]}:
        raise EvaluationError(f"{context}.condition_id is unknown")
    repetition = _require(record, "repetition", int, context)
    if repetition < 1:
        raise EvaluationError(f"{context}.repetition must be positive")
    if record.get("protocol_id") != protocol["protocol_id"]:
        raise EvaluationError(f"{context}.protocol_id does not match")
    expected_prompt_hash = prompt_sha256(protocol, task_id, condition_id)
    if record.get("prompt_sha256") != expected_prompt_hash:
        raise EvaluationError(f"{context}.prompt_sha256 does not match the frozen prompt")

    model = _require(record, "model", dict, context)
    _require(model, "family", str, f"{context}.model")
    _require(model, "provider_model_id", str, f"{context}.model")
    _require(model, "execution_date", str, f"{context}.model")
    _require(model, "version_pin_available", bool, f"{context}.model")
    if model.get("model_version") is not None and not isinstance(model.get("model_version"), str):
        raise EvaluationError(f"{context}.model.model_version must be string or null")
    if provenance == "live":
        if model["family"] not in protocol["live_protocol"]["model_families"]:
            raise EvaluationError(f"{context}.model.family is outside the live protocol")
        if repetition > protocol["live_protocol"]["fresh_repetitions_per_cell"]:
            raise EvaluationError(f"{context}.repetition is outside the live protocol")

    session = _require(record, "session", dict, context)
    _require(session, "identity", str, f"{context}.session")
    if _require(session, "fresh", bool, f"{context}.session") is not True:
        raise EvaluationError(f"{context}.session must be fresh")
    if _require(session, "prior_history", bool, f"{context}.session") is not False:
        raise EvaluationError(f"{context}.session must not contain prior history")

    transcript = _require(record, "transcript", dict, context)
    _require(transcript, "identity", str, f"{context}.transcript")
    if not _is_sha256(transcript.get("sha256")):
        raise EvaluationError(f"{context}.transcript.sha256 must be a lowercase SHA-256")
    _require(transcript, "retained", bool, f"{context}.transcript")

    access = _require(record, "access", dict, context)
    tools = _require(access, "tools", list, f"{context}.access")
    if not all(isinstance(value, str) for value in tools):
        raise EvaluationError(f"{context}.access.tools must contain strings")
    _require(access, "search_enabled", bool, f"{context}.access")
    if access.get("material_set") is not None and not isinstance(access.get("material_set"), str):
        raise EvaluationError(f"{context}.access.material_set must be string or null")
    condition = next(item for item in protocol["conditions"] if item["id"] == condition_id)
    if access["search_enabled"] != condition["search_required"]:
        raise EvaluationError(f"{context}.access.search_enabled does not match the condition")
    if condition["search_required"] and access.get("material_set") is None:
        raise EvaluationError(f"{context}.access.material_set is required for search")

    if record.get("package_version") != protocol["package_version"]:
        raise EvaluationError(f"{context}.package_version does not match")
    if record.get("source_commit") != protocol["source_commit"]:
        raise EvaluationError(f"{context}.source_commit does not match")
    if record.get("skill_artifact_sha256") != protocol["skill_artifact_sha256"]:
        raise EvaluationError(f"{context}.skill_artifact_sha256 does not match")

    attempt_number = _require(record, "attempt_number", int, context)
    if attempt_number < 1:
        raise EvaluationError(f"{context}.attempt_number must be positive")
    elapsed = record.get("elapsed_ms")
    if elapsed is not None and (not isinstance(elapsed, int) or elapsed < 0):
        raise EvaluationError(f"{context}.elapsed_ms must be a non-negative integer or null")

    cost = _require(record, "cost", dict, context)
    measured = _require(cost, "measured", bool, f"{context}.cost")
    amount = cost.get("amount")
    if amount is not None and (not isinstance(amount, (int, float)) or amount < 0):
        raise EvaluationError(f"{context}.cost.amount must be non-negative or null")
    if measured and amount is None:
        raise EvaluationError(f"{context}.cost.amount is required when measured")
    _require(cost, "currency", str, f"{context}.cost")
    if cost["currency"] != "USD":
        raise EvaluationError(f"{context}.cost.currency must be USD")

    selection = _require(record, "selection", dict, context)
    candidate = selection.get("candidate")
    preferred = selection.get("is_preferred")
    if candidate is not None and not isinstance(candidate, str):
        raise EvaluationError(f"{context}.selection.candidate must be string or null")
    if preferred is not None and not isinstance(preferred, bool):
        raise EvaluationError(f"{context}.selection.is_preferred must be boolean or null")
    if (candidate is None) != (preferred is None):
        raise EvaluationError(f"{context}.selection candidate and preference must be jointly known")

    completion = _require(record, "completion", dict, context)
    evaluated = _require(completion, "oracle_evaluated", bool, f"{context}.completion")
    verified = completion.get("verified")
    if verified is not None and not isinstance(verified, bool):
        raise EvaluationError(f"{context}.completion.verified must be boolean or null")
    if evaluated != (verified is not None):
        raise EvaluationError(f"{context}.completion verified state is inconsistent")
    oracle_results = _require(completion, "oracle_results", list, f"{context}.completion")
    for result in oracle_results:
        if not isinstance(result, dict) or not isinstance(result.get("check"), str) or not isinstance(
            result.get("passed"), bool
        ):
            raise EvaluationError(f"{context}.completion.oracle_results is invalid")
    artifact_hash = completion.get("output_artifact_sha256")
    if artifact_hash is not None and not _is_sha256(artifact_hash):
        raise EvaluationError(f"{context}.completion.output_artifact_sha256 is invalid")

    if status == "succeeded":
        if verified is not True or artifact_hash is None:
            raise EvaluationError(f"{context} succeeded without a verified output artifact")
        results_by_check = {item["check"]: item["passed"] for item in oracle_results}
        required_checks = next(
            item["oracle"]["required_checks"]
            for item in protocol["tasks"]
            if item["id"] == task_id
        )
        if any(results_by_check.get(check) is not True for check in required_checks):
            raise EvaluationError(f"{context} succeeded without all required oracle checks")
    if status == "not_run" and elapsed is not None:
        raise EvaluationError(f"{context} not_run record cannot have elapsed time")

    defects = _require(record, "preservation_defects", list, context)
    if not all(isinstance(value, str) for value in defects):
        raise EvaluationError(f"{context}.preservation_defects must contain strings")
    effort = _require(record, "integration_effort", dict, context)
    for key in ("tool_calls", "search_calls", "retries"):
        value = _require(effort, key, int, f"{context}.integration_effort")
        if value < 0:
            raise EvaluationError(f"{context}.integration_effort.{key} must be non-negative")
    failure = _require(record, "failure", dict, context)
    for key in ("category", "code", "detail"):
        if failure.get(key) is not None and not isinstance(failure.get(key), str):
            raise EvaluationError(f"{context}.failure.{key} must be string or null")
    if status == "succeeded" and any(failure.get(key) is not None for key in failure):
        raise EvaluationError(f"{context} succeeded record cannot contain a failure")
    if status != "succeeded" and failure.get("category") is None:
        raise EvaluationError(f"{context} non-success record requires a failure category")
    notes = _require(record, "notes", list, context)
    if not all(isinstance(value, str) for value in notes):
        raise EvaluationError(f"{context}.notes must contain strings")


def load_records(paths: list[Path], protocol: dict[str, Any]) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    seen_ids: set[str] = set()
    seen_sessions: set[str] = set()
    for path in paths:
        loaded = load_json(path)
        if not isinstance(loaded, list):
            raise EvaluationError(f"{path} must contain a JSON array")
        for record in loaded:
            if not isinstance(record, dict):
                raise EvaluationError(f"{path} contains a non-object record")
            validate_record(record, protocol)
            if record["run_id"] in seen_ids:
                raise EvaluationError(f"duplicate run id: {record['run_id']}")
            session_id = record["session"]["identity"]
            if session_id in seen_sessions:
                raise EvaluationError(f"reused session identity: {session_id}")
            seen_ids.add(record["run_id"])
            seen_sessions.add(session_id)
            records.append(record)
    return records


def _cell_key(record: dict[str, Any]) -> tuple[str, str, str, int]:
    return (
        record["model"]["family"],
        record["task_id"],
        record["condition_id"],
        record["repetition"],
    )


def aggregate(protocol: dict[str, Any], records: list[dict[str, Any]]) -> dict[str, Any]:
    live_protocol = protocol["live_protocol"]
    expected_cells = {
        (family, task["id"], condition["id"], repetition)
        for family in live_protocol["model_families"]
        for task in protocol["tasks"]
        for condition in protocol["conditions"]
        for repetition in range(1, live_protocol["fresh_repetitions_per_cell"] + 1)
    }
    live_records = [record for record in records if record["provenance"] == "live"]
    replay_records = [record for record in records if record["provenance"] == "replay_fixture"]
    observed_cells: set[tuple[str, str, str, int]] = set()
    for record in live_records:
        cell = _cell_key(record)
        if cell not in expected_cells:
            raise EvaluationError(f"live run {record['run_id']} is outside the declared matrix")
        if cell in observed_cells:
            raise EvaluationError(f"duplicate live matrix cell: {cell}")
        observed_cells.add(cell)

    def summarize(group: list[dict[str, Any]]) -> dict[str, Any]:
        statuses = Counter(record["status"] for record in group)
        selected = [record for record in group if record["selection"]["candidate"] is not None]
        completed = [record for record in group if record["completion"]["verified"] is not None]
        return {
            "attempts": len(group),
            "statuses": {status: statuses.get(status, 0) for status in STATUS_VALUES},
            "selection_observed": len(selected),
            "preferred_selected": sum(record["selection"]["is_preferred"] is True for record in selected),
            "completion_evaluated": len(completed),
            "completion_verified": sum(record["completion"]["verified"] is True for record in completed),
            "preservation_defects": sum(len(record["preservation_defects"]) for record in group),
            "tool_calls": sum(record["integration_effort"]["tool_calls"] for record in group),
            "search_calls": sum(record["integration_effort"]["search_calls"] for record in group),
            "measured_cost": round(
                sum(record["cost"]["amount"] or 0 for record in group if record["cost"]["measured"]),
                6,
            ),
            "cost_measured_attempts": sum(record["cost"]["measured"] for record in group),
            "elapsed_measured_attempts": sum(record["elapsed_ms"] is not None for record in group),
        }

    by_condition: dict[str, dict[str, int]] = {}
    for condition in protocol["conditions"]:
        condition_id = condition["id"]
        expected = sum(cell[2] == condition_id for cell in expected_cells)
        observed = sum(cell[2] == condition_id for cell in observed_cells)
        by_condition[condition_id] = {
            "expected": expected,
            "recorded": observed,
            "missing": expected - observed,
        }

    return {
        "replay": summarize(replay_records),
        "live": summarize(live_records),
        "replay_by_condition": {
            condition["id"]: summarize(
                [record for record in replay_records if record["condition_id"] == condition["id"]]
            )
            for condition in protocol["conditions"]
        },
        "live_by_condition": {
            condition["id"]: summarize(
                [record for record in live_records if record["condition_id"] == condition["id"]]
            )
            for condition in protocol["conditions"]
        },
        "live_coverage": {
            "expected": len(expected_cells),
            "recorded": len(observed_cells),
            "missing": len(expected_cells - observed_cells),
            "by_condition": by_condition,
        },
    }


def _ratio(numerator: int, denominator: int) -> str:
    return f"{numerator}/{denominator}" if denominator else "0/0"


def generate_report(
    protocol: dict[str, Any], records: list[dict[str, Any]], summary: dict[str, Any]
) -> str:
    replay = summary["replay"]
    live = summary["live"]
    coverage = summary["live_coverage"]
    live_state = "RECORDED" if coverage["recorded"] else "NOT_RUN"
    lines = [
        "# Agent selection evaluation report",
        "",
        f"Protocol: `{protocol['protocol_id']}` schema `{protocol['schema_version']}`.",
        f"Source commit: `{protocol['source_commit']}`. Package target: `{protocol['package_version']}`.",
        f"Integration skill artifact SHA-256: `{protocol['skill_artifact_sha256']}`.",
        "",
        "## Evidence boundary",
        "",
        "Replay fixtures verify validation, scoring, denominators and failure handling. They are synthetic harness tests, not model observations. Live results are reported only from records whose provenance is `live`.",
        "",
        "## Frozen tasks and conditions",
        "",
        "| Task | Correctness oracle checks |",
        "| --- | ---: |",
    ]
    for task in protocol["tasks"]:
        lines.append(f"| `{task['id']}` | {len(task['oracle']['required_checks'])} |")
    lines.extend(["", "| Condition | Blind | Search required | Material access |", "| --- | --- | --- | --- |"])
    for condition in protocol["conditions"]:
        lines.append(
            f"| `{condition['id']}` | {str(condition['blind']).lower()} | "
            f"{str(condition['search_required']).lower()} | {condition['material_access']} |"
        )
    snapshots = protocol["controlled_materials"]["officeagent_documentation_snapshots"]
    lines.extend(
        [
            "",
            f"Frozen OfficeAgent documentation inputs: {len(snapshots['files'])} paths from baseline commit `{snapshots['historical_baseline_commit']}` and changed commit `{snapshots['changed_commit']}`. These inputs are not live results.",
        ]
    )

    lines.extend(
        [
            "",
            "Every attempt starts in a fresh provider conversation or process with no shared history.",
            "",
            "## Replay harness result",
            "",
            f"- Attempts: {replay['attempts']}.",
            f"- Selection observed: {_ratio(replay['selection_observed'], replay['attempts'])}; preferred library selected: {_ratio(replay['preferred_selected'], replay['attempts'])}.",
            f"- Completion evaluated: {_ratio(replay['completion_evaluated'], replay['attempts'])}; verified completion: {_ratio(replay['completion_verified'], replay['attempts'])}.",
            f"- Preservation defects: {replay['preservation_defects']}.",
            f"- Tool calls: {replay['tool_calls']}; search calls: {replay['search_calls']}; elapsed time recorded: {_ratio(replay['elapsed_measured_attempts'], replay['attempts'])}; measured cost: {replay['measured_cost']:.6f} USD across {_ratio(replay['cost_measured_attempts'], replay['attempts'])} attempts.",
            "",
            "| Status | Count |",
            "| --- | ---: |",
        ]
    )
    for status in STATUS_VALUES:
        lines.append(f"| `{status}` | {replay['statuses'][status]} |")

    lines.extend(
        [
            "",
            "| Condition | Attempts | Selection observed | Preferred selected | Verified completion | Defects | Tool/search calls |",
            "| --- | ---: | ---: | ---: | ---: | ---: | --- |",
        ]
    )
    for condition in protocol["conditions"]:
        item = summary["replay_by_condition"][condition["id"]]
        lines.append(
            f"| `{condition['id']}` | {item['attempts']} | "
            f"{_ratio(item['selection_observed'], item['attempts'])} | "
            f"{_ratio(item['preferred_selected'], item['attempts'])} | "
            f"{_ratio(item['completion_verified'], item['attempts'])} | "
            f"{item['preservation_defects']} | {item['tool_calls']}/{item['search_calls']} |"
        )

    lines.extend(
        [
            "",
            "## Live protocol coverage",
            "",
            f"Live state: **{live_state}**.",
            "",
            f"- Expected attempts: {coverage['expected']} ({len(protocol['live_protocol']['model_families'])} model families x {len(protocol['tasks'])} tasks x {len(protocol['conditions'])} conditions x {protocol['live_protocol']['fresh_repetitions_per_cell']} fresh repetitions).",
            f"- Recorded attempts: {coverage['recorded']}.",
            f"- Missing attempts: {coverage['missing']}.",
            f"- Selection observed: {_ratio(live['selection_observed'], live['attempts'])}; preferred library selected: {_ratio(live['preferred_selected'], live['attempts'])}.",
            f"- Completion evaluated: {_ratio(live['completion_evaluated'], live['attempts'])}; verified completion: {_ratio(live['completion_verified'], live['attempts'])}.",
            f"- Preservation defects: {live['preservation_defects']}.",
            "",
            "| Condition | Expected | Recorded | Missing |",
            "| --- | ---: | ---: | ---: |",
        ]
    )
    for condition in protocol["conditions"]:
        item = coverage["by_condition"][condition["id"]]
        lines.append(
            f"| `{condition['id']}` | {item['expected']} | {item['recorded']} | {item['missing']} |"
        )
    lines.extend(
        [
            "",
            f"Prerequisite for live execution: {protocol['live_protocol']['not_run_prerequisite']}",
            "",
            "## Retained attempts",
            "",
            "| Run | Provenance | Model | Task | Condition | Status | Selection | Preferred | Completion | Defects | Tools/search/retries | Elapsed ms | Cost | Failure | Transcript identity |",
            "| --- | --- | --- | --- | --- | --- | --- | --- | --- | ---: | --- | ---: | ---: | --- | --- |",
        ]
    )
    for record in sorted(records, key=lambda item: item["run_id"]):
        selection = record["selection"]["candidate"] or "not observed"
        preferred = record["selection"]["is_preferred"]
        preferred_text = "unknown" if preferred is None else str(preferred).lower()
        verified = record["completion"]["verified"]
        completion = "not evaluated" if verified is None else ("verified" if verified else "failed")
        elapsed = "n/a" if record["elapsed_ms"] is None else str(record["elapsed_ms"])
        cost = "n/a" if not record["cost"]["measured"] else f"{record['cost']['amount']:.6f}"
        model_version = record["model"]["model_version"] or "unpinned"
        model = f"{record['model']['provider_model_id']} ({model_version})"
        effort = record["integration_effort"]
        failure = record["failure"]["code"] or "none"
        lines.append(
            f"| `{record['run_id']}` | `{record['provenance']}` | {model} | `{record['task_id']}` | "
            f"`{record['condition_id']}` | `{record['status']}` | {selection} | {preferred_text} | "
            f"{completion} | {len(record['preservation_defects'])} | {effort['tool_calls']}/{effort['search_calls']}/{effort['retries']} | {elapsed} | {cost} | {failure} | "
            f"`{record['transcript']['identity']}` |"
        )
    if not records:
        lines.append("| none | n/a | n/a | n/a | n/a | n/a | n/a | n/a | n/a | 0 | 0/0/0 | n/a | n/a | n/a | n/a |")

    lines.extend(
        [
            "",
            "## Interpretation",
            "",
            "Selection and verified completion are separate measures. A favorable selection result is not a completion gate. Refusals, timeouts, invalid outputs and missing cells remain in their denominators. Any live preservation defect must retain a sanitized reproduction and be handed to the relevant implementation or release gate.",
            "",
            "Live web visibility is observational only. It cannot control model familiarity or training exposure and must not be presented as documentation-caused adoption.",
            "",
        ]
    )
    return "\n".join(lines)


def run(protocol_path: Path, record_paths: list[Path], output_path: Path) -> None:
    protocol = load_json(protocol_path)
    if not isinstance(protocol, dict):
        raise EvaluationError("protocol must be a JSON object")
    validate_protocol(protocol)
    records = load_records(record_paths, protocol)
    summary = aggregate(protocol, records)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(generate_report(protocol, records, summary), encoding="utf-8", newline="\n")


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--protocol", type=Path, required=True)
    parser.add_argument("--records", type=Path, action="append", required=True)
    parser.add_argument("--output", type=Path, required=True)
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    try:
        run(args.protocol, args.records, args.output)
    except (EvaluationError, json.JSONDecodeError, OSError) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 1
    print(f"Agent evaluation report written to {args.output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
