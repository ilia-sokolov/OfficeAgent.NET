#!/usr/bin/env python3
"""
Summarise raw benchmark results, and compare a candidate against a baseline.

The harness that produces the raw results does not compute statistics, decide whether a run
is acceptable, or write the report. That separation is deliberate: a component that both
produces numbers and judges them can hide a bad run inside its own arithmetic, and none of
its arithmetic can be tested without running a benchmark. Everything here is a pure function
of a JSON file, so every rule below is tested against fixed inputs.

Four rules are load bearing, and each of them came from a measurement rather than a
preference.

**Every iteration counts.** Failed iterations are reported, never dropped, and a scenario
with any failure is reported as failed regardless of how fast its successful iterations
were. Reporting the speed of the cases that happened to work is the classic way to publish a
number that is true and useless.

**Variance is measured between runs, not between repetitions.** The first version of this
policy derived its threshold from the spread of repetitions inside one baseline run. Two
consecutive runs of the same commit then reported two scenarios as regressed by 24%.
Repetitions inside one process share a warm cache, a warm JIT, one thermal state and one set
of background neighbours, so their spread describes the machine at one moment and badly
understates how far the same benchmark moves when it is run again.

**Wall-clock time is reported; allocation is gated.** Measured across five runs of one
commit on an ordinary developer workstation, the per-scenario medians moved by 9% to 21%
between runs, and a five-run baseline at three standard deviations still produced a 46% false
regression. Allocated bytes moved by 0.016% to 0.052% across the same runs, which is three
orders of magnitude steadier, because allocation is a property of the code rather than of the
machine's mood. So a timing change is reported and never blocks, and an allocation change is
what actually gates. A timing gate on this class of hardware would fail for no reason, and a
gate that fails for no reason is one people learn to ignore.

**Protocol cost is gated, and call counts exactly.** For the installed MCP path, the number of
tool calls and the bytes sent were identical across every measured iteration, while bytes
received moved by about 1.2% because generated content differs slightly run to run. So a change
in the number of round trips an agent pays for is compared exactly and any increase blocks,
while byte totals get a threshold. This is the metric that catches an adapter change doubling
the calls needed to do the same work, which no timing number on a noisy machine would show.
"""

from __future__ import annotations

import argparse
import json
import statistics
import sys
from pathlib import Path
from typing import Any

# 1.0 is the direct .NET harness. 1.1 adds optional protocol counters from the installed MCP
# path. Both are accepted so published 1.0 baselines keep loading; widening 1.0 in place would
# have made the version meaningless.
SUPPORTED_SCHEMAS = {"1.0", "1.1"}
SCHEMA_VERSION = "1.0"

# Wall-clock floor. Timings on a shared machine are not stable below roughly this level.
MINIMUM_TIME_THRESHOLD = 0.10

# Allocation floor. Measured between-run spread is under 0.06%, so this is a wide margin that
# still catches any change worth calling a regression.
MINIMUM_ALLOCATION_THRESHOLD = 0.05

# No scenario gets an unlimited allowance because its baseline was noisy. A scenario that
# unstable should be fixed or dropped, not gated loosely.
MAXIMUM_THRESHOLD = 1.00

# How many multiples of the baseline's between-run spread count as a real change.
VARIANCE_MULTIPLIER = 3.0

# A baseline of one run cannot show how much a benchmark moves between runs, which is the only
# variance a regression has to be distinguished from. One run is accepted, because refusing to
# report anything is worse, but it gets the floor and is marked provisional.
MINIMUM_BASELINE_RUNS = 2

# Protocol byte totals carry a little content-driven variance, measured at about 1.2% between
# iterations, so they get a threshold rather than exact equality. Tool call counts do not vary
# at all and are compared exactly.
MINIMUM_PROTOCOL_THRESHOLD = 0.05

# An allocation or protocol regression blocks; a timing regression does not. See the docstring.
BLOCKING = {"failed", "allocation-regressed", "protocol-regressed", "baseline-unusable",
            "inputs-changed", "missing"}


class ReportError(Exception):
    """A malformed or unusable results file."""


def load(path: Path) -> dict[str, Any]:
    document = json.loads(path.read_text(encoding="utf-8"))
    version = document.get("SchemaVersion")
    if version not in SUPPORTED_SCHEMAS:
        raise ReportError(
            f"unsupported results schema {version!r}, expected one of {sorted(SUPPORTED_SCHEMAS)}")
    if not document.get("Scenarios"):
        raise ReportError("results contain no scenarios")
    return document


def summarize_scenario(scenario: dict[str, Any]) -> dict[str, Any]:
    """Reduce one scenario's iterations to the figures the report prints."""
    iterations = scenario.get("Iterations") or []
    if not iterations:
        raise ReportError(f"scenario {scenario.get('Name')!r} recorded no iterations")

    succeeded = [i for i in iterations if i.get("Succeeded")]
    failures = [i for i in iterations if not i.get("Succeeded")]

    summary: dict[str, Any] = {
        "name": scenario.get("Name", ""),
        "description": scenario.get("Description", ""),
        "inputs": scenario.get("Inputs", []),
        "samples": len(iterations),
        "failures": len(failures),
        "failure_messages": sorted({str(i.get("Failure")) for i in failures if i.get("Failure")}),
        "warmup": scenario.get("WarmupIterations", 0),
    }

    if not succeeded:
        # Every iteration failed. There is no timing to report, and deriving one from failed
        # attempts would describe how long it takes to not work.
        summary.update(
            {k: None for k in
             ("min_ms", "median_ms", "p95_ms", "max_ms", "within_run_spread",
              "median_allocated_bytes", "median_output_bytes", "median_input_bytes",
              "work_units", "input_bytes_per_second", "work_units_per_second",
              "tool_calls", "median_bytes_sent", "median_bytes_received")}
        )
        summary["concurrency"] = int(scenario.get("ConcurrencyLevel", 1) or 1)
        summary["preserved_parts"] = list(scenario.get("PreservedParts", []) or [])
        return summary

    times = sorted(float(i["ElapsedMilliseconds"]) for i in succeeded)
    summary.update(
        {
            "min_ms": times[0],
            "median_ms": statistics.median(times),
            "p95_ms": percentile(times, 0.95),
            "max_ms": times[-1],
            # Reported for context only. It is not what any threshold is derived from.
            "within_run_spread": relative_spread(times),
            "median_allocated_bytes": statistics.median(int(i["AllocatedBytes"]) for i in succeeded),
            "median_output_bytes": statistics.median(int(i["OutputBytes"]) for i in succeeded),
            "median_input_bytes": statistics.median(int(i.get("InputBytes", 0)) for i in succeeded),
            "work_units": statistics.median(int(i["WorkUnits"]) for i in succeeded),
        }
    )

    # Throughput, from the median rather than the mean, so one slow outlier cannot flatter or
    # spoil it. Two denominators because neither alone describes every scenario: a comparison
    # produces no output bytes, and a batch's work is its item count rather than its size.
    seconds = summary["median_ms"] / 1000.0 if summary["median_ms"] else 0.0
    summary["input_bytes_per_second"] = (
        summary["median_input_bytes"] / seconds if seconds > 0 and summary["median_input_bytes"] else None)
    summary["work_units_per_second"] = (
        summary["work_units"] / seconds if seconds > 0 and summary["work_units"] else None)
    summary["concurrency"] = int(scenario.get("ConcurrencyLevel", 1) or 1)
    summary["preserved_parts"] = list(scenario.get("PreservedParts", []) or [])

    # Optional, and present only for the installed MCP path.
    protocols = [i["Protocol"] for i in succeeded if isinstance(i.get("Protocol"), dict)]
    if protocols:
        summary["tool_calls"] = statistics.median(int(p.get("ToolCalls", 0)) for p in protocols)
        summary["median_bytes_sent"] = statistics.median(
            int(p.get("BytesSent", 0)) for p in protocols)
        summary["median_bytes_received"] = statistics.median(
            int(p.get("BytesReceived", 0)) for p in protocols)
    else:
        summary["tool_calls"] = None
        summary["median_bytes_sent"] = None
        summary["median_bytes_received"] = None
    return summary


def percentile(sorted_values: list[float], fraction: float) -> float:
    """Nearest-rank percentile, which needs no interpolation story to explain."""
    if not sorted_values:
        raise ReportError("cannot take a percentile of no values")
    rank = max(1, min(len(sorted_values), int(round(fraction * len(sorted_values) + 0.5))))
    return sorted_values[rank - 1]


def relative_spread(values: list[float]) -> float:
    """
    Standard deviation as a fraction of the median.

    One value has no spread, reported as 0.0 and handled by the threshold floor rather than by
    pretending a single measurement is precise.
    """
    values = [v for v in values if v is not None]
    if len(values) < 2:
        return 0.0
    median = statistics.median(values)
    if median <= 0:
        return 0.0
    return statistics.stdev(values) / median


def aggregate_baseline(runs: list[dict[str, Any]]) -> dict[str, dict[str, Any]]:
    """
    Reduce several baseline runs to one entry per scenario, keyed by name.

    What matters is how far each scenario moved between runs, because that is the noise a real
    regression has to exceed. A scenario missing from any run, or failing in any run, is marked
    unusable: a baseline stitched together from whichever runs happened to contain a scenario
    would quietly compare against a different population.
    """
    entries: dict[str, dict[str, Any]] = {}
    for index, run in enumerate(runs):
        for raw in run["Scenarios"]:
            summary = summarize_scenario(raw)
            entry = entries.setdefault(
                summary["name"],
                {
                    "name": summary["name"],
                    "medians": [],
                    "allocations": [],
                    "tool_calls": [],
                    "bytes": [],
                    "inputs": summary["inputs"],
                    "runs_present": 0,
                    "unusable": False,
                    "reason": None,
                },
            )
            entry["runs_present"] += 1
            if summary["inputs"] != entry["inputs"]:
                entry["unusable"] = True
                entry["reason"] = f"run {index + 1} used different inputs"
            if summary["failures"]:
                entry["unusable"] = True
                entry["reason"] = f"run {index + 1} recorded {summary['failures']} failed iteration(s)"
            elif summary["median_ms"] is None:
                entry["unusable"] = True
                entry["reason"] = f"run {index + 1} produced no usable timing"
            else:
                entry["medians"].append(float(summary["median_ms"]))
                entry["allocations"].append(float(summary["median_allocated_bytes"]))
                if summary.get("tool_calls") is not None:
                    entry["tool_calls"].append(int(summary["tool_calls"]))
                    entry["bytes"].append(
                        float(summary["median_bytes_sent"]) + float(summary["median_bytes_received"]))

    for entry in entries.values():
        if entry["runs_present"] != len(runs):
            entry["unusable"] = True
            entry["reason"] = f"present in {entry['runs_present']} of {len(runs)} baseline runs"
        entry["runs"] = len(runs)
        entry["median_ms"] = statistics.median(entry["medians"]) if entry["medians"] else None
        entry["median_allocated_bytes"] = (
            statistics.median(entry["allocations"]) if entry["allocations"] else None)
        entry["time_spread"] = relative_spread(entry["medians"])
        entry["allocation_spread"] = relative_spread(entry["allocations"])
        entry["tool_calls"] = statistics.median(entry["tool_calls"]) if entry["tool_calls"] else None
        entry["median_protocol_bytes"] = statistics.median(entry["bytes"]) if entry["bytes"] else None
        entry["protocol_spread"] = relative_spread(entry["bytes"])
        entry["provisional"] = len(entry["medians"]) < MINIMUM_BASELINE_RUNS
    return entries


def threshold_for(entry: dict[str, Any], *, metric: str) -> float:
    """
    The change a scenario must exceed before it is called a regression in `metric`.

    Derived from how far that metric moved between baseline runs. A scenario that moves by s
    from one run to the next cannot distinguish a change smaller than a few multiples of s
    from its own noise.
    """
    floor = {
        "time": MINIMUM_TIME_THRESHOLD,
        "allocation": MINIMUM_ALLOCATION_THRESHOLD,
        "protocol": MINIMUM_PROTOCOL_THRESHOLD,
    }[metric]
    if entry.get("provisional"):
        return floor
    spread = entry.get(
        {"time": "time_spread", "allocation": "allocation_spread",
         "protocol": "protocol_spread"}[metric])
    if spread is None:
        return floor
    return min(MAXIMUM_THRESHOLD, max(floor, VARIANCE_MULTIPLIER * float(spread)))


def change_between(current: float | None, previous: float | None) -> float | None:
    if current is None or not previous:
        return None
    return (current - previous) / previous


def compare(candidate: dict[str, Any], baselines: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Compare like for like, and say plainly when a comparison is not like for like."""
    aggregated = aggregate_baseline(baselines)
    rows: list[dict[str, Any]] = []

    for scenario in candidate["Scenarios"]:
        current = summarize_scenario(scenario)
        previous = aggregated.get(current["name"])
        row: dict[str, Any] = {
            "name": current["name"],
            "candidate": current,
            "baseline": previous,
            "time_change": None,
            "time_threshold": None,
            "allocation_change": None,
            "allocation_threshold": None,
            "protocol_change": None,
            "protocol_threshold": None,
            "tool_call_change": None,
            "provisional": bool(previous and previous.get("provisional")),
            "status": "ok",
        }

        if current["failures"]:
            row["status"] = "failed"
        elif previous is None:
            # Not a pass and not a regression: there is nothing to compare against, and
            # calling it either would be a guess.
            row["status"] = "new"
        elif previous["unusable"]:
            row["status"] = "baseline-unusable"
            row["reason"] = previous["reason"]
        elif current["inputs"] != previous["inputs"]:
            row["status"] = "inputs-changed"
        elif current["median_ms"] is None or not previous["median_ms"]:
            row["status"] = "baseline-unusable"
        else:
            row["time_threshold"] = threshold_for(previous, metric="time")
            row["time_change"] = change_between(current["median_ms"], previous["median_ms"])
            row["allocation_threshold"] = threshold_for(previous, metric="allocation")
            row["allocation_change"] = change_between(
                current["median_allocated_bytes"], previous["median_allocated_bytes"])

            if previous.get("tool_calls") is not None and current.get("tool_calls") is not None:
                row["tool_call_change"] = int(current["tool_calls"]) - int(previous["tool_calls"])
                row["protocol_threshold"] = threshold_for(previous, metric="protocol")
                row["protocol_change"] = change_between(
                    (current["median_bytes_sent"] or 0) + (current["median_bytes_received"] or 0),
                    previous.get("median_protocol_bytes"))

            allocation = row["allocation_change"]
            timing = row["time_change"]
            protocol = row["protocol_change"]
            calls = row["tool_call_change"]
            if calls:
                # Exact: the number of round trips an agent pays for does not vary between
                # runs, so any change in it is a change in the product, not in the weather.
                row["status"] = "protocol-regressed" if calls > 0 else "fewer-calls"
            elif protocol is not None and protocol > row["protocol_threshold"]:
                row["status"] = "protocol-regressed"
            elif allocation is not None and allocation > row["allocation_threshold"]:
                row["status"] = "allocation-regressed"
            elif timing is not None and timing > row["time_threshold"]:
                # Reported, not blocking. Wall-clock noise on ordinary hardware is far larger
                # than any threshold worth setting; see the module docstring.
                row["status"] = "slower"
            elif timing is not None and timing < -row["time_threshold"]:
                row["status"] = "faster"

        rows.append(row)

    present = {s["Name"] for s in candidate["Scenarios"]}
    for name in sorted(set(aggregated) - present):
        rows.append(
            {
                "name": name,
                "candidate": None,
                "baseline": aggregated[name],
                "time_change": None,
                "time_threshold": None,
                "allocation_change": None,
                "allocation_threshold": None,
                "protocol_change": None,
                "protocol_threshold": None,
                "tool_call_change": None,
                "provisional": False,
                "status": "missing",
            }
        )

    return rows


def render(candidate: dict[str, Any], rows: list[dict[str, Any]] | None,
           baseline_runs: int = 0) -> str:
    env = candidate.get("Environment", {})
    summaries = [summarize_scenario(s) for s in candidate["Scenarios"]]

    lines: list[str] = [
        "# Benchmark results",
        "",
        "Generated by `scripts/benchmark_report.py` from raw results. Every figure is a median",
        "over the recorded iterations of one scenario, with failures counted, not dropped.",
        "",
        "## Run",
        "",
        "| Field | Value |",
        "| --- | --- |",
        f"| Commit | `{env.get('Commit', '')}` |",
        f"| Working tree clean | {yes_no(env.get('CommitIsClean'))} |",
        f"| Runtime | {env.get('Runtime', '')} |",
        f"| Operating system | {env.get('OperatingSystem', '')} |",
        f"| Architecture | {env.get('Architecture', '')} |",
        f"| Logical processors | {env.get('ProcessorCount', '')} |",
        f"| Server garbage collection | {yes_no(env.get('ServerGarbageCollection'))} |",
        f"| Build configuration | {env.get('Configuration', '')} |",
        f"| Repetitions per scenario | {env.get('RepetitionsPerScenario', '')} |",
        f"| Warmup iterations, discarded | {env.get('WarmupIterations', '')} |",
        f"| Started | {env.get('StartedUtc', '')} |",
        "",
        f"Measurement method: {env.get('MeasurementMethod', '')}",
        "",
        "## Scenarios",
        "",
        "| Scenario | Conc | Samples | Failures | Median ms | p95 ms | Median alloc | Input/s | Work/s |",
        "| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |",
    ]

    for summary in summaries:
        lines.append(
            f"| `{summary['name']}` | {summary.get('concurrency', 1)} | {summary['samples']} | "
            f"{summary['failures']} | {number(summary['median_ms'])} | {number(summary['p95_ms'])} | "
            f"{size(summary['median_allocated_bytes'])} | "
            f"{rate(summary.get('input_bytes_per_second'))} | "
            f"{number(summary.get('work_units_per_second'))} |"
        )

    preserved = [s for s in summaries if s.get("preserved_parts")]
    if preserved:
        lines += [
            "",
            "### Parts asserted unchanged",
            "",
            "Correctness and preservation are different claims. These parts were compared byte",
            "for byte between the staged input and the produced output on every iteration.",
            "",
            "| Scenario | Parts |",
            "| --- | --- |",
        ]
        for summary in preserved:
            parts = ", ".join(f"`{part}`" for part in summary["preserved_parts"])
            lines.append(f"| `{summary['name']}` | {parts} |")

    protocol = [s for s in summaries if s.get("tool_calls") is not None]
    if protocol:
        lines += [
            "",
            "### Protocol cost",
            "",
            "What an agent pays to drive this over stdio. Tool calls and bytes sent do not vary",
            "between runs at all; bytes received carries about 1% content-driven variance.",
            "",
            "| Scenario | Tool calls | Bytes sent | Bytes received |",
            "| --- | ---: | ---: | ---: |",
        ]
        for summary in protocol:
            lines.append(
                f"| `{summary['name']}` | {int(summary['tool_calls'])} | "
                f"{size(summary['median_bytes_sent'])} | {size(summary['median_bytes_received'])} |"
            )

    lines += [
        "",
        f"Total iterations: {sum(s['samples'] for s in summaries)}. "
        f"Failed iterations: {sum(s['failures'] for s in summaries)}.",
    ]

    failing = [s for s in summaries if s["failures"]]
    if failing:
        lines += ["", "### Failures", ""]
        for summary in failing:
            for message in summary["failure_messages"]:
                lines.append(f"- `{summary['name']}`: {message}")

    if rows is not None:
        lines += [
            "",
            "## Against the baseline",
            "",
            f"Baseline runs: {baseline_runs}.",
            "",
            "Thresholds are three times how far each metric moved **between baseline runs**, not",
            "between repetitions inside one run, clamped to a floor and a ceiling. Repetitions",
            "inside one process share a warm cache, a warm JIT and one thermal state, so their",
            "spread understates how far the same benchmark moves when run again.",
            "",
            "**Allocation gates; wall-clock time is reported only.** Measured over five runs of a",
            "single commit, scenario medians moved 9% to 21% between runs while allocated bytes",
            "moved under 0.06%. A timing gate on this class of hardware fails for no reason, and",
            "a gate that fails for no reason is one people learn to ignore. A `slower` row is a",
            "real signal worth investigating and does not block.",
            "",
            "Protocol call counts are compared exactly, because they do not vary between runs.",
            "",
            "| Scenario | Status | Calls | Protocol bytes | Alloc change | Time change | Baseline |",
            "| --- | --- | ---: | ---: | ---: | ---: | --- |",
        ]
        for row in rows:
            entry = row.get("baseline") or {}
            if row.get("provisional"):
                basis = "provisional, one run"
            elif entry:
                basis = (f"{entry.get('runs', 0)} runs, alloc {percent(entry.get('allocation_spread'))}, "
                         f"time {percent(entry.get('time_spread'))}")
            else:
                basis = "none"
            calls = row.get("tool_call_change")
            lines.append(
                f"| `{row['name']}` | {row['status']} | "
                f"{'n/a' if calls is None else f'{calls:+d}'} | "
                f"{signed_percent(row['protocol_change'])} | "
                f"{signed_percent(row['allocation_change'])} | "
                f"{signed_percent(row['time_change'])} | {basis} |"
            )

        blocking = [row for row in rows if row["status"] in BLOCKING]
        lines += ["", f"Blocking rows: {len(blocking)}."]
        if blocking:
            lines += [
                "",
                "A blocking row stops the release unless a scoped decision is recorded against it.",
                "`ok`, `new`, `faster`, `fewer-calls` and `slower` do not block.",
            ]

    lines.append("")
    return "\n".join(lines)


def yes_no(value: Any) -> str:
    return "unknown" if value is None else ("yes" if value else "no")


def number(value: float | None) -> str:
    return "n/a" if value is None else f"{value:.2f}"


def percent(value: float | None) -> str:
    return "n/a" if value is None else f"{value:.2%}"


def signed_percent(value: float | None) -> str:
    return "n/a" if value is None else f"{value:+.2%}"


def rate(value: float | None) -> str:
    """Bytes per second, rendered like a size so the units are obvious."""
    return "n/a" if value is None else f"{size(value)}/s"


def size(value: float | None) -> str:
    if value is None:
        return "n/a"
    scaled = float(value)
    for unit in ["B", "KiB", "MiB", "GiB"]:
        if scaled < 1024 or unit == "GiB":
            return f"{scaled:.1f} {unit}"
        scaled /= 1024
    return f"{scaled:.1f} GiB"


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Summarise and compare benchmark results.")
    parser.add_argument("--results", required=True, type=Path, help="Raw candidate results JSON.")
    parser.add_argument(
        "--baseline",
        type=Path,
        nargs="+",
        help="Raw baseline results JSON files. Give several runs so thresholds can be derived "
        "from between-run variance; one run is accepted but gets the floor and is marked "
        "provisional.",
    )
    parser.add_argument("--output", type=Path, help="Markdown report path; stdout when omitted.")
    parser.add_argument(
        "--fail-on-blocking",
        action="store_true",
        help="Exit non-zero when any row blocks, for use as a release gate.",
    )
    args = parser.parse_args(argv)

    try:
        candidate = load(args.results)
        baselines = [load(path) for path in args.baseline] if args.baseline else []
        rows = compare(candidate, baselines) if baselines else None
        report = render(candidate, rows, len(baselines))
    except (ReportError, json.JSONDecodeError, OSError) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2

    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(report, encoding="utf-8", newline="\n")
        print(f"report={args.output}")
    else:
        print(report)

    failed = sum(s["failures"] for s in (summarize_scenario(x) for x in candidate["Scenarios"]))
    blocking = [row for row in rows if row["status"] in BLOCKING] if rows else []
    print(f"failed-iterations={failed} blocking-rows={len(blocking)}", file=sys.stderr)
    if failed or (args.fail_on_blocking and blocking):
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
