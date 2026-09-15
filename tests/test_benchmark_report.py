"""
Tests for the benchmark reporter's arithmetic and its threshold policy.

These run against fixed inputs and never execute a benchmark, which is the point of keeping
the statistics out of the harness: the rules that decide whether a release is blocked are
testable without a stopwatch, and nothing here can be affected by the machine it runs on.
"""

from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

import benchmark_report as report  # noqa: E402


def scenario(name: str, times: list[float], *, failures: int = 0,
             inputs: list[str] | None = None, allocated: int = 1000,
             protocol: dict | None = None) -> dict:
    iterations = [
        {
            "Index": index,
            "ElapsedMilliseconds": value,
            "AllocatedBytes": allocated,
            "OutputBytes": 2000,
            "WorkUnits": 3,
            "Succeeded": True,
            "Failure": None,
            **({"Protocol": protocol} if protocol else {}),
        }
        for index, value in enumerate(times)
    ]
    for index in range(failures):
        iterations.append(
            {
                "Index": len(iterations),
                "ElapsedMilliseconds": 0.0,
                "AllocatedBytes": 0,
                "OutputBytes": 0,
                "WorkUnits": 0,
                "Succeeded": False,
                "Failure": f"assertion {index} failed",
            }
        )
    return {
        "Name": name,
        "Description": "",
        "Inputs": inputs if inputs is not None else ["a.docx"],
        "WarmupIterations": 2,
        "Iterations": iterations,
        "FailedIterations": failures,
        "ProcessPeakWorkingSetBytes": 0,
    }


def run(*scenarios: dict) -> dict:
    return {
        "SchemaVersion": report.SCHEMA_VERSION,
        "Environment": {"Commit": "abc", "CommitIsClean": True},
        "Scenarios": list(scenarios),
    }


class SummaryTests(unittest.TestCase):
    def test_failed_iterations_are_counted_not_dropped(self) -> None:
        """A fast scenario that failed twice is a failing scenario, not a fast one."""
        summary = report.summarize_scenario(scenario("s", [10.0, 10.0, 10.0], failures=2))
        self.assertEqual(5, summary["samples"])
        self.assertEqual(2, summary["failures"])
        self.assertEqual(10.0, summary["median_ms"])
        self.assertEqual(["assertion 0 failed", "assertion 1 failed"], summary["failure_messages"])

    def test_a_scenario_that_never_succeeded_reports_no_timing(self) -> None:
        """Timing derived from failed attempts would describe how long it takes to not work."""
        summary = report.summarize_scenario(scenario("s", [], failures=3))
        self.assertEqual(3, summary["failures"])
        self.assertIsNone(summary["median_ms"])
        self.assertIsNone(summary["within_run_spread"])

    def test_a_scenario_with_no_iterations_is_an_error(self) -> None:
        with self.assertRaises(report.ReportError):
            report.summarize_scenario(scenario("s", []))

    def test_percentile_uses_nearest_rank(self) -> None:
        values = [1.0, 2.0, 3.0, 4.0, 5.0]
        self.assertEqual(5.0, report.percentile(values, 0.95))
        self.assertEqual(3.0, report.percentile(values, 0.5))

    def test_a_single_sample_has_no_spread(self) -> None:
        self.assertEqual(0.0, report.relative_spread([42.0]))


class ThresholdTests(unittest.TestCase):
    """
    The threshold is a property of the baseline as a set of runs, not of one summary.

    The two tests this class replaced asserted that a single run's internal spread widened
    its own threshold. That is the policy that let two consecutive runs of the same commit
    report a 24% regression, so the tests were rewritten to the corrected contract rather
    than adjusted to keep passing.
    """

    @staticmethod
    def entry(*run_medians: float) -> dict:
        return report.aggregate_baseline([run(scenario("s", [value] * 3)) for value in run_medians])["s"]

    def test_a_steady_baseline_gets_the_floor(self) -> None:
        self.assertEqual(
            report.MINIMUM_TIME_THRESHOLD,
            report.threshold_for(self.entry(100.0, 100.0, 100.0), metric="time"))

    def test_a_baseline_that_moves_between_runs_widens_its_threshold(self) -> None:
        self.assertGreater(
            report.threshold_for(self.entry(100.0, 140.0, 180.0), metric="time"), report.MINIMUM_TIME_THRESHOLD)

    def test_no_scenario_gets_an_unlimited_allowance(self) -> None:
        """A scenario this unstable should be fixed or dropped, not gated loosely."""
        self.assertEqual(
            report.MAXIMUM_THRESHOLD, report.threshold_for(self.entry(1.0, 100.0, 1000.0), metric="time"))


class ComparisonTests(unittest.TestCase):
    def test_a_timing_regression_is_reported_but_does_not_block(self) -> None:
        """
        Wall clock is far too noisy on ordinary hardware to gate on. Five runs of one commit
        moved 9% to 21% between runs here, so a timing gate fails for no reason.
        """
        baseline = run(scenario("s", [100.0, 100.0, 100.0]))
        candidate = run(scenario("s", [200.0, 200.0, 200.0]))
        row = report.compare(candidate, [baseline, baseline])[0]
        self.assertEqual("slower", row["status"])
        self.assertNotIn("slower", report.BLOCKING)
        self.assertAlmostEqual(1.0, row["time_change"])

    def test_an_allocation_regression_blocks(self) -> None:
        """Allocation is a property of the code, so it can carry the gate."""
        baseline = run(scenario("s", [100.0] * 3, allocated=1000))
        candidate = run(scenario("s", [100.0] * 3, allocated=2000))
        row = report.compare(candidate, [baseline, baseline])[0]
        self.assertEqual("allocation-regressed", row["status"])
        self.assertIn("allocation-regressed", report.BLOCKING)
        self.assertAlmostEqual(1.0, row["allocation_change"])

    def test_allocation_takes_precedence_over_timing(self) -> None:
        """A run that got slower and allocates more is blocked on the signal that is real."""
        baseline = run(scenario("s", [100.0] * 3, allocated=1000))
        candidate = run(scenario("s", [200.0] * 3, allocated=2000))
        self.assertEqual("allocation-regressed", report.compare(candidate, [baseline, baseline])[0]["status"])

    def test_a_change_inside_the_threshold_does_not_block(self) -> None:
        baseline = run(scenario("s", [100.0, 100.0, 100.0]))
        candidate = run(scenario("s", [105.0, 105.0, 105.0]))
        row = report.compare(candidate, [baseline, baseline])[0]
        self.assertEqual("ok", row["status"])

    def test_an_improvement_is_reported_and_does_not_block(self) -> None:
        baseline = run(scenario("s", [100.0, 100.0, 100.0]))
        candidate = run(scenario("s", [50.0, 50.0, 50.0]))
        row = report.compare(candidate, [baseline, baseline])[0]
        self.assertEqual("faster", row["status"])
        self.assertNotIn("faster", report.BLOCKING)

    def test_a_failing_candidate_blocks_however_fast_it_was(self) -> None:
        baseline = run(scenario("s", [100.0, 100.0, 100.0]))
        candidate = run(scenario("s", [1.0, 1.0, 1.0], failures=1))
        row = report.compare(candidate, [baseline, baseline])[0]
        self.assertEqual("failed", row["status"])

    def test_changed_inputs_are_flagged_rather_than_compared(self) -> None:
        """Same name, different fixture: the numbers are not comparable and must not be."""
        baseline = run(scenario("s", [100.0] * 3, inputs=["a.docx"]))
        candidate = run(scenario("s", [10.0] * 3, inputs=["b.docx"]))
        row = report.compare(candidate, [baseline, baseline])[0]
        self.assertEqual("inputs-changed", row["status"])
        self.assertIn("inputs-changed", report.BLOCKING)

    def test_a_new_scenario_is_neither_a_pass_nor_a_regression(self) -> None:
        baseline = run(scenario("old", [100.0] * 3))
        candidate = run(scenario("old", [100.0] * 3), scenario("new", [100.0] * 3))
        rows = {row["name"]: row for row in report.compare(candidate, [baseline, baseline])}
        self.assertEqual("new", rows["new"]["status"])
        self.assertNotIn("new", report.BLOCKING)

    def test_a_scenario_that_disappeared_blocks(self) -> None:
        """Deleting a slow scenario must not look like an improvement."""
        baseline = run(scenario("kept", [100.0] * 3), scenario("dropped", [100.0] * 3))
        candidate = run(scenario("kept", [100.0] * 3))
        rows = {row["name"]: row for row in report.compare(candidate, [baseline, baseline])}
        self.assertEqual("missing", rows["dropped"]["status"])

    def test_an_unusable_baseline_blocks_rather_than_passing(self) -> None:
        baseline = run(scenario("s", [100.0] * 3, failures=1))
        candidate = run(scenario("s", [100.0] * 3))
        row = report.compare(candidate, [baseline, baseline])[0]
        self.assertEqual("baseline-unusable", row["status"])


class LoadAndRenderTests(unittest.TestCase):
    def test_an_unknown_schema_is_refused(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "r.json"
            path.write_text(json.dumps({"SchemaVersion": "99.0", "Scenarios": []}), encoding="utf-8")
            with self.assertRaises(report.ReportError):
                report.load(path)

    def test_results_without_scenarios_are_refused(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "r.json"
            path.write_text(
                json.dumps({"SchemaVersion": report.SCHEMA_VERSION, "Scenarios": []}), encoding="utf-8"
            )
            with self.assertRaises(report.ReportError):
                report.load(path)

    def test_the_report_states_sample_size_and_failures(self) -> None:
        text = report.render(run(scenario("s", [100.0] * 3, failures=1)), None, 0)
        self.assertIn("Failed iterations: 1", text)
        self.assertIn("Total iterations: 4", text)
        self.assertIn("assertion 0 failed", text)

    def test_the_cli_fails_when_an_iteration_failed(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "r.json"
            path.write_text(json.dumps(run(scenario("s", [1.0], failures=1))), encoding="utf-8")
            out = Path(directory) / "report.md"
            self.assertEqual(1, report.main(["--results", str(path), "--output", str(out)]))
            self.assertTrue(out.exists())

    def test_the_cli_gate_blocks_only_when_asked(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory) / "base.json"
            cand = Path(directory) / "cand.json"
            base.write_text(json.dumps(run(scenario("s", [100.0] * 3))), encoding="utf-8")
            cand.write_text(json.dumps(run(scenario("s", [500.0] * 3))), encoding="utf-8")
            out = Path(directory) / "report.md"
            args = ["--results", str(cand), "--baseline", str(base), str(base), "--output", str(out)]
            # A timing move alone never blocks, with or without the gate flag.
            self.assertEqual(0, report.main(args))
            self.assertEqual(0, report.main(args + ["--fail-on-blocking"]))

    def test_the_cli_gate_blocks_an_allocation_regression(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory) / "base.json"
            cand = Path(directory) / "cand.json"
            out = Path(directory) / "report.md"
            base.write_text(json.dumps(run(scenario("s", [100.0] * 3, allocated=1000))), encoding="utf-8")
            cand.write_text(json.dumps(run(scenario("s", [100.0] * 3, allocated=5000))), encoding="utf-8")
            args = ["--results", str(cand), "--baseline", str(base), str(base), "--output", str(out)]
            self.assertEqual(0, report.main(args))
            self.assertEqual(1, report.main(args + ["--fail-on-blocking"]))


class BaselineRunTests(unittest.TestCase):
    """
    The threshold must come from between-run movement, not from within-run repetitions.

    This is the defect that produced the policy: the first version derived the threshold from
    the spread of repetitions inside one run, and two consecutive runs of the same commit
    reported two scenarios as regressed by 24%.
    """

    def test_a_single_baseline_run_is_marked_provisional_and_gets_the_floor(self) -> None:
        aggregated = report.aggregate_baseline([run(scenario("s", [100.0] * 15))])
        entry = aggregated["s"]
        self.assertTrue(entry["provisional"])
        self.assertEqual(report.MINIMUM_TIME_THRESHOLD, report.threshold_for(entry, metric="time"))

    def test_within_run_spread_does_not_widen_the_threshold(self) -> None:
        """A single run that wobbles internally still gets the floor, not a wide allowance."""
        aggregated = report.aggregate_baseline([run(scenario("s", [50.0, 100.0, 150.0]))])
        self.assertEqual(
            report.MINIMUM_TIME_THRESHOLD,
            report.threshold_for(aggregated["s"], metric="time"))

    def test_between_run_movement_widens_the_threshold(self) -> None:
        runs = [
            run(scenario("s", [100.0] * 3)),
            run(scenario("s", [140.0] * 3)),
            run(scenario("s", [180.0] * 3)),
        ]
        entry = report.aggregate_baseline(runs)["s"]
        self.assertFalse(entry["provisional"])
        self.assertGreater(entry["time_spread"], 0.0)
        self.assertGreater(report.threshold_for(entry, metric="time"), report.MINIMUM_TIME_THRESHOLD)

    def test_a_steady_baseline_across_runs_keeps_the_floor(self) -> None:
        runs = [run(scenario("s", [100.0] * 3)) for _ in range(4)]
        entry = report.aggregate_baseline(runs)["s"]
        self.assertEqual(0.0, entry["time_spread"])
        self.assertEqual(report.MINIMUM_TIME_THRESHOLD, report.threshold_for(entry, metric="time"))

    def test_a_scenario_absent_from_one_run_is_unusable(self) -> None:
        """A baseline stitched from whichever runs had a scenario compares two populations."""
        runs = [run(scenario("a", [100.0] * 3), scenario("b", [100.0] * 3)),
                run(scenario("a", [100.0] * 3))]
        aggregated = report.aggregate_baseline(runs)
        self.assertFalse(aggregated["a"]["unusable"])
        self.assertTrue(aggregated["b"]["unusable"])

    def test_a_failure_in_any_baseline_run_makes_it_unusable(self) -> None:
        runs = [run(scenario("s", [100.0] * 3)), run(scenario("s", [100.0] * 3, failures=1))]
        self.assertTrue(report.aggregate_baseline(runs)["s"]["unusable"])

    def test_the_median_is_taken_across_run_medians(self) -> None:
        runs = [run(scenario("s", [100.0] * 3)),
                run(scenario("s", [200.0] * 3)),
                run(scenario("s", [300.0] * 3))]
        self.assertEqual(200.0, report.aggregate_baseline(runs)["s"]["median_ms"])


def protocol(calls: int, sent: int = 1000, received: int = 5000) -> dict:
    return {"ToolCalls": calls, "BytesSent": sent, "BytesReceived": received,
            "FramesSent": calls, "FramesReceived": calls}


class ProtocolTests(unittest.TestCase):
    """
    Protocol cost is the one metric on the MCP path worth gating.

    Measured over five iterations against an installed server, tool calls and bytes sent were
    identical every time and bytes received moved about 1.2%. So calls are compared exactly and
    bytes get a threshold.
    """

    def test_schema_1_1_is_accepted_alongside_1_0(self) -> None:
        """Published 1.0 baselines must keep loading after the protocol fields were added."""
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "r.json"
            document = run(scenario("s", [1.0]))
            document["SchemaVersion"] = "1.1"
            path.write_text(json.dumps(document), encoding="utf-8")
            self.assertEqual("1.1", report.load(path)["SchemaVersion"])

    def test_protocol_counters_are_summarised_when_present(self) -> None:
        summary = report.summarize_scenario(
            scenario("s", [1.0, 1.0], protocol=protocol(6, sent=100, received=900)))
        self.assertEqual(6, summary["tool_calls"])
        self.assertEqual(100, summary["median_bytes_sent"])
        self.assertEqual(900, summary["median_bytes_received"])

    def test_a_direct_scenario_carries_no_protocol_figures(self) -> None:
        summary = report.summarize_scenario(scenario("s", [1.0, 1.0]))
        self.assertIsNone(summary["tool_calls"])

    def test_one_extra_round_trip_blocks(self) -> None:
        """An adapter that needs more calls for the same work is a regression, exactly."""
        baseline = run(scenario("s", [100.0] * 3, protocol=protocol(6)))
        candidate = run(scenario("s", [100.0] * 3, protocol=protocol(7)))
        row = report.compare(candidate, [baseline, baseline])[0]
        self.assertEqual("protocol-regressed", row["status"])
        self.assertEqual(1, row["tool_call_change"])
        self.assertIn("protocol-regressed", report.BLOCKING)

    def test_fewer_round_trips_is_reported_and_does_not_block(self) -> None:
        baseline = run(scenario("s", [100.0] * 3, protocol=protocol(6)))
        candidate = run(scenario("s", [100.0] * 3, protocol=protocol(4)))
        row = report.compare(candidate, [baseline, baseline])[0]
        self.assertEqual("fewer-calls", row["status"])
        self.assertNotIn("fewer-calls", report.BLOCKING)

    def test_identical_call_counts_do_not_block(self) -> None:
        baseline = run(scenario("s", [100.0] * 3, protocol=protocol(6)))
        candidate = run(scenario("s", [100.0] * 3, protocol=protocol(6)))
        self.assertEqual("ok", report.compare(candidate, [baseline, baseline])[0]["status"])

    def test_a_payload_that_balloons_blocks_even_with_the_same_call_count(self) -> None:
        baseline = run(scenario("s", [100.0] * 3, protocol=protocol(6, sent=100, received=900)))
        candidate = run(scenario("s", [100.0] * 3, protocol=protocol(6, sent=100, received=9000)))
        row = report.compare(candidate, [baseline, baseline])[0]
        self.assertEqual("protocol-regressed", row["status"])

    def test_small_payload_variance_is_within_the_threshold(self) -> None:
        """Bytes received move about 1.2% between iterations; that must not be a regression."""
        baseline = run(scenario("s", [100.0] * 3, protocol=protocol(6, sent=100, received=1000)))
        candidate = run(scenario("s", [100.0] * 3, protocol=protocol(6, sent=100, received=1012)))
        self.assertEqual("ok", report.compare(candidate, [baseline, baseline])[0]["status"])

    def test_protocol_takes_precedence_over_allocation_and_timing(self) -> None:
        baseline = run(scenario("s", [100.0] * 3, allocated=1000, protocol=protocol(6)))
        candidate = run(scenario("s", [900.0] * 3, allocated=9000, protocol=protocol(9)))
        self.assertEqual(
            "protocol-regressed", report.compare(candidate, [baseline, baseline])[0]["status"])


if __name__ == "__main__":
    unittest.main()
