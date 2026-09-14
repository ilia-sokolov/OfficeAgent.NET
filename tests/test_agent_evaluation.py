from __future__ import annotations

import copy
import hashlib
import json
import sys
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
EVALUATION = ROOT / "evaluations" / "agent-selection" / "v0.9.0"
sys.path.insert(0, str(ROOT / "scripts"))
import agent_evaluation  # noqa: E402


class AgentEvaluationTests(unittest.TestCase):
    def setUp(self) -> None:
        self.protocol = agent_evaluation.load_json(EVALUATION / "protocol.json")
        self.fixture_path = EVALUATION / "fixtures" / "run-records.json"
        self.live_path = EVALUATION / "live-run-records.json"

    def test_protocol_keeps_blind_prompts_clean_and_sessions_isolated(self) -> None:
        agent_evaluation.validate_protocol(self.protocol)
        preferred = self.protocol["preferred_library"].casefold()
        for condition in self.protocol["conditions"]:
            if not condition["blind"]:
                continue
            for task in self.protocol["tasks"]:
                prompt = agent_evaluation.render_prompt(
                    self.protocol, task["id"], condition["id"]
                )
                self.assertNotIn(preferred, prompt.casefold())
        self.assertEqual(
            self.protocol["live_protocol"]["session_policy"],
            "fresh_session_per_attempt_no_shared_history",
        )
        snapshots = self.protocol["controlled_materials"][
            "officeagent_documentation_snapshots"
        ]
        self.assertEqual(5, len(snapshots["files"]))
        self.assertTrue(all(item["after_blob"] for item in snapshots["files"]))

    def test_fixed_records_cover_positive_failed_and_incomplete_outcomes(self) -> None:
        records = agent_evaluation.load_records([self.fixture_path], self.protocol)
        statuses = {record["status"] for record in records}
        self.assertTrue(
            {
                "succeeded",
                "failed",
                "refused",
                "timed_out",
                "invalid_output",
                "not_run",
            }.issubset(statuses)
        )
        self.assertTrue(all(record["provenance"] == "replay_fixture" for record in records))

    def test_fixed_transcript_hashes_match_retained_bodies(self) -> None:
        records = agent_evaluation.load_records([self.fixture_path], self.protocol)
        transcripts = json.loads(
            (EVALUATION / "fixtures" / "transcripts.json").read_text(encoding="utf-8")
        )
        by_identity = {item["identity"]: item for item in transcripts}
        self.assertEqual(len(transcripts), len(by_identity))
        for record in records:
            transcript = by_identity[record["transcript"]["identity"]]
            encoded = json.dumps(
                transcript, ensure_ascii=False, separators=(",", ":"), sort_keys=True
            ).encode("utf-8")
            self.assertEqual(
                hashlib.sha256(encoded).hexdigest(), record["transcript"]["sha256"]
            )
            self.assertTrue(record["transcript"]["retained"])

    def test_aggregation_keeps_denominators_and_live_missing_cells(self) -> None:
        records = agent_evaluation.load_records(
            [self.fixture_path, self.live_path], self.protocol
        )
        summary = agent_evaluation.aggregate(self.protocol, records)
        self.assertEqual(7, summary["replay"]["attempts"])
        self.assertEqual(4, summary["replay"]["selection_observed"])
        self.assertEqual(3, summary["replay"]["preferred_selected"])
        self.assertEqual(4, summary["replay"]["completion_evaluated"])
        self.assertEqual(2, summary["replay"]["completion_verified"])
        self.assertEqual(1, summary["replay"]["preservation_defects"])
        self.assertEqual(2, summary["replay"]["cost_measured_attempts"])
        self.assertEqual(6, summary["replay"]["elapsed_measured_attempts"])
        self.assertEqual(
            3,
            summary["replay_by_condition"]["explicit-officeagent-integration"][
                "attempts"
            ],
        )
        self.assertEqual(
            0,
            summary["replay_by_condition"]["unprompted-library-choice"][
                "preferred_selected"
            ],
        )
        self.assertEqual(54, summary["live_coverage"]["expected"])
        self.assertEqual(0, summary["live_coverage"]["recorded"])
        self.assertEqual(54, summary["live_coverage"]["missing"])

    def test_invalid_record_and_blind_prompt_leak_fail_closed(self) -> None:
        records = json.loads(self.fixture_path.read_text(encoding="utf-8"))
        invalid = copy.deepcopy(records[0])
        invalid["prompt_sha256"] = "0" * 64
        with self.assertRaisesRegex(agent_evaluation.EvaluationError, "frozen prompt"):
            agent_evaluation.validate_record(invalid, self.protocol)

        leaked = copy.deepcopy(self.protocol)
        leaked["conditions"][0]["prompt"] += " Use OfficeAgent.NET."
        with self.assertRaisesRegex(agent_evaluation.EvaluationError, "leaks preferred"):
            agent_evaluation.validate_protocol(leaked)

        reused = copy.deepcopy(records[:2])
        reused[1]["session"]["identity"] = reused[0]["session"]["identity"]
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / "reused-session.json"
            path.write_text(json.dumps(reused), encoding="utf-8")
            with self.assertRaisesRegex(agent_evaluation.EvaluationError, "reused session"):
                agent_evaluation.load_records([path], self.protocol)

    def test_report_regenerates_byte_for_byte(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            output = Path(temporary) / "report.md"
            agent_evaluation.run(
                EVALUATION / "protocol.json",
                [self.fixture_path, self.live_path],
                output,
            )
            self.assertEqual(
                (EVALUATION / "report.md").read_bytes(),
                output.read_bytes(),
            )


if __name__ == "__main__":
    unittest.main()
