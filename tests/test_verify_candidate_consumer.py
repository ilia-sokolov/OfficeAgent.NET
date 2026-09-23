"""The candidate consumer check reads stored records and restores published package ids; both must
stay where it expects them, or the check fails for a reason unrelated to the candidate."""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

import verify_candidate_consumer as consumer  # noqa: E402


class CandidateConsumerTests(unittest.TestCase):
    def test_every_stored_record_the_consumer_reads_exists(self) -> None:
        missing = [str(path.relative_to(ROOT)) for path in consumer.REQUIRED_RECORDS if not path.is_file()]
        self.assertEqual([], missing)

    def test_the_consumer_references_only_published_packages(self) -> None:
        packable = {
            project.stem
            for project in (ROOT / "src").glob("*/*.csproj")
            if "<IsPackable>false</IsPackable>" not in project.read_text(encoding="utf-8")
        }
        self.assertTrue(set(consumer.PACKAGES) <= packable, set(consumer.PACKAGES) - packable)

    def test_the_v09_case_the_consumer_replays_is_still_recorded(self) -> None:
        manifest = (consumer.CORPUS_V09 / "manifest.json").read_text(encoding="utf-8")
        self.assertIn('"mixed-runs-tracked-edit"', manifest)

    def test_a_failed_check_fails_the_process(self) -> None:
        self.assertIn("return failures.Count == 0 ? 0 : 1;", consumer.PROGRAM)
        self.assertIn('"candidate-consumer=passed"', consumer.PROGRAM)


if __name__ == "__main__":
    unittest.main()
