"""The candidate consumer check reads stored records and restores published package ids; both must
stay where it expects them, or the check fails for a reason unrelated to the candidate."""

from __future__ import annotations

import sys
import tempfile
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


def everyone(version: str | None) -> dict[str, str | None]:
    return {name: version for name in consumer.ASSEMBLIES}


class ExactVersionTests(unittest.TestCase):
    """Only the source-control suffix after '+' is removed; the rest must be equal."""

    def test_the_candidate_with_its_source_control_suffix_passes(self) -> None:
        self.assertEqual([], consumer.version_mismatches(everyone("1.0.0-rc.3+8264f64"), "1.0.0-rc.3"))

    def test_a_longer_prerelease_label_fails(self) -> None:
        problems = consumer.version_mismatches(everyone("1.0.0-rc.30+8264f64"), "1.0.0-rc.3")
        self.assertEqual(len(consumer.ASSEMBLIES), len(problems))
        self.assertIn("1.0.0-rc.30+8264f64", problems[0])

    def test_a_prerelease_fails_for_the_final_version(self) -> None:
        self.assertEqual(len(consumer.ASSEMBLIES),
                         len(consumer.version_mismatches(everyone("1.0.0-rc.3+8264f64"), "1.0.0")))

    def test_the_final_version_with_its_suffix_passes(self) -> None:
        self.assertEqual([], consumer.version_mismatches(everyone("1.0.0+8264f64"), "1.0.0"))

    def test_a_missing_informational_version_fails_and_names_the_assembly(self) -> None:
        observed = everyone("1.0.0")
        observed["OfficeAgent.Word"] = None
        self.assertEqual(["OfficeAgent.Word: no informational version"], consumer.version_mismatches(observed, "1.0.0"))

    def test_an_assembly_the_consumer_did_not_report_fails(self) -> None:
        observed = everyone("1.0.0")
        del observed["OfficeAgent.Core"]
        self.assertEqual(["OfficeAgent.Core: not reported by the consumer"], consumer.version_mismatches(observed, "1.0.0"))

    def test_leading_whitespace_fails(self) -> None:
        self.assertEqual(len(consumer.ASSEMBLIES),
                         len(consumer.version_mismatches(everyone(" 1.0.0-rc.4+8264f64"), "1.0.0-rc.4")))

    def test_trailing_whitespace_before_the_suffix_fails(self) -> None:
        self.assertEqual(len(consumer.ASSEMBLIES),
                         len(consumer.version_mismatches(everyone("1.0.0-rc.4 +8264f64"), "1.0.0-rc.4")))

    def test_trailing_whitespace_without_a_suffix_fails(self) -> None:
        self.assertEqual(len(consumer.ASSEMBLIES),
                         len(consumer.version_mismatches(everyone("1.0.0-rc.4 "), "1.0.0-rc.4")))

    def test_trailing_whitespace_after_the_build_metadata_fails(self) -> None:
        problems = consumer.version_mismatches(everyone("1.0.0-rc.4+8264f64 "), "1.0.0-rc.4")
        self.assertEqual(len(consumer.ASSEMBLIES), len(problems))
        self.assertIn("'1.0.0-rc.4+8264f64 '", problems[0])

    def test_whitespace_only_build_metadata_fails(self) -> None:
        self.assertEqual("1.0.0-rc.4+ ", consumer.normalized_informational_version("1.0.0-rc.4+ "))
        self.assertEqual(len(consumer.ASSEMBLIES),
                         len(consumer.version_mismatches(everyone("1.0.0-rc.4+ "), "1.0.0-rc.4")))

    def test_empty_build_metadata_fails(self) -> None:
        self.assertEqual(len(consumer.ASSEMBLIES),
                         len(consumer.version_mismatches(everyone("1.0.0-rc.4+"), "1.0.0-rc.4")))

    def test_a_whitespace_only_version_fails_as_a_mismatch(self) -> None:
        problems = consumer.version_mismatches(everyone("   "), "1.0.0-rc.4")
        self.assertEqual(len(consumer.ASSEMBLIES), len(problems))
        self.assertIn("informational version '   '", problems[0])

    def test_dotted_build_metadata_passes(self) -> None:
        self.assertEqual("1.0.0-rc.4", consumer.normalized_informational_version("1.0.0-rc.4+sha.032a68a.build-7"))

    def test_an_empty_informational_version_is_missing(self) -> None:
        self.assertIsNone(consumer.normalized_informational_version(""))
        self.assertIsNone(consumer.normalized_informational_version(None))

    def test_case_is_significant(self) -> None:
        self.assertEqual(len(consumer.ASSEMBLIES),
                         len(consumer.version_mismatches(everyone("1.0.0-RC.4+8264f64"), "1.0.0-rc.4")))

    def test_the_rc4_candidate_with_its_commit_passes(self) -> None:
        self.assertEqual([], consumer.version_mismatches(everyone("1.0.0-rc.4+032a68a"), "1.0.0-rc.4"))

    def test_reported_lines_are_parsed_per_assembly(self) -> None:
        stdout = "PASS x\nASSEMBLY\tOfficeAgent.Core\t1.0.0-rc.3+abc\nASSEMBLY\tOfficeAgent.Word\t\n"
        self.assertEqual({"OfficeAgent.Core": "1.0.0-rc.3+abc", "OfficeAgent.Word": None},
                         consumer.reported_assembly_versions(stdout))


class RuntimeTests(unittest.TestCase):
    def test_no_requirement_accepts_any_runtime(self) -> None:
        self.assertIsNone(consumer.runtime_mismatch("RUNTIME\t8.0.21\n", None))

    def test_the_required_major_passes_and_another_fails(self) -> None:
        self.assertIsNone(consumer.runtime_mismatch("RUNTIME\t10.0.11\n", "10"))
        self.assertIn("8.0.21", consumer.runtime_mismatch("RUNTIME\t8.0.21\n", "10"))

    def test_an_unreported_runtime_fails_when_one_is_required(self) -> None:
        self.assertIsNotNone(consumer.runtime_mismatch("PASS x\n", "10"))


class CandidateWorkflowTests(unittest.TestCase):
    """Ordinary CI packs the committed version; only the candidate job is candidate evidence, and it
    must hand the requested version to every step rather than let a verifier infer another."""

    def job(self) -> str:
        workflow = (ROOT / ".github" / "workflows" / "build.yml").read_text(encoding="utf-8")
        return workflow[workflow.index("  candidate-packages:"):workflow.index("  renderer-sandbox:")]

    def test_the_candidate_version_reaches_the_pack_and_every_verifier(self) -> None:
        job = self.job()
        for step in ('-p:Version="$CANDIDATE_VERSION"',
                     'smoke_packaged_artifacts.py --artifacts ./candidate --version "$CANDIDATE_VERSION"',
                     'verify_candidate_consumer.py --artifacts ./candidate --version "$CANDIDATE_VERSION"',
                     'package_samples.py --output ./candidate --version "$CANDIDATE_VERSION"',
                     'verify_quickedit_sample.py --artifacts ./candidate --version "$CANDIDATE_VERSION"',
                     'verify_integration_kit.py --artifacts ./candidate --version "$CANDIDATE_VERSION"'):
            self.assertIn(step, job)

    def test_the_job_requires_all_nine_packages_and_covers_the_runtime_matrix(self) -> None:
        job = self.job()
        self.assertIn("test \"$count\" -eq 9", job)
        for cell in ("{ os: ubuntu-latest, runtime: '8' }", "{ os: windows-latest, runtime: '8' }",
                     "{ os: macos-latest, runtime: '8' }", "{ os: ubuntu-latest, runtime: '10' }",
                     "{ os: windows-latest, runtime: '10' }"):
            self.assertIn(cell, job)
        self.assertIn("OFFICEAGENT_EXPECT_RUNTIME_MAJOR: ${{ matrix.runtime }}", job)


class MissingCandidateTests(unittest.TestCase):
    def test_a_run_fails_when_the_requested_candidate_packages_are_absent(self) -> None:
        with tempfile.TemporaryDirectory() as empty:
            with self.assertRaises(consumer.SmokeError) as raised:
                consumer.verify(Path(empty), "1.0.0-rc.3", "dotnet")
        self.assertIn("OfficeAgent.Core.1.0.0-rc.3.nupkg", str(raised.exception))


if __name__ == "__main__":
    unittest.main()
