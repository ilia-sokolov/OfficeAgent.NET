"""The QuickEdit sample is candidate evidence only when it ran on exactly the requested packages:
present before anything runs, restored into a fresh cache, and resolved at exactly that version."""

from __future__ import annotations

import os
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

import package_samples  # noqa: E402
import verify_quickedit_sample as quickedit  # noqa: E402


def assets(versions: dict[str, str], kind: str = "package") -> dict:
    libraries = {f"{package}/{version}": {"type": kind} for package, version in versions.items()}
    libraries["System.Text.Json/8.0.5"] = {"type": "package"}
    return {"libraries": libraries}


def everyone(version: str) -> dict[str, str]:
    return {package: version for package in quickedit.RESOLVED_PACKAGES}


class ResolvedVersionTests(unittest.TestCase):
    def test_exact_core_and_word_versions_pass(self) -> None:
        self.assertEqual([], quickedit.resolved_version_problems(assets(everyone("1.0.0-rc.5")), "1.0.0-rc.5"))

    def test_the_sample_requires_core_and_word(self) -> None:
        self.assertTrue({"OfficeAgent.Core", "OfficeAgent.Word"} <= set(quickedit.RESOLVED_PACKAGES))

    def test_a_wrong_resolved_version_fails_and_names_the_package(self) -> None:
        versions = everyone("1.0.0-rc.5")
        versions["OfficeAgent.Word"] = "1.0.0-rc.50"
        problems = quickedit.resolved_version_problems(assets(versions), "1.0.0-rc.5")
        self.assertEqual(1, len(problems))
        self.assertIn("OfficeAgent.Word", problems[0])
        self.assertIn("1.0.0-rc.50", problems[0])

    def test_a_higher_upstream_version_fails(self) -> None:
        versions = everyone("1.0.0-rc.5")
        versions["OfficeAgent.Core"] = "1.0.0"
        self.assertIn("OfficeAgent.Core", quickedit.resolved_version_problems(assets(versions), "1.0.0-rc.5")[0])

    def test_a_package_absent_from_the_assets_fails(self) -> None:
        versions = everyone("1.0.0-rc.5")
        del versions["OfficeAgent.Core"]
        problems = quickedit.resolved_version_problems(assets(versions), "1.0.0-rc.5")
        self.assertEqual(["OfficeAgent.Core: resolved nothing, expected 'OfficeAgent.Core/1.0.0-rc.5'"], problems)

    def test_a_project_reference_is_not_a_resolved_package(self) -> None:
        problems = quickedit.resolved_version_problems(assets(everyone("1.0.0-rc.5"), kind="project"), "1.0.0-rc.5")
        self.assertEqual(len(quickedit.RESOLVED_PACKAGES), len(problems))

    def test_versions_compare_exactly(self) -> None:
        self.assertEqual(len(quickedit.RESOLVED_PACKAGES),
                         len(quickedit.resolved_version_problems(assets(everyone("1.0.0-RC.5")), "1.0.0-rc.5")))

    def test_package_ids_compare_as_nuget_does(self) -> None:
        lowered = {package.lower(): "1.0.0-rc.5" for package in quickedit.RESOLVED_PACKAGES}
        self.assertEqual([], quickedit.resolved_version_problems(assets(lowered), "1.0.0-rc.5"))

    def test_an_assets_file_without_libraries_fails_for_every_package(self) -> None:
        self.assertEqual(len(quickedit.RESOLVED_PACKAGES), len(quickedit.resolved_version_problems({}, "1.0.0-rc.5")))


class SampleProjectTests(unittest.TestCase):
    def test_the_packaged_project_references_the_requested_version(self) -> None:
        with tempfile.TemporaryDirectory() as output:
            archive = package_samples.package_quickedit(Path(output), "1.0.0-rc.5")
            with zipfile.ZipFile(archive) as package:
                project = package.read("quickedit-sample/QuickEdit.csproj").decode()
        self.assertEqual([], quickedit.project_version_problems(project, "1.0.0-rc.5"))
        problems = quickedit.project_version_problems(project, "1.0.0-rc.50")
        self.assertEqual(["OfficeAgent.Core", "OfficeAgent.Word"], [problem.split(":", 1)[0] for problem in problems])


class MissingPackageTests(unittest.TestCase):
    def test_missing_candidate_packages_are_rejected_before_anything_runs(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            artifacts = Path(folder)
            package_samples.package_quickedit(artifacts, "1.0.0-rc.5")
            (artifacts / "OfficeAgent.Core.1.0.0-rc.5.nupkg").write_bytes(b"")
            with mock.patch.object(quickedit.subprocess, "run") as run:
                with self.assertRaises(RuntimeError) as raised:
                    quickedit.verify(artifacts, "1.0.0-rc.5", "dotnet")
            run.assert_not_called()
        message = str(raised.exception)
        self.assertIn("OfficeAgent.Word.1.0.0-rc.5.nupkg", message)
        self.assertIn("OfficeAgent.Abstractions.1.0.0-rc.5.nupkg", message)
        self.assertNotIn("OfficeAgent.Core.1.0.0-rc.5.nupkg", message)

    def test_packages_of_another_version_do_not_count(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            artifacts = Path(folder)
            for package in quickedit.RESOLVED_PACKAGES:
                (artifacts / f"{package}.1.0.0-rc.50.nupkg").write_bytes(b"")
            self.assertEqual(len(quickedit.RESOLVED_PACKAGES), len(quickedit.missing_packages(artifacts, "1.0.0-rc.5")))
            self.assertEqual([], quickedit.missing_packages(artifacts, "1.0.0-rc.50"))

    def test_a_sample_packed_for_another_version_is_rejected_before_restore(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            artifacts = Path(folder)
            package_samples.package_quickedit(artifacts, "1.0.0-rc.4")
            for package in quickedit.RESOLVED_PACKAGES:
                (artifacts / f"{package}.1.0.0-rc.5.nupkg").write_bytes(b"")
            with mock.patch.object(quickedit.subprocess, "run") as run:
                with self.assertRaises(RuntimeError) as raised:
                    quickedit.verify(artifacts, "1.0.0-rc.5", "dotnet")
            run.assert_not_called()
        self.assertIn("does not reference version '1.0.0-rc.5'", str(raised.exception))


class IsolatedRestoreTests(unittest.TestCase):
    def test_a_stale_ambient_cache_cannot_satisfy_the_restore(self) -> None:
        ambient = {"NUGET_PACKAGES": "/ambient/nuget/packages", "NUGET_FALLBACK_PACKAGES": "/ambient/fallback"}
        with tempfile.TemporaryDirectory() as folder, mock.patch.dict(os.environ, ambient):
            root = Path(folder)
            artifacts = root / "artifacts"
            env, config = quickedit.restore_environment(root, artifacts)
            cache = Path(env["NUGET_PACKAGES"])
            self.assertEqual(root / "nuget-packages", cache)
            self.assertTrue(cache.is_dir())
            self.assertEqual([], list(cache.iterdir()))
            self.assertNotIn("NUGET_FALLBACK_PACKAGES", env)
            self.assertEqual(str(config), env["RESTORE_CONFIG_FILE"])
            text = config.read_text(encoding="utf-8")
        self.assertIn("<clear />", text)
        local = text.index(f'value="{artifacts.as_posix()}"')
        upstream = text.index("https://api.nuget.org/v3/index.json")
        self.assertLess(local, upstream, "the candidate feed must be the first source")

    def test_restore_uses_the_fresh_cache_and_a_wrong_resolution_stops_the_sample(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            artifacts = Path(folder)
            package_samples.package_quickedit(artifacts, "1.0.0-rc.5")
            for package in quickedit.RESOLVED_PACKAGES:
                (artifacts / f"{package}.1.0.0-rc.5.nupkg").write_bytes(b"")
            calls = []

            def fake_run(command, cwd, env, check):  # noqa: ANN001
                calls.append((command, env))
                if command[1] == "restore":
                    (Path(cwd) / "obj").mkdir()
                    (Path(cwd) / "obj" / "project.assets.json").write_text(
                        '{"libraries": {"OfficeAgent.Core/1.0.0-rc.4": {"type": "package"}}}', encoding="utf-8")
                return mock.Mock(returncode=0)

            with mock.patch.object(quickedit.subprocess, "run", side_effect=fake_run):
                with self.assertRaises(RuntimeError) as raised:
                    quickedit.verify(artifacts, "1.0.0-rc.5", "dotnet")
        self.assertEqual(1, len(calls), "the sample must not run after a wrong resolution")
        self.assertTrue(calls[0][1]["NUGET_PACKAGES"].endswith("nuget-packages"))
        self.assertIn("OfficeAgent.Core: resolved ['OfficeAgent.Core/1.0.0-rc.4']", str(raised.exception))


class CommandTests(unittest.TestCase):
    """The workflow and the runbook hand the same version to packaging and to verification."""

    def test_the_candidate_job_packs_and_verifies_the_same_version(self) -> None:
        workflow = (ROOT / ".github" / "workflows" / "build.yml").read_text(encoding="utf-8")
        job = workflow[workflow.index("  candidate-packages:"):workflow.index("  renderer-sandbox:")]
        self.assertIn('package_samples.py --output ./candidate --version "$CANDIDATE_VERSION"', job)
        self.assertIn('verify_quickedit_sample.py --artifacts ./candidate --version "$CANDIDATE_VERSION"', job)

    def test_the_runbook_packs_and_verifies_the_same_version(self) -> None:
        runbook = (ROOT / "docs" / "releasing.md").read_text(encoding="utf-8")
        self.assertIn('package_samples.py --output ./candidate --version "$CANDIDATE"', runbook)
        self.assertIn('verify_quickedit_sample.py --artifacts ./candidate --version "$CANDIDATE"', runbook)
        self.assertNotIn("verify_quickedit_sample.py --artifacts ./candidate\n", runbook.replace("\r\n", "\n"))

    def test_the_release_workflow_verifies_the_released_version(self) -> None:
        workflow = (ROOT / ".github" / "workflows" / "publish.yml").read_text(encoding="utf-8")
        self.assertIn('verify_quickedit_sample.py --artifacts artifacts --version "${{ steps.release.outputs.version }}"',
                      workflow)


if __name__ == "__main__":
    unittest.main()
