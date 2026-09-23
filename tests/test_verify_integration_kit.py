"""The integration kit's recipes are candidate evidence only when they ran on exactly the requested
packages: present, restored into a fresh cache, and resolved at exactly that version before they run."""

from __future__ import annotations

import json
import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

import package_skills  # noqa: E402
import release_evidence  # noqa: E402
import verify_integration_kit as kit  # noqa: E402


def artifacts_with(folder: Path, version: str) -> Path:
    artifacts = folder / "artifacts"
    package_skills.package_skills(release_evidence.load_inventory(release_evidence.DEFAULT_INVENTORY), artifacts)
    for package in kit.RESOLVED_PACKAGES:
        (artifacts / f"{package}.{version}.nupkg").write_bytes(b"")
    return artifacts


class IntegrationKitRestoreTests(unittest.TestCase):
    def test_missing_candidate_packages_are_rejected_before_anything_runs(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            artifacts = artifacts_with(Path(folder), "1.0.0-rc.50")
            with mock.patch.object(kit.subprocess, "run") as run:
                with self.assertRaises(kit.VerificationError) as raised:
                    kit.verify(artifacts, "1.0.0-rc.5", "dotnet")
            run.assert_not_called()
        self.assertIn("OfficeAgent.Core.1.0.0-rc.5.nupkg", str(raised.exception))

    def test_a_stale_ambient_cache_cannot_satisfy_the_restore(self) -> None:
        ambient = {"NUGET_PACKAGES": "/ambient/nuget/packages", "NUGET_FALLBACK_PACKAGES": "/ambient/fallback"}
        with tempfile.TemporaryDirectory() as folder, mock.patch.dict(os.environ, ambient):
            root = Path(folder)
            consumer = root / "consumer"
            consumer.mkdir()
            artifacts = root / "artifacts"
            env = kit.restore_environment(root, consumer, artifacts)
            cache = Path(env["NUGET_PACKAGES"])
            self.assertEqual(root / "nuget-packages", cache)
            self.assertEqual([], list(cache.iterdir()))
            self.assertNotIn("NUGET_FALLBACK_PACKAGES", env)
            self.assertEqual(str(consumer / "NuGet.Config"), env["RESTORE_CONFIG_FILE"])
            text = (consumer / "NuGet.Config").read_text(encoding="utf-8")
        self.assertIn("<clear />", text)
        self.assertLess(text.index(f'value="{artifacts.as_posix()}"'), text.index("https://api.nuget.org/v3/index.json"),
                        "the candidate feed must be the first source")

    def test_a_wrong_resolution_stops_the_recipes_and_restore_uses_the_fresh_cache(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            artifacts = artifacts_with(Path(folder), "1.0.0-rc.5")
            calls = []

            def fake_run(command, cwd, env, text, capture_output, check):  # noqa: ANN001
                calls.append((command, env))
                if command[1] == "restore":
                    (Path(cwd) / "obj").mkdir()
                    (Path(cwd) / "obj" / "project.assets.json").write_text(
                        '{"libraries": {"OfficeAgent.Core/1.0.0-rc.4": {"type": "package"}}}', encoding="utf-8")
                return mock.Mock(returncode=0, stdout="", stderr="")

            with mock.patch.object(kit.subprocess, "run", side_effect=fake_run):
                with self.assertRaises(kit.VerificationError) as raised:
                    kit.verify(artifacts, "1.0.0-rc.5", "dotnet")
        self.assertEqual(["restore"], [command[1] for command, _ in calls], "the recipes must not run")
        self.assertTrue(calls[0][1]["NUGET_PACKAGES"].endswith("nuget-packages"))
        self.assertIn("OfficeAgent.Core: resolved ['OfficeAgent.Core/1.0.0-rc.4']", str(raised.exception))

    def test_the_exact_candidate_resolution_lets_the_recipes_run(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            artifacts = artifacts_with(Path(folder), "1.0.0-rc.5")
            calls = []

            def fake_run(command, cwd, env, text, capture_output, check):  # noqa: ANN001
                calls.append(command[1])
                if command[1] == "restore":
                    (Path(cwd) / "obj").mkdir()
                    libraries = {f"{package}/1.0.0-rc.5": {"type": "package"} for package in kit.RESOLVED_PACKAGES}
                    (Path(cwd) / "obj" / "project.assets.json").write_text(
                        json.dumps({"libraries": libraries}), encoding="utf-8")
                return mock.Mock(returncode=0, stdout="\n".join(sorted(kit.EXPECTED_OUTPUT)), stderr="")

            with mock.patch.object(kit.subprocess, "run", side_effect=fake_run):
                kit.verify(artifacts, "1.0.0-rc.5", "dotnet")
        self.assertEqual(["restore", "run"], calls)


class CommandTests(unittest.TestCase):
    """Every workflow that verifies installed packages names the version it packed."""

    def test_the_candidate_job_passes_the_candidate_version(self) -> None:
        workflow = (ROOT / ".github" / "workflows" / "build.yml").read_text(encoding="utf-8")
        job = workflow[workflow.index("  candidate-packages:"):workflow.index("  renderer-sandbox:")]
        self.assertIn('verify_integration_kit.py --artifacts ./candidate --version "$CANDIDATE_VERSION"', job)

    def test_the_release_workflow_passes_the_released_version(self) -> None:
        workflow = (ROOT / ".github" / "workflows" / "publish.yml").read_text(encoding="utf-8")
        self.assertIn('verify_integration_kit.py --artifacts artifacts --version "${{ steps.release.outputs.version }}"',
                      workflow)


if __name__ == "__main__":
    unittest.main()
