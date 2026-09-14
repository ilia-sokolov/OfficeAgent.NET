from __future__ import annotations

import json
import sys
import tempfile
import unittest
import xml.etree.ElementTree as ET
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import release_evidence  # noqa: E402


class ReleaseEvidenceTests(unittest.TestCase):
    VERSION = "0.9.0"
    SHA = "1" * 40
    REF = "refs/tags/v0.9.0"

    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.output = Path(self.temporary.name)
        self.inventory = {
            "schemaVersion": 1,
            "repository": "owner/repository",
            "workflow": ".github/workflows/publish.yml",
            "sbomTool": {
                "package": "CycloneDX",
                "version": "5.5.0",
                "format": "CycloneDX",
                "specVersion": "1.6",
            },
            "packages": [
                {
                    "id": "OfficeAgent.Test",
                    "project": "src/OfficeAgent.Test/OfficeAgent.Test.csproj",
                    "frameworks": ["netstandard2.0", "net8.0"],
                }
            ],
            "skills": [
                {
                    "name": "test-skill",
                    "archive": "test-skill.zip",
                    "source": "skills/test-skill",
                    "license": "MIT",
                },
                {
                    "name": "integration-skill",
                    "archive": "integration-skill.zip",
                    "source": "skills/integration-skill",
                    "license": "MIT",
                },
            ],
            "container": {
                "name": "ghcr.io/owner/test",
                "workflow": ".github/workflows/publish-image.yml",
            },
        }
        for product in release_evidence.expected_products(self.inventory, self.VERSION):
            (self.output / product["name"]).write_bytes(product["name"].encode())
        for name in release_evidence.expected_sboms(self.inventory, self.VERSION):
            component, framework = release_evidence.sbom_identity(
                self.inventory, name, self.VERSION
            )
            bom = {
                "bomFormat": "CycloneDX",
                "specVersion": "1.6",
                "version": 1,
                "metadata": {
                    "component": {
                        "type": "library",
                        "name": component,
                        "version": self.VERSION,
                        "properties": [
                            {"name": "officeagent:targetFramework", "value": framework}
                        ],
                    }
                },
                "components": [],
                "dependencies": [],
            }
            (self.output / name).write_text(json.dumps(bom), encoding="utf-8")
        release_evidence.create_evidence(
            self.inventory, self.VERSION, self.SHA, self.REF, self.output
        )

    def verify(self, *, release_assets: set[str] | None = None) -> None:
        release_evidence.verify_evidence(
            self.inventory,
            self.output,
            self.inventory["repository"],
            self.inventory["workflow"],
            self.REF,
            self.SHA,
            release_assets,
        )

    def test_genuine_release_evidence_passes(self) -> None:
        manifest = release_evidence.read_json(self.output / "release-manifest.json")
        self.verify(release_assets=set(manifest["releaseAssets"]))

    def test_altered_artifact_fails(self) -> None:
        path = self.output / f"OfficeAgent.Test.{self.VERSION}.nupkg"
        path.write_bytes(path.read_bytes() + b"tampered")
        with self.assertRaisesRegex(release_evidence.EvidenceError, "digest"):
            self.verify()

    def test_wrong_source_identity_fails(self) -> None:
        with self.assertRaisesRegex(release_evidence.EvidenceError, "identity"):
            release_evidence.verify_evidence(
                self.inventory,
                self.output,
                self.inventory["repository"],
                self.inventory["workflow"],
                self.REF,
                "2" * 40,
            )

    def test_wrong_source_ref_fails(self) -> None:
        with self.assertRaisesRegex(release_evidence.EvidenceError, "identity"):
            release_evidence.verify_evidence(
                self.inventory,
                self.output,
                self.inventory["repository"],
                self.inventory["workflow"],
                "refs/tags/v0.9.1",
                self.SHA,
            )

    def test_missing_sbom_entry_fails(self) -> None:
        path = self.output / "release-manifest.json"
        manifest = release_evidence.read_json(path)
        manifest["verificationMaterials"].pop()
        path.write_text(json.dumps(manifest), encoding="utf-8")
        with self.assertRaisesRegex(release_evidence.EvidenceError, "SBOM inventory"):
            self.verify()

    def test_incomplete_release_attachment_fails(self) -> None:
        manifest = release_evidence.read_json(self.output / "release-manifest.json")
        assets = set(manifest["releaseAssets"])
        assets.remove(f"OfficeAgent.Test.{self.VERSION}.net8.0.cdx.json")
        with self.assertRaisesRegex(release_evidence.EvidenceError, "attachments differ"):
            self.verify(release_assets=assets)

    def test_each_skill_has_its_own_product_and_sbom(self) -> None:
        manifest = release_evidence.read_json(self.output / "release-manifest.json")
        artifacts = {item["name"]: item for item in manifest["artifacts"]}
        self.assertEqual(
            artifacts["integration-skill.zip"]["dependencyInventory"],
            [f"integration-skill.{self.VERSION}.cdx.json"],
        )
        self.assertIn(
            f"integration-skill.{self.VERSION}.cdx.json",
            {item["name"] for item in manifest["verificationMaterials"]},
        )


class RepositoryReleaseInventoryTests(unittest.TestCase):
    def test_inventory_covers_every_packable_source_project_and_framework(self) -> None:
        inventory = release_evidence.load_inventory(release_evidence.DEFAULT_INVENTORY)
        inventoried = {package["project"]: package for package in inventory["packages"]}
        projects = sorted((ROOT / "src").glob("*/*.csproj"))
        relative_projects = {project.relative_to(ROOT).as_posix() for project in projects}
        self.assertEqual(relative_projects, set(inventoried))

        for relative, package in inventoried.items():
            root = ET.parse(ROOT / relative).getroot()
            frameworks = root.findtext(".//TargetFrameworks") or root.findtext(".//TargetFramework")
            self.assertIsNotNone(frameworks)
            self.assertEqual(frameworks.split(";"), package["frameworks"])

    def test_tool_manifest_pins_expected_cyclonedx_version(self) -> None:
        inventory = release_evidence.load_inventory(release_evidence.DEFAULT_INVENTORY)
        tools = release_evidence.read_json(ROOT / ".config" / "dotnet-tools.json")["tools"]
        self.assertEqual(inventory["sbomTool"]["version"], tools["cyclonedx"]["version"])
        self.assertFalse(tools["cyclonedx"]["rollForward"])

    def test_inventory_covers_every_packaged_skill(self) -> None:
        inventory = release_evidence.load_inventory(release_evidence.DEFAULT_INVENTORY)
        expected = {path.name for path in (ROOT / "skills").iterdir() if path.is_dir()}
        actual = {skill["name"] for skill in inventory["skills"]}
        self.assertEqual(expected, actual)
        for skill in inventory["skills"]:
            self.assertEqual(skill["archive"], f"{skill['name']}.zip")
            self.assertEqual(skill["source"], f"skills/{skill['name']}")
            self.assertTrue((ROOT / skill["source"] / "SKILL.md").is_file())


if __name__ == "__main__":
    unittest.main()
