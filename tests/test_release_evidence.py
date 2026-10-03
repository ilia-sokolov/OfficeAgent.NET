from __future__ import annotations

import io
import json
import sys
import tarfile
import tempfile
import unittest
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import release_evidence  # noqa: E402


def fake_single_file(rid: str, version: str, framework: str = "net10.0") -> bytes:
    """Bytes laid out like a .NET single-file executable: host, marker, deps.json, header."""
    deps = json.dumps({
        "runtimeTarget": {"name": f".NETCoreApp,Version=v{framework.removeprefix('net')}/{rid}"},
        "libraries": {
            f"OfficeAgent.Test/{version}": {"type": "project"},
            f"OfficeAgent.Dependency/{version}": {"type": "project"},
            "ModelContextProtocol/1.4.0": {"type": "package"},
            f"runtimepack.Microsoft.NETCore.App.Runtime.{rid}/10.0.5": {"type": "runtimepack"},
            "Reference.Only/1.0.0": {"type": "reference"},
        },
    }).encode()
    host = b"host" + bytes(8) + release_evidence.SINGLE_FILE_MARKER + b"code"
    header_offset = len(host) + len(deps)
    header = (
        (6).to_bytes(4, "little") + (0).to_bytes(4, "little") + (1).to_bytes(4, "little")
        + bytes([3]) + b"abc"
        + len(host).to_bytes(8, "little") + len(deps).to_bytes(8, "little") + bytes(24)
    )
    content = bytearray(host + deps + header)
    content[4:12] = header_offset.to_bytes(8, "little")
    return bytes(content)


def write_standalone(path: Path, member: str, content: bytes, mode: int = 0o755) -> None:
    if path.name.endswith(".tar.gz"):
        with tarfile.open(path, "w:gz") as archive:
            info = tarfile.TarInfo(member)
            info.size, info.mode = len(content), mode
            archive.addfile(info, io.BytesIO(content))
        return
    with zipfile.ZipFile(path, "w") as archive:
        info = zipfile.ZipInfo(member)
        info.create_system = 3
        info.external_attr = (0o100000 | mode) << 16
        archive.writestr(info, content)
        if path.suffix == ".mcpb":
            archive.writestr("manifest.json", json.dumps({
                "version": ReleaseEvidenceTests.VERSION, "server": {"entry_point": member},
            }))


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
                },
                {
                    "id": "OfficeAgent.Dependency",
                    "project": "src/OfficeAgent.Dependency/OfficeAgent.Dependency.csproj",
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
            "samples": [
                {
                    "name": "test-sample",
                    "archive": "test-sample.zip",
                    "license": "MIT",
                }
            ],
            "standalone": {
                "name": "test-server",
                "project": "src/OfficeAgent.Test/OfficeAgent.Test.csproj",
                "framework": "net8.0",
                "license": "MIT",
                "runtimes": ["win-x64", "linux-x64"],
                "bundles": ["win-x64"],
            },
            "container": {
                "name": "ghcr.io/owner/test",
                "workflow": ".github/workflows/publish.yml",
            },
        }
        standalone = self.inventory["standalone"]
        standalone_files = {}
        for rid in standalone["runtimes"]:
            standalone_files[release_evidence.standalone_archive(standalone, rid)] = (
                release_evidence.standalone_archive_member(standalone, rid), rid)
            standalone_files[release_evidence.standalone_bundle(standalone, rid)] = (
                release_evidence.standalone_bundle_member(standalone, rid), rid)
        for product in release_evidence.expected_products(self.inventory, self.VERSION):
            path = self.output / product["name"]
            if product["name"] in standalone_files:
                member, rid = standalone_files[product["name"]]
                write_standalone(path, member, fake_single_file(rid, self.VERSION, "net8.0"))
            elif product["name"] == "test-sample.zip":
                with zipfile.ZipFile(path, "w") as archive:
                    archive.writestr(
                        "test-sample/Test.csproj",
                        '<Project><ItemGroup><PackageReference Include="OfficeAgent.Test" '
                        'Version="0.9.0" /></ItemGroup></Project>',
                    )
            elif product["name"].endswith(".nupkg"):
                package_id = product["name"].removesuffix(f".{self.VERSION}.nupkg")
                dependency = (
                    '<dependency id="OfficeAgent.Dependency" version="[0.9.0, )" />'
                    if package_id == "OfficeAgent.Test" else ""
                )
                nuspec = f"""<?xml version="1.0"?>
<package xmlns="http://schemas.microsoft.com/packaging/2013/05/nuspec.xsd">
  <metadata><id>{package_id}</id><version>{self.VERSION}</version><dependencies>
    <group targetFramework=".NETStandard2.0">{dependency}</group>
    <group targetFramework="net8.0">{dependency}</group>
  </dependencies></metadata>
</package>"""
                with zipfile.ZipFile(path, "w") as archive:
                    archive.writestr(f"{package_id}.nuspec", nuspec)
            else:
                path.write_bytes(product["name"].encode())
        for name in release_evidence.expected_sboms(self.inventory, self.VERSION):
            rid = next((rid for rid in standalone["runtimes"]
                        if name == release_evidence.standalone_sbom(standalone, rid, self.VERSION)), None)
            if rid:
                bom = release_evidence.build_standalone_sbom(
                    self.inventory, rid, self.VERSION, fake_single_file(rid, self.VERSION, "net8.0"))
                (self.output / name).write_text(json.dumps(bom), encoding="utf-8")
                continue
            component, framework = release_evidence.sbom_identity(
                self.inventory, name, self.VERSION
            )
            expected_dependencies = release_evidence.expected_sbom_dependencies(
                self.inventory, self.output, name, self.VERSION
            )
            root_ref = f"pkg:generic/{component}@{self.VERSION}"
            dependency_components = [
                {
                    "type": "library",
                    "name": package_id,
                    "version": package_version,
                    "bom-ref": f"pkg:nuget/{package_id}@{package_version}",
                }
                for package_id, package_version in expected_dependencies.items()
            ]
            bom = {
                "bomFormat": "CycloneDX",
                "specVersion": "1.6",
                "serialNumber": release_evidence.sbom_serial_number(name),
                "version": 1,
                "metadata": {
                    "component": {
                        "type": "library",
                        "name": component,
                        "version": self.VERSION,
                        "bom-ref": root_ref,
                        "properties": [
                            {"name": "officeagent:targetFramework", "value": framework}
                        ],
                    }
                },
                "components": dependency_components,
                "dependencies": [{
                    "ref": root_ref,
                    "dependsOn": [item["bom-ref"] for item in dependency_components],
                }],
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

    def test_missing_cyclonedx_serial_number_fails(self) -> None:
        path = self.output / f"OfficeAgent.Test.{self.VERSION}.net8.0.cdx.json"
        bom = release_evidence.read_json(path)
        del bom["serialNumber"]
        path.write_text(json.dumps(bom), encoding="utf-8")
        with self.assertRaisesRegex(release_evidence.EvidenceError, "serial number"):
            release_evidence.validate_sbom(
                path,
                "OfficeAgent.Test",
                self.VERSION,
                "1.6",
                "net8.0",
                {"OfficeAgent.Dependency": self.VERSION},
            )

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

    def test_missing_declared_first_party_component_fails(self) -> None:
        path = self.output / f"OfficeAgent.Test.{self.VERSION}.net8.0.cdx.json"
        bom = release_evidence.read_json(path)
        bom["components"] = []
        path.write_text(json.dumps(bom), encoding="utf-8")
        with self.assertRaisesRegex(release_evidence.EvidenceError, "omits declared dependency"):
            release_evidence.validate_sbom(
                path,
                "OfficeAgent.Test",
                self.VERSION,
                "1.6",
                "net8.0",
                {"OfficeAgent.Dependency": self.VERSION},
            )

    def test_missing_declared_first_party_root_edge_fails(self) -> None:
        path = self.output / f"OfficeAgent.Test.{self.VERSION}.net8.0.cdx.json"
        bom = release_evidence.read_json(path)
        bom["dependencies"][0]["dependsOn"] = []
        path.write_text(json.dumps(bom), encoding="utf-8")
        with self.assertRaisesRegex(release_evidence.EvidenceError, "omits the root edge"):
            release_evidence.validate_sbom(
                path,
                "OfficeAgent.Test",
                self.VERSION,
                "1.6",
                "net8.0",
                {"OfficeAgent.Dependency": self.VERSION},
            )

    def test_sample_has_its_own_product_and_sbom(self) -> None:
        manifest = release_evidence.read_json(self.output / "release-manifest.json")
        artifacts = {item["name"]: item for item in manifest["artifacts"]}
        self.assertEqual(
            artifacts["test-sample.zip"]["dependencyInventory"],
            [f"test-sample.{self.VERSION}.cdx.json"],
        )

    def test_standalone_archive_and_bundle_share_the_runtime_sbom(self) -> None:
        manifest = release_evidence.read_json(self.output / "release-manifest.json")
        artifacts = {item["name"]: item for item in manifest["artifacts"]}
        expected = [f"test-server-win-x64.{self.VERSION}.cdx.json"]
        self.assertEqual(artifacts["test-server-win-x64.zip"]["kind"], "standalone-executable")
        self.assertEqual(artifacts["test-server-win-x64.zip"]["dependencyInventory"], expected)
        self.assertEqual(artifacts["test-server-win-x64.mcpb"]["kind"], "mcpb-bundle")
        self.assertEqual(artifacts["test-server-win-x64.mcpb"]["dependencyInventory"], expected)
        self.assertIn("test-server-linux-x64.tar.gz", artifacts)
        self.assertNotIn("test-server-linux-x64.mcpb", artifacts)

    def test_standalone_sbom_is_read_from_the_executable(self) -> None:
        bom = release_evidence.read_json(self.output / f"test-server-linux-x64.{self.VERSION}.cdx.json")
        components = {item["name"]: item["version"] for item in bom["components"]}
        self.assertEqual(components, {
            "Microsoft.NETCore.App.Runtime.linux-x64": "10.0.5",
            "ModelContextProtocol": "1.4.0",
            "OfficeAgent.Dependency": self.VERSION,
        })
        properties = bom["metadata"]["component"]["properties"]
        self.assertIn({"name": "officeagent:runtimeIdentifier", "value": "linux-x64"}, properties)

    def test_bundle_with_a_different_executable_fails(self) -> None:
        standalone = self.inventory["standalone"]
        write_standalone(
            self.output / "test-server-win-x64.mcpb",
            release_evidence.standalone_bundle_member(standalone, "win-x64"),
            fake_single_file("win-x64", self.VERSION, "net8.0") + b"other",
        )
        with self.assertRaisesRegex(release_evidence.EvidenceError, "different executable"):
            release_evidence.verify_standalone_payloads(self.inventory, self.output, self.VERSION)

    def test_executable_that_differs_from_its_sbom_fails(self) -> None:
        standalone = self.inventory["standalone"]
        write_standalone(
            self.output / "test-server-linux-x64.tar.gz",
            release_evidence.standalone_archive_member(standalone, "linux-x64"),
            fake_single_file("linux-x64", self.VERSION, "net8.0") + b"rebuilt",
        )
        with self.assertRaisesRegex(release_evidence.EvidenceError, "does not match its SBOM"):
            release_evidence.verify_standalone_payloads(self.inventory, self.output, self.VERSION)

    def test_unix_executable_without_execute_bit_fails(self) -> None:
        standalone = self.inventory["standalone"]
        write_standalone(
            self.output / "test-server-linux-x64.tar.gz",
            release_evidence.standalone_archive_member(standalone, "linux-x64"),
            fake_single_file("linux-x64", self.VERSION, "net8.0"),
            mode=0o644,
        )
        with self.assertRaisesRegex(release_evidence.EvidenceError, "not executable"):
            release_evidence.verify_standalone_payloads(self.inventory, self.output, self.VERSION)

    def test_executable_for_another_runtime_fails(self) -> None:
        with self.assertRaisesRegex(release_evidence.EvidenceError, "not built for"):
            release_evidence.single_file_dependencies(
                fake_single_file("osx-arm64", self.VERSION, "net8.0"),
                self.inventory["standalone"],
                "linux-x64",
            )

    def test_framework_dependent_executable_fails(self) -> None:
        with self.assertRaisesRegex(release_evidence.EvidenceError, "not a .NET single-file"):
            release_evidence.single_file_deps(b"MZ" + bytes(64), "apphost")

    def test_standalone_inventory_rejects_a_linux_bundle(self) -> None:
        self.inventory["standalone"]["bundles"] = ["linux-x64"]
        with self.assertRaisesRegex(release_evidence.EvidenceError, "Windows or macOS"):
            release_evidence.validate_standalone_inventory(
                self.inventory["standalone"], self.inventory["packages"])

    def test_standalone_inventory_rejects_an_unbuilt_framework(self) -> None:
        self.inventory["standalone"]["framework"] = "net10.0"
        with self.assertRaisesRegex(release_evidence.EvidenceError, "framework of a packaged project"):
            release_evidence.validate_standalone_inventory(
                self.inventory["standalone"], self.inventory["packages"])



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

    def test_standalone_server_ships_six_runtimes_and_desktop_bundles(self) -> None:
        inventory = release_evidence.load_inventory(release_evidence.DEFAULT_INVENTORY)
        standalone = inventory["standalone"]
        self.assertEqual(standalone["project"], "src/OfficeAgent.Mcp/OfficeAgent.Mcp.csproj")
        self.assertEqual(standalone["framework"], "net10.0")
        self.assertEqual(
            set(standalone["runtimes"]),
            {f"{system}-{arch}" for system in ("win", "osx", "linux") for arch in ("x64", "arm64")},
        )
        self.assertEqual(
            set(standalone["bundles"]),
            {rid for rid in standalone["runtimes"] if not rid.startswith("linux-")},
        )


if __name__ == "__main__":
    unittest.main()
