from __future__ import annotations

import json
import re
import sys
import tarfile
import tempfile
import unittest
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import package_standalone as package  # noqa: E402
import release_evidence  # noqa: E402


STANDALONE = release_evidence.load_inventory(release_evidence.DEFAULT_INVENTORY)["standalone"]


def mach_o(cputype: int, signed: bool) -> bytes:
    commands = [(0x19, 72)] + ([(package.LC_CODE_SIGNATURE, 16)] if signed else [])
    body = b"".join(
        command.to_bytes(4, "little") + size.to_bytes(4, "little") + bytes(size - 8)
        for command, size in commands
    )
    header = (
        b"\xcf\xfa\xed\xfe" + cputype.to_bytes(4, "little") + bytes(8)
        + len(commands).to_bytes(4, "little") + len(body).to_bytes(4, "little") + bytes(8)
    )
    return header + body


def portable_executable(machine: int) -> bytes:
    content = bytearray(b"MZ" + bytes(0x80))
    content[0x3C:0x40] = (0x40).to_bytes(4, "little")
    content[0x40:0x44] = b"PE\0\0"
    content[0x44:0x46] = machine.to_bytes(2, "little")
    return bytes(content)


def elf(machine: int) -> bytes:
    return b"\x7fELF" + bytes(14) + machine.to_bytes(2, "little") + bytes(44)


class ExecutableFormatTests(unittest.TestCase):
    def test_each_runtime_accepts_its_own_format_and_architecture(self) -> None:
        package.check_executable_format(portable_executable(0x8664), "win-x64")
        package.check_executable_format(portable_executable(0xAA64), "win-arm64")
        package.check_executable_format(elf(0x3E), "linux-x64")
        package.check_executable_format(elf(0xB7), "linux-arm64")
        package.check_executable_format(mach_o(0x01000007, signed=True), "osx-x64")
        package.check_executable_format(mach_o(0x0100000C, signed=True), "osx-arm64")

    def test_wrong_architecture_fails(self) -> None:
        with self.assertRaisesRegex(package.PackagingError, "not a linux-arm64 binary"):
            package.check_executable_format(elf(0x3E), "linux-arm64")
        with self.assertRaisesRegex(package.PackagingError, "not a win-x64 binary"):
            package.check_executable_format(elf(0x3E), "win-x64")

    def test_unsigned_macos_executable_fails(self) -> None:
        with self.assertRaisesRegex(package.PackagingError, "no code signature"):
            package.check_executable_format(mach_o(0x0100000C, signed=False), "osx-arm64")


class BundleManifestTests(unittest.TestCase):
    def manifest(self, rid: str) -> dict:
        return package.bundle_manifest(STANDALONE, rid, "1.2.0")

    def test_every_bundle_runs_its_own_executable(self) -> None:
        for rid in STANDALONE["bundles"]:
            manifest = self.manifest(rid)
            server = manifest["server"]
            self.assertEqual(server["type"], "binary")
            self.assertEqual(server["entry_point"], release_evidence.standalone_bundle_member(STANDALONE, rid))
            self.assertEqual(server["mcp_config"]["command"], "${__dirname}/" + server["entry_point"])
            self.assertEqual(server["mcp_config"]["args"], ["--stdio"])
            platform = "win32" if rid.startswith("win-") else "darwin"
            self.assertEqual(manifest["compatibility"]["platforms"], [platform])
            self.assertEqual(manifest["version"], "1.2.0")

    def test_every_setting_the_command_uses_is_declared(self) -> None:
        manifest = self.manifest("osx-arm64")
        text = json.dumps(manifest["server"]["mcp_config"])
        used = set(re.findall(r"\$\{user_config\.([^}]+)\}", text))
        self.assertEqual(used, set(manifest["user_config"]))

    def test_creation_stays_off_until_the_user_allows_it(self) -> None:
        settings = self.manifest("win-x64")["user_config"]
        self.assertIs(settings["allow_creation"]["default"], False)
        self.assertEqual(settings["documents_folder"]["type"], "directory")
        self.assertTrue(settings["documents_folder"]["required"])

    def test_environment_names_bind_to_server_options(self) -> None:
        options = (ROOT / "src" / "OfficeAgent.Mcp" / "OfficeAgentMcpOptions.cs").read_text(encoding="utf-8")
        env = self.manifest("win-x64")["server"]["mcp_config"]["env"]
        for name in env:
            parts = [part for part in name.split("__")[1:] if not part.isdigit()]
            for part in parts:
                self.assertRegex(options, rf"public [^\n]+ {part} \{{", f"{name} binds to no option")
        extensions = sorted(value for key, value in env.items() if "AllowedExtensions" in key)
        self.assertEqual(extensions, [".docx", ".pptx", ".xlsx"])

    def test_listed_tools_are_unique(self) -> None:
        names = [tool["name"] for tool in self.manifest("win-x64")["tools"]]
        self.assertEqual(len(names), len(set(names)))
        self.assertTrue(self.manifest("win-x64")["tools_generated"])


class ArchiveTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.output = Path(self.temporary.name)

    def package(self, rid: str, executable: bytes = b"server", bundle: bool = True) -> list[Path]:
        return package.package_runtime(STANDALONE, rid, "1.2.0", executable, self.output, bundle)

    def test_archives_are_reproducible(self) -> None:
        for rid in ("win-x64", "linux-arm64"):
            first = [path.read_bytes() for path in self.package(rid)]
            second = [path.read_bytes() for path in self.package(rid)]
            self.assertEqual(first, second, rid)

    def test_unix_archive_keeps_the_execute_bit(self) -> None:
        archive, = self.package("linux-x64", bundle=False)
        self.assertEqual(archive.name, "officeagent-mcp-linux-x64.tar.gz")
        with tarfile.open(archive) as contents:
            members = {member.name: member for member in contents.getmembers()}
        executable = members["officeagent-mcp-linux-x64/officeagent-mcp"]
        self.assertEqual(executable.mode, 0o755)
        self.assertEqual(members["officeagent-mcp-linux-x64/LICENSE"].mode, 0o644)
        self.assertIn("officeagent-mcp-linux-x64/README.txt", members)

    def test_bundle_carries_manifest_icon_license_and_executable(self) -> None:
        archive, bundle = self.package("osx-arm64", b"same bytes")
        self.assertEqual(bundle.name, "officeagent-mcp-osx-arm64.mcpb")
        with zipfile.ZipFile(bundle) as contents:
            self.assertEqual(
                sorted(contents.namelist()),
                ["LICENSE", "icon.png", "manifest.json", "server/officeagent-mcp"],
            )
            entry = contents.getinfo("server/officeagent-mcp")
            self.assertEqual(entry.external_attr >> 16, 0o100755)
            self.assertEqual(contents.read(entry), b"same bytes")
            manifest = json.loads(contents.read("manifest.json"))
        self.assertEqual(manifest["version"], "1.2.0")
        with tarfile.open(archive) as contents:
            self.assertEqual(
                contents.extractfile("officeagent-mcp-osx-arm64/officeagent-mcp").read(), b"same bytes"
            )

    def test_windows_readme_uses_windows_line_endings(self) -> None:
        archive, _ = self.package("win-arm64")
        with zipfile.ZipFile(archive) as contents:
            readme = contents.read("officeagent-mcp-win-arm64/README.txt")
        self.assertIn(b"officeagent-mcp.exe --stdio\r\n", readme)
        self.assertNotIn(b"\r\r", readme)


if __name__ == "__main__":
    unittest.main()
