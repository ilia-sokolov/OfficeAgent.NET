from __future__ import annotations

import base64
import hashlib
import io
import sys
import tarfile
import tempfile
import unittest
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import package_wheels as wheels  # noqa: E402
import release_evidence  # noqa: E402

STANDALONE = release_evidence.load_inventory(release_evidence.DEFAULT_INVENTORY)["standalone"]
LICENSE = b"MIT License\n"


def fake_release(directory: Path, runtimes: list[str]) -> None:
    """Archives shaped like the release's, each with a small stand-in executable, plus SHA256SUMS."""
    sums = []
    for rid in runtimes:
        member = release_evidence.standalone_archive_member(STANDALONE, rid)
        folder = member.rsplit("/", 1)[0]
        executable = f"server for {rid}".encode()
        archive = directory / release_evidence.standalone_archive(STANDALONE, rid)
        if archive.name.endswith(".zip"):
            with zipfile.ZipFile(archive, "w") as target:
                target.writestr(member, executable)
                target.writestr(f"{folder}/LICENSE", LICENSE)
        else:
            with tarfile.open(archive, "w:gz") as target:
                for name, data, mode in ((member, executable, 0o755), (f"{folder}/LICENSE", LICENSE, 0o644)):
                    info = tarfile.TarInfo(name)
                    info.size, info.mode = len(data), mode
                    target.addfile(info, io.BytesIO(data))
        sums.append(f"{hashlib.sha256(archive.read_bytes()).hexdigest()}  {archive.name}")
    (directory / "SHA256SUMS").write_text("\n".join(sums) + "\n", encoding="utf-8")


class WheelBuildTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        self.root = Path(self.temporary.name)
        self.release = self.root / "release"
        self.release.mkdir()
        self.output = self.root / "dist"

    def tearDown(self) -> None:
        self.temporary.cleanup()

    def build(self, version: str = "1.2.0") -> list[Path]:
        fake_release(self.release, STANDALONE["runtimes"])
        return wheels.build(self.release, version, self.output)

    def test_one_wheel_per_released_runtime_with_its_platform_tag(self) -> None:
        built = sorted(path.name for path in self.build())
        expected = sorted(f"officeagent_mcp-1.2.0-py3-none-{wheels.PLATFORM_TAGS[rid]}.whl" for rid in STANDALONE["runtimes"])
        self.assertEqual(expected, built)
        self.assertEqual(set(STANDALONE["runtimes"]), set(wheels.PLATFORM_TAGS))

    def test_a_wheel_carries_the_release_executable_byte_for_byte_and_executable(self) -> None:
        wheel = next(path for path in self.build() if "manylinux_2_27_x86_64" in path.name)
        with zipfile.ZipFile(wheel) as archive:
            info = archive.getinfo("officeagent_mcp/bin/officeagent-mcp")
            self.assertEqual(b"server for linux-x64", archive.read(info))
            self.assertEqual(0o755, (info.external_attr >> 16) & 0o777)
            windows = next(path for path in self.output.iterdir() if path.name.endswith("win_amd64.whl"))
        with zipfile.ZipFile(windows) as archive:
            self.assertEqual(b"server for win-x64", archive.read("officeagent_mcp/bin/officeagent-mcp.exe"))

    def test_the_record_lists_every_file_with_its_hash_and_size(self) -> None:
        wheel = self.build()[0]
        with zipfile.ZipFile(wheel) as archive:
            names = archive.namelist()
            record_name = next(name for name in names if name.endswith(".dist-info/RECORD"))
            rows = [line.split(",") for line in archive.read(record_name).decode().splitlines()]
            self.assertEqual(sorted(names), sorted(row[0] for row in rows))
            for name, digest, size in rows:
                if name == record_name:
                    self.assertEqual(("", ""), (digest, size))
                    continue
                data = archive.read(name)
                expected = "sha256=" + base64.urlsafe_b64encode(hashlib.sha256(data).digest()).rstrip(b"=").decode()
                self.assertEqual((expected, str(len(data))), (digest, size), name)

    def test_metadata_entry_point_and_launcher_version(self) -> None:
        wheel = next(path for path in self.build("1.3.0-rc.2") if "win_amd64" in path.name)
        self.assertTrue(wheel.name.startswith("officeagent_mcp-1.3.0rc2-"))
        with zipfile.ZipFile(wheel) as archive:
            metadata = archive.read("officeagent_mcp-1.3.0rc2.dist-info/METADATA").decode()
            self.assertIn("Name: officeagent-mcp\n", metadata)
            self.assertIn("Version: 1.3.0rc2\n", metadata)
            self.assertIn("License-Expression: MIT\n", metadata)
            self.assertIn("/blob/v1.3.0-rc.2/CHANGELOG.md", metadata)
            self.assertIn("uvx", metadata)
            wheel_file = archive.read("officeagent_mcp-1.3.0rc2.dist-info/WHEEL").decode()
            self.assertIn("Root-Is-Purelib: false\n", wheel_file)
            self.assertIn("Tag: py3-none-win_amd64\n", wheel_file)
            self.assertEqual("[console_scripts]\nofficeagent-mcp = officeagent_mcp.__main__:main\n",
                             archive.read("officeagent_mcp-1.3.0rc2.dist-info/entry_points.txt").decode())
            self.assertEqual(LICENSE, archive.read("officeagent_mcp-1.3.0rc2.dist-info/licenses/LICENSE"))
            self.assertIn('__version__ = "1.3.0rc2"', archive.read("officeagent_mcp/__init__.py").decode())

    def test_builds_are_reproducible(self) -> None:
        first = {path.name: path.read_bytes() for path in self.build()}
        for path in self.output.iterdir():
            path.unlink()
        second = {path.name: path.read_bytes() for path in wheels.build(self.release, "1.2.0", self.output)}
        self.assertEqual(first, second)

    def test_an_archive_that_does_not_match_the_checksums_is_refused(self) -> None:
        fake_release(self.release, STANDALONE["runtimes"])
        archive = self.release / release_evidence.standalone_archive(STANDALONE, "linux-x64")
        archive.write_bytes(archive.read_bytes() + b"tampered")
        with self.assertRaisesRegex(wheels.PackagingError, "does not match SHA256SUMS"):
            wheels.build(self.release, "1.2.0", self.output)

    def test_without_checksums_nothing_is_built(self) -> None:
        fake_release(self.release, STANDALONE["runtimes"])
        (self.release / "SHA256SUMS").unlink()
        with self.assertRaisesRegex(wheels.PackagingError, "SHA256SUMS is missing"):
            wheels.build(self.release, "1.2.0", self.output)
        self.assertFalse(self.output.exists())

    def test_only_release_and_candidate_versions_are_accepted(self) -> None:
        self.assertEqual("1.2.0", wheels.python_version("1.2.0"))
        self.assertEqual("1.2.0rc4", wheels.python_version("1.2.0-rc.4"))
        for bad in ("1.2", "v1.2.0", "1.2.0-beta.1", "1.2.0-rc4"):
            with self.assertRaises(wheels.PackagingError, msg=bad):
                wheels.python_version(bad)


class LauncherTests(unittest.TestCase):
    def test_the_launcher_finds_its_executable_beside_it(self) -> None:
        sys.path.insert(0, str(ROOT / "packaging" / "pypi"))
        try:
            import officeagent_mcp
            path = officeagent_mcp.executable()
            self.assertEqual(ROOT / "packaging" / "pypi" / "officeagent_mcp" / "bin", path.parent)
            self.assertEqual("officeagent-mcp.exe" if sys.platform == "win32" else "officeagent-mcp", path.name)
            self.assertEqual("__VERSION__", officeagent_mcp.__version__)
        finally:
            sys.path.remove(str(ROOT / "packaging" / "pypi"))
            sys.modules.pop("officeagent_mcp", None)


if __name__ == "__main__":
    unittest.main()
