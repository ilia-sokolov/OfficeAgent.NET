from __future__ import annotations

import hashlib
import shutil
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import package_skills  # noqa: E402
import release_evidence  # noqa: E402
import verify_integration_kit  # noqa: E402


class PackageSkillsTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.output = Path(self.temporary.name) / "artifacts"
        self.inventory = release_evidence.load_inventory(release_evidence.DEFAULT_INVENTORY)

    def test_archives_are_deterministic_and_have_declared_layout(self) -> None:
        first = package_skills.package_skills(self.inventory, self.output)
        first_hashes = {path.name: hashlib.sha256(path.read_bytes()).hexdigest() for path in first}
        second = package_skills.package_skills(self.inventory, self.output)
        second_hashes = {path.name: hashlib.sha256(path.read_bytes()).hexdigest() for path in second}
        self.assertEqual(first_hashes, second_hashes)

        for skill in self.inventory["skills"]:
            with zipfile.ZipFile(self.output / skill["archive"]) as archive:
                names = archive.namelist()
            self.assertIn(f"{skill['name']}/SKILL.md", names)
            self.assertTrue(all(name.startswith(f"{skill['name']}/") for name in names))

    def test_integration_skill_installs_resolves_references_and_removes(self) -> None:
        package_skills.package_skills(self.inventory, self.output)
        install_home = Path(self.temporary.name) / "agent-home" / "skills"
        install_home.mkdir(parents=True)
        with zipfile.ZipFile(self.output / "officeagent-integration.zip") as archive:
            archive.extractall(install_home)
        installed = install_home / "officeagent-integration"
        verify_integration_kit.validate_installed_references(installed)
        self.assertTrue((installed / "assets" / "Program.cs").is_file())
        shutil.rmtree(installed)
        self.assertFalse(installed.exists())


if __name__ == "__main__":
    unittest.main()
