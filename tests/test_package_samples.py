from __future__ import annotations

import sys
import tempfile
import unittest
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import package_samples  # noqa: E402


class PackageSamplesTests(unittest.TestCase):
    def test_quickedit_archive_is_deterministic_and_package_backed(self) -> None:
        with tempfile.TemporaryDirectory() as first, tempfile.TemporaryDirectory() as second:
            one = package_samples.package_quickedit(Path(first), "0.9.0")
            two = package_samples.package_quickedit(Path(second), "0.9.0")
            self.assertEqual(one.read_bytes(), two.read_bytes())

            with zipfile.ZipFile(one) as archive:
                self.assertEqual(
                    sorted(archive.namelist()),
                    [
                        "quickedit-sample/Program.cs",
                        "quickedit-sample/QuickEdit.csproj",
                        "quickedit-sample/README.md",
                        "quickedit-sample/services-agreement.docx",
                    ],
                )
                project = archive.read("quickedit-sample/QuickEdit.csproj").decode()
                program = archive.read("quickedit-sample/Program.cs").decode()

            self.assertNotIn("ProjectReference", project)
            self.assertIn("<ImplicitUsings>enable</ImplicitUsings>", project)
            self.assertIn('Include="OfficeAgent.Core" Version="0.9.0"', project)
            self.assertIn('Include="OfficeAgent.Word" Version="0.9.0"', project)
            self.assertIn('Author = "QuickEdit"', program)


if __name__ == "__main__":
    unittest.main()
