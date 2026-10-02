from __future__ import annotations

import re
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


# Sample projects CI does not build, each with its reason. Keep this short: a sample outside
# the solution stops compiling unnoticed, as samples/ExcelRecalculation did when an API it
# called became internal in 1.0.
NOT_IN_SOLUTION = {
    "samples/RequirementsReview/RequirementsReview.csproj": "in progress; add it to the solution when it lands",
}


class SampleSolutionTests(unittest.TestCase):
    def test_every_sample_project_is_built_by_the_solution(self) -> None:
        solution = (ROOT / "OfficeAgent.NET.sln").read_text(encoding="utf-8-sig")
        listed = {
            path.replace("\\", "/")
            for path in re.findall(r'"(samples\\[^"]+?\.csproj)"', solution)
        }
        projects = {
            path.relative_to(ROOT).as_posix() for path in (ROOT / "samples").glob("*/*.csproj")
        }

        self.assertEqual(
            [],
            sorted(projects - listed - set(NOT_IN_SOLUTION)),
            "Add these sample projects to OfficeAgent.NET.sln so CI builds them, "
            "or list them in NOT_IN_SOLUTION with a reason",
        )
        self.assertEqual(
            [],
            sorted(path for path in NOT_IN_SOLUTION if path not in projects or path in listed),
            "Remove these entries from NOT_IN_SOLUTION: the project is gone or already built",
        )


if __name__ == "__main__":
    unittest.main()
