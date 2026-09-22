"""
Keeps the published support, security and runtime promises consistent with each other and with
the repository they describe.

The policy lives in several files: SECURITY.md, SUPPORT.md, CONTRIBUTING.md, the compatibility
and maintenance guides, the issue templates, the CI workflow and the container image. Before 1.0
they had drifted apart: SECURITY.md still called the project pre-1.0, named 0.8.x as the
supported line after 0.9.0 shipped, and said no SBOM was published while SUPPORT.md described the
SBOMs every release carries. Each check below fails on one such disagreement.
"""

from __future__ import annotations

import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

# Microsoft's support policy, https://dotnet.microsoft.com/platform/support/policy/dotnet-core,
# fetched 2026-09-22. Every document that quotes an end-of-support date must quote these.
END_OF_SUPPORT = {"8": "2026-11-10", "10": "2028-11-14"}

POLICY_DOCUMENTS = (
    "SECURITY.md",
    "SUPPORT.md",
    "CONTRIBUTING.md",
    "docs/compatibility.md",
    "docs/maintenance.md",
    "docs/releasing.md",
)


def read(relative: str) -> str:
    return (ROOT / relative).read_text(encoding="utf-8")


def repository_version() -> tuple[int, int]:
    match = re.search(r"<Version>(\d+)\.(\d+)\.\d+</Version>", read("Directory.Build.props"))
    assert match, "Directory.Build.props has no <Version>"
    return int(match.group(1)), int(match.group(2))


def packable_projects() -> set[str]:
    names = set()
    for project in (ROOT / "src").glob("*/*.csproj"):
        text = project.read_text(encoding="utf-8")
        if "<IsPackable>false</IsPackable>" not in text:
            names.add(project.stem)
    return names


class SupportPolicyTests(unittest.TestCase):
    def test_security_table_supports_the_version_this_repository_builds(self) -> None:
        major, minor = repository_version()
        rows = re.findall(r"^\| ([^|]+?) \| ([^|]+?) \|\s*$", read("SECURITY.md"), re.MULTILINE)
        support = {version.strip(): status.strip() for version, status in rows}
        line = f"{major}.{minor}.x"
        self.assertIn(line, support, f"SECURITY.md has no row for {line}")
        self.assertTrue(support[line].startswith("Supported"), f"{line} is '{support[line]}'")
        unsupported_before = f"{major}.{minor - 1}.x and earlier" if minor > 0 else None
        if unsupported_before:
            self.assertEqual("Unsupported", support.get(unsupported_before))

    def test_no_policy_document_still_describes_the_project_as_pre_release(self) -> None:
        for relative in ("SECURITY.md", "SUPPORT.md", "CONTRIBUTING.md"):
            with self.subTest(document=relative):
                self.assertNotRegex(read(relative), r"(?i)\bpre-1\.0\b")

    def test_security_release_statement_matches_the_support_statement(self) -> None:
        security = read("SECURITY.md")
        self.assertNotIn("does not currently publish an SBOM", security)
        self.assertIn("CycloneDX", security)
        self.assertIn("CycloneDX", read("SUPPORT.md"))
        self.assertIn("macOS", security)

    def test_every_quoted_end_of_support_date_matches_the_microsoft_policy(self) -> None:
        pattern = re.compile(r"\.NET (\d+)[^.\n]{0,60}?(?:support|supported)[^.\n]{0,40}?(\d{4}-\d{2}-\d{2})")
        quoted = 0
        for relative in POLICY_DOCUMENTS:
            text = read(relative).replace("\r\n", " ").replace("\n", " ")
            for runtime, date in pattern.findall(text):
                if runtime in END_OF_SUPPORT:
                    quoted += 1
                    with self.subTest(document=relative, runtime=runtime):
                        self.assertEqual(END_OF_SUPPORT[runtime], date)
        self.assertGreater(quoted, 0, "no document quotes an end-of-support date")

    def test_support_compatibility_and_maintenance_name_the_same_runtimes(self) -> None:
        for relative in ("SUPPORT.md", "docs/compatibility.md", "docs/maintenance.md"):
            with self.subTest(document=relative):
                text = read(relative)
                self.assertIn(".NET 10", text)
                self.assertIn(END_OF_SUPPORT["8"], text)

    def test_the_dotnet_10_leg_asserts_the_runtime_it_claims(self) -> None:
        workflow = read(".github/workflows/build.yml")
        job = workflow[workflow.index("  runtime-dotnet10:"):workflow.index("  renderer-sandbox:")]
        self.assertIn("10.0.x", job)
        self.assertIn("DOTNET_ROLL_FORWARD: LatestMajor", job)
        self.assertIn('OFFICEAGENT_EXPECT_RUNTIME_MAJOR: "10"', job)
        self.assertIn("dotnet test OfficeAgent.NET.sln", job)
        self.assertIn("smoke_packaged_artifacts.py", job)

    def test_the_container_runs_on_the_runtime_the_documents_state(self) -> None:
        runtime_image = re.search(r"^FROM mcr\.microsoft\.com/dotnet/aspnet:(\d+)\.\d+ AS runtime$",
                                  read("Dockerfile"), re.MULTILINE)
        self.assertIsNotNone(runtime_image)
        self.assertEqual("10", runtime_image.group(1))
        self.assertIn("container image runs on the .NET 10 runtime", " ".join(read("SUPPORT.md").split()))
        mcp = read("src/OfficeAgent.Mcp/OfficeAgent.Mcp.csproj")
        self.assertIn("<RollForward>LatestMajor</RollForward>", mcp)

    def test_bug_template_lists_every_published_package_and_the_current_version(self) -> None:
        template = read(".github/ISSUE_TEMPLATE/bug.yml")
        listed = set(re.findall(r"^\s+- (OfficeAgent\.[A-Za-z]+)\s*$", template, re.MULTILINE))
        self.assertEqual(packable_projects(), listed)
        major, minor = repository_version()
        placeholder = re.search(r'placeholder: "(\d+)\.(\d+)\.\d+"', template)
        self.assertIsNotNone(placeholder)
        self.assertEqual((major, minor), (int(placeholder.group(1)), int(placeholder.group(2))))

    def test_security_reports_are_routed_away_from_public_issues(self) -> None:
        config = read(".github/ISSUE_TEMPLATE/config.yml")
        self.assertIn("/security/advisories/new", config)

    def test_commercial_service_is_separated_from_library_guarantees(self) -> None:
        for relative in ("README.md", "SUPPORT.md"):
            with self.subTest(document=relative):
                text = read(relative).replace("\r\n", " ").replace("\n", " ")
                self.assertIn("dotaction", text)
                self.assertRegex(text, r"(?i)own agreement")


if __name__ == "__main__":
    unittest.main()
