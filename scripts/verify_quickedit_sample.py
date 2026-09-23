#!/usr/bin/env python3
"""Run the packaged QuickEdit sample outside the repository against local packages.

The sample is candidate evidence only when it ran on exactly the requested packages, so:
- the requested OfficeAgent packages must be in the artifacts folder before anything runs;
- the packaged sample project must reference exactly that version;
- restore uses a fresh package cache inside the temporary directory, the artifacts folder first and
  nuget.org only for third-party dependencies, so a stale package in the machine's cache cannot
  satisfy it;
- after restore, obj/project.assets.json must resolve OfficeAgent.Abstractions, Core and Word as
  packages at exactly that version, or the sample never runs.
"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

sys.path.insert(0, str(Path(__file__).resolve().parent))
from smoke_packaged_artifacts import isolated_env, repository_version, write_nuget_config  # noqa: E402


WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# The packages the sample references directly.
REFERENCED_PACKAGES = ("OfficeAgent.Core", "OfficeAgent.Word")
# The OfficeAgent packages that must resolve at the requested version. Abstractions arrives through Core.
RESOLVED_PACKAGES = ("OfficeAgent.Abstractions",) + REFERENCED_PACKAGES


def run(command: list[str], cwd: Path, env: dict[str, str]) -> None:
    result = subprocess.run(command, cwd=cwd, env=env, check=False)
    if result.returncode:
        raise RuntimeError(f"command failed ({result.returncode}): {' '.join(command)}")


def missing_packages(artifacts: Path, version: str) -> list[str]:
    """The requested OfficeAgent package files absent from the artifacts folder."""
    names = [f"{package}.{version}.nupkg" for package in RESOLVED_PACKAGES]
    return [name for name in names if not (artifacts / name).is_file()]


def project_version_problems(project: str, version: str) -> list[str]:
    """Every directly referenced OfficeAgent package the sample project does not pin to the version."""
    return [
        f"{package}: the sample project does not reference version {version!r}"
        for package in REFERENCED_PACKAGES
        if f'<PackageReference Include="{package}" Version="{version}" />' not in project
    ]


def resolved_version_problems(assets: dict, version: str) -> list[str]:
    """Every OfficeAgent package that restore did not resolve as a package at exactly the version.

    Package ids compare case-insensitively, as NuGet does; versions compare exactly. A missing
    package, another version, or a project reference in place of the package is a problem.
    """
    libraries = assets.get("libraries") or {}
    problems = []
    for package in RESOLVED_PACKAGES:
        resolved = sorted(
            key for key in libraries if key.split("/", 1)[0].lower() == package.lower()
        )
        expected = f"{package}/{version}"
        exact = [key for key in resolved if key.split("/", 1)[1] == version]
        if not exact or len(resolved) != 1:
            problems.append(f"{package}: resolved {resolved or 'nothing'}, expected {expected!r}")
        elif libraries[exact[0]].get("type") != "package":
            problems.append(f"{package}: resolved {exact[0]!r} as {libraries[exact[0]].get('type')!r}, not a package")
    return problems


def restore_environment(root: Path, artifacts: Path) -> tuple[dict[str, str], Path]:
    """An environment and NuGet.Config whose package cache is new and inside root.

    The artifacts folder is the first source; nuget.org is the second, for third-party dependencies.
    """
    cache = root / "nuget-packages"
    cache.mkdir()
    config = root / "NuGet.Config"
    write_nuget_config(config, artifacts, include_nuget_org=True)
    return isolated_env(cache, config), config


def verify_output(path: Path) -> None:
    with zipfile.ZipFile(path) as package:
        document = ET.fromstring(package.read("word/document.xml"))
        names = set(package.namelist())
    text = "".join(node.text or "" for node in document.iter(f"{{{WORD_NS}}}t"))
    deleted = "".join(node.text or "" for node in document.iter(f"{{{WORD_NS}}}delText"))
    authors = {
        node.attrib.get(f"{{{WORD_NS}}}author")
        for node in document.iter()
        if node.tag in {f"{{{WORD_NS}}}ins", f"{{{WORD_NS}}}del"}
    }
    if "within forty-five days of receipt" not in text:
        raise RuntimeError("QuickEdit output omits the replacement text")
    if "within thirty days of receipt" not in deleted:
        raise RuntimeError("QuickEdit output omits the tracked deletion")
    if "QuickEdit" not in authors:
        raise RuntimeError("QuickEdit output omits its explicit revision author")
    if "word/comments.xml" not in names:
        raise RuntimeError("QuickEdit output did not preserve the existing comment part")


def verify(artifacts: Path, version: str, dotnet: str) -> None:
    archive = artifacts / "quickedit-sample.zip"
    if not archive.is_file():
        raise RuntimeError(f"missing {archive}")
    missing = missing_packages(artifacts, version)
    if missing:
        raise RuntimeError(f"missing candidate packages in {artifacts}: {', '.join(missing)}")

    with tempfile.TemporaryDirectory(prefix="officeagent-quickedit-") as temporary:
        root = Path(temporary)
        with zipfile.ZipFile(archive) as package:
            package.extractall(root)
        sample = root / "quickedit-sample"
        problems = project_version_problems((sample / "QuickEdit.csproj").read_text(encoding="utf-8"), version)
        if problems:
            raise RuntimeError("; ".join(problems))

        env, config = restore_environment(root, artifacts)
        run([dotnet, "restore", "QuickEdit.csproj", "--configfile", str(config)], sample, env)
        assets = json.loads((sample / "obj" / "project.assets.json").read_text(encoding="utf-8"))
        problems = resolved_version_problems(assets, version)
        if problems:
            raise RuntimeError("restore did not resolve the candidate packages: " + "; ".join(problems))
        print(f"resolved-packages=passed expected={version} packages={len(RESOLVED_PACKAGES)}", flush=True)

        output = sample / "quickedit-output.docx"
        run(
            [
                dotnet,
                "run",
                "--project",
                "QuickEdit.csproj",
                "--no-restore",
                "--",
                "services-agreement.docx",
                str(output),
            ],
            sample,
            env,
        )
        verify_output(output)
        missing_output = sample / "must-not-exist.docx"
        missing_source = subprocess.run(
            [
                dotnet,
                "run",
                "--project",
                "QuickEdit.csproj",
                "--no-restore",
                "--",
                "services-agreement.docx",
                str(missing_output),
                "text that is absent from the fixture",
                "replacement",
            ],
            cwd=sample,
            env=env,
            check=False,
        )
        if missing_source.returncode != 4 or missing_output.exists():
            raise RuntimeError("QuickEdit missing-source path did not fail closed without an output")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--artifacts", type=Path, required=True)
    parser.add_argument(
        "--version",
        default=None,
        help="OfficeAgent package version the sample must restore; defaults to Directory.Build.props",
    )
    parser.add_argument("--dotnet", default="dotnet")
    args = parser.parse_args()
    try:
        version = args.version or repository_version()
        verify(args.artifacts.resolve(), version, args.dotnet)
    except (OSError, RuntimeError, ValueError, zipfile.BadZipFile, ET.ParseError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    print(f"Packaged QuickEdit sample passed outside-repository verification with OfficeAgent {version}.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
