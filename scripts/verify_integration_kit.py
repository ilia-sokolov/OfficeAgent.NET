#!/usr/bin/env python3
"""Verify the integration skill in a disposable installed-package consumer."""

from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path, PurePosixPath


ROOT = Path(__file__).resolve().parents[1]
LINK_RE = re.compile(r"!?\[[^\]]*\]\(([^\s)]+)(?:\s+[^)]*)?\)")
EXPECTED_OUTPUT = {
    "tracked-edit=passed",
    "refusal=unsupported-operation bytes-unchanged=True",
    "template-population=passed",
    "complete-comparison=passed",
    "sdk-interop=passed",
    "sdk-conflict=version-conflict document-unchanged=True",
    "all-recipes=passed",
}


class VerificationError(ValueError):
    pass


def repository_version() -> str:
    text = (ROOT / "Directory.Build.props").read_text(encoding="utf-8")
    match = re.search(r"<Version>([^<]+)</Version>", text)
    if not match:
        raise VerificationError("Directory.Build.props has no Version")
    return match.group(1)


def run(command: list[str], cwd: Path) -> subprocess.CompletedProcess[str]:
    result = subprocess.run(command, cwd=cwd, text=True, capture_output=True, check=False)
    if result.returncode:
        raise VerificationError(
            f"command failed ({result.returncode}): {' '.join(command)}\n{result.stdout}\n{result.stderr}"
        )
    return result


def validate_installed_references(skill_root: Path) -> None:
    markdown = sorted(skill_root.rglob("*.md"))
    if not markdown or not (skill_root / "SKILL.md").is_file():
        raise VerificationError("installed skill has no SKILL.md")
    for path in markdown:
        for line_number, line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
            for match in LINK_RE.finditer(line):
                href = match.group(1).strip("<>")
                if re.match(r"^[a-z][a-z0-9+.-]*:", href, re.IGNORECASE):
                    continue
                relative = href.partition("#")[0]
                target = (path.parent / relative).resolve() if relative else path.resolve()
                try:
                    target.relative_to(skill_root.resolve())
                except ValueError as exc:
                    raise VerificationError(f"installed reference escapes skill: {path.name}:{line_number}") from exc
                if not target.exists():
                    raise VerificationError(f"missing installed reference: {path.name}:{line_number}: {href}")


def write_nuget_config(path: Path, package_source: Path) -> None:
    path.write_text(
        """<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <packageSources>
    <clear />
    <add key="officeagent-local" value="{source}" />
    <add key="nuget.org" value="https://api.nuget.org/v3/index.json" protocolVersion="3" />
  </packageSources>
</configuration>
""".format(source=package_source.as_posix()),
        encoding="utf-8",
        newline="\n",
    )


def verify(artifacts: Path, version: str, dotnet: str) -> None:
    archive = artifacts / "officeagent-integration.zip"
    required_packages = [artifacts / f"OfficeAgent.{name}.{version}.nupkg" for name in ("Abstractions", "Core", "Word")]
    missing = [path.name for path in [archive, *required_packages] if not path.is_file()]
    if missing:
        raise VerificationError(f"missing verification artifacts: {', '.join(missing)}")

    with tempfile.TemporaryDirectory(prefix="officeagent-integration-kit-") as temporary:
        fixture = Path(temporary)
        install_home = fixture / "agent-home" / "skills"
        install_home.mkdir(parents=True)
        with zipfile.ZipFile(archive) as skill_archive:
            names = skill_archive.namelist()
            safe_names = all(
                not PurePosixPath(name).is_absolute()
                and ".." not in PurePosixPath(name).parts
                and PurePosixPath(name).parts[0] == "officeagent-integration"
                for name in names
            )
            if not names or len(names) != len(set(names)) or not safe_names:
                raise VerificationError("skill archive layout is invalid")
            skill_archive.extractall(install_home)
        installed = install_home / "officeagent-integration"
        validate_installed_references(installed)

        consumer = fixture / "consumer"
        consumer.mkdir()
        for name in ("IntegrationRecipes.csproj", "Program.cs"):
            shutil.copy2(installed / "assets" / name, consumer / name)
        write_nuget_config(consumer / "NuGet.Config", artifacts.resolve())
        package_property = f"-p:OfficeAgentPackageVersion={version}"
        run([dotnet, "restore", "--configfile", "NuGet.Config", package_property], consumer)
        execution = run(
            [dotnet, "run", "--configuration", "Release", "--no-restore", package_property],
            consumer,
        )
        observed = set(execution.stdout.splitlines())
        missing_output = sorted(EXPECTED_OUTPUT - observed)
        if missing_output:
            raise VerificationError(f"recipe output assertions were not observed: {missing_output}\n{execution.stdout}")

        assets = json.loads((consumer / "obj" / "project.assets.json").read_text(encoding="utf-8"))
        resolved = set(assets.get("libraries", {}))
        for package in ("OfficeAgent.Abstractions", "OfficeAgent.Core", "OfficeAgent.Word"):
            if f"{package}/{version}" not in resolved:
                raise VerificationError(f"consumer did not resolve {package} {version}")

        shutil.rmtree(installed)
        if installed.exists():
            raise VerificationError("skill removal did not remove the installed directory")
        print(f"install-layout={install_home}")
        print(f"package-version={version}")
        print("relative-references=passed")
        print(execution.stdout.strip())
        print("removal=passed")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--artifacts", type=Path, required=True)
    parser.add_argument("--version", default=None)
    parser.add_argument("--dotnet", default="dotnet")
    args = parser.parse_args()
    try:
        verify(args.artifacts.resolve(), args.version or repository_version(), args.dotnet)
    except (OSError, VerificationError, zipfile.BadZipFile) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    print("Integration kit verification completed successfully.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
