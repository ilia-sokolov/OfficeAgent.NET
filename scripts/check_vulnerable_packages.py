#!/usr/bin/env python3
"""Fail when NuGet reports a vulnerability without a current documented exception."""

from __future__ import annotations

import argparse
from datetime import date
import json
from pathlib import Path
import subprocess
import sys


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_ALLOWLIST = ROOT / "security" / "dependency-audit-allowlist.json"


def load_allowlist(path: Path) -> dict[tuple[str, str], dict[str, str]]:
    if not path.exists():
        return {}
    value = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(value, list):
        raise ValueError(f"{path} must contain a JSON array")
    allowed: dict[tuple[str, str], dict[str, str]] = {}
    for entry in value:
        required = {"package", "advisory", "expires", "reason"}
        if not isinstance(entry, dict) or not required.issubset(entry):
            raise ValueError(f"every entry in {path} needs package, advisory, expires, and reason")
        expiry = date.fromisoformat(str(entry["expires"]))
        if expiry < date.today():
            raise ValueError(f"expired dependency exception: {entry['package']} {entry['advisory']}")
        key = (str(entry["package"]).lower(), str(entry["advisory"]))
        allowed[key] = {name: str(entry[name]) for name in required}
    return allowed


def findings(report: dict[str, object]) -> list[dict[str, str]]:
    found: list[dict[str, str]] = []
    for project in report.get("projects", []):
        if not isinstance(project, dict):
            continue
        project_path = str(project.get("path", "<unknown project>"))
        for framework in project.get("frameworks", []):
            if not isinstance(framework, dict):
                continue
            for group in ("topLevelPackages", "transitivePackages"):
                for package in framework.get(group, []):
                    if not isinstance(package, dict):
                        continue
                    for vulnerability in package.get("vulnerabilities", []):
                        if not isinstance(vulnerability, dict):
                            continue
                        found.append(
                            {
                                "project": project_path,
                                "package": str(package.get("id", "<unknown package>")),
                                "version": str(package.get("resolvedVersion", "<unknown version>")),
                                "severity": str(vulnerability.get("severity", "unknown")),
                                "advisory": str(vulnerability.get("advisoryurl", "<unknown advisory>")),
                            }
                        )
    return found


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--allowlist", type=Path, default=DEFAULT_ALLOWLIST)
    args = parser.parse_args()

    try:
        allowed = load_allowlist(args.allowlist)
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        print(f"ERROR: invalid dependency exception file: {exc}")
        return 1

    command = [
        "dotnet",
        "list",
        str(ROOT / "OfficeAgent.NET.sln"),
        "package",
        "--vulnerable",
        "--include-transitive",
        "--format",
        "json",
    ]
    completed = subprocess.run(command, cwd=ROOT, text=True, capture_output=True, check=False)
    if completed.returncode != 0:
        print(completed.stdout, end="")
        print(completed.stderr, end="", file=sys.stderr)
        print("ERROR: NuGet vulnerability audit did not complete")
        return 1

    try:
        report = json.loads(completed.stdout)
    except json.JSONDecodeError as exc:
        print(f"ERROR: NuGet vulnerability audit returned invalid JSON: {exc}")
        return 1

    unapproved: list[dict[str, str]] = []
    used_exceptions: set[tuple[str, str]] = set()
    for finding in findings(report):
        key = (finding["package"].lower(), finding["advisory"])
        if key in allowed:
            used_exceptions.add(key)
            print(
                f"EXCEPTION: {finding['package']} {finding['version']} "
                f"{finding['advisory']} until {allowed[key]['expires']}"
            )
        else:
            unapproved.append(finding)

    stale = set(allowed) - used_exceptions
    for key in sorted(stale):
        print(f"ERROR: stale dependency exception no longer matches a finding: {key[0]} {key[1]}")
    for finding in unapproved:
        print(
            "ERROR: vulnerable dependency: "
            f"{finding['package']} {finding['version']} ({finding['severity']}) "
            f"{finding['advisory']} in {finding['project']}"
        )

    if unapproved or stale:
        return 1
    print("NuGet vulnerability audit passed with no unapproved findings.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
