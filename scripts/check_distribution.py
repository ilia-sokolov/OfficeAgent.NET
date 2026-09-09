#!/usr/bin/env python3
"""Verify that a released OfficeAgent.Mcp version is aligned across public channels."""

from __future__ import annotations

import argparse
import json
import re
import sys
from urllib.parse import quote
from urllib.request import Request, urlopen


REPOSITORY = "ilia-sokolov/OfficeAgent.NET"
SERVER_NAME = "io.github.ilia-sokolov/officeagent"
PACKAGE_ID = "officeagent.mcp"
USER_AGENT = "OfficeAgent.NET-distribution-check"


def get_json(url: str, timeout: int) -> object:
    request = Request(url, headers={"Accept": "application/json", "User-Agent": USER_AGENT})
    with urlopen(request, timeout=timeout) as response:
        return json.load(response)


def github_release(timeout: int) -> dict[str, object]:
    value = get_json(f"https://api.github.com/repos/{REPOSITORY}/releases/latest", timeout)
    if not isinstance(value, dict):
        raise ValueError("GitHub returned an unexpected release response")
    return value


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--version",
        help="expected major.minor.patch version; defaults to the latest GitHub release",
    )
    parser.add_argument("--timeout", type=int, default=30, help="HTTP timeout in seconds")
    args = parser.parse_args()

    errors: list[str] = []
    try:
        release = github_release(args.timeout)
        latest_tag = str(release.get("tag_name", ""))
        expected = args.version or latest_tag.removeprefix("v")
        if not re.fullmatch(r"\d+\.\d+\.\d+(?:[-+][0-9A-Za-z.-]+)?", expected):
            print(f"ERROR: invalid expected version: {expected!r}")
            return 2
        if latest_tag != f"v{expected}":
            errors.append(f"latest GitHub release is {latest_tag or '<missing>'}, expected v{expected}")

        nuget = get_json(
            f"https://api.nuget.org/v3-flatcontainer/{PACKAGE_ID}/index.json", args.timeout
        )
        versions = nuget.get("versions", []) if isinstance(nuget, dict) else []
        if expected not in versions:
            errors.append(f"NuGet does not expose {PACKAGE_ID} {expected}")

        encoded_name = quote(SERVER_NAME, safe="")
        registry = get_json(
            "https://registry.modelcontextprotocol.io/v0.1/servers/"
            f"{encoded_name}/versions/latest",
            args.timeout,
        )
        server = registry.get("server", {}) if isinstance(registry, dict) else {}
        if server.get("version") != expected:
            errors.append(
                "MCP Registry latest version is "
                f"{server.get('version', '<missing>')}, expected {expected}"
            )
        packages = server.get("packages", []) if isinstance(server, dict) else []
        matching_packages = [
            package
            for package in packages
            if isinstance(package, dict)
            and package.get("identifier", "").lower() == PACKAGE_ID
        ]
        if not matching_packages:
            errors.append(f"MCP Registry does not advertise package {PACKAGE_ID}")
        elif matching_packages[0].get("version") != expected:
            errors.append(
                "MCP Registry package version is "
                f"{matching_packages[0].get('version', '<missing>')}, expected {expected}"
            )
        repository = server.get("repository", {}) if isinstance(server, dict) else {}
        if repository.get("url") != f"https://github.com/{REPOSITORY}":
            errors.append("MCP Registry repository URL does not match the canonical repository")
    except Exception as exc:  # Network and response failures must make this gate fail closed.
        print(f"ERROR: distribution check could not complete: {exc}")
        return 1

    if errors:
        for error in errors:
            print(f"ERROR: {error}")
        return 1

    print(
        f"Distribution is aligned for {expected}: GitHub release, NuGet package, "
        "and MCP Registry entry."
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
