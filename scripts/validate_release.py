#!/usr/bin/env python3
"""Fail a release when its tag, changelog, package, and registry metadata drift."""

from __future__ import annotations

import re
import sys

import validate_docs


def main() -> int:
    if len(sys.argv) != 2 or not re.fullmatch(r"v\d+\.\d+\.\d+", sys.argv[1]):
        print("usage: validate_release.py v<major>.<minor>.<patch>")
        return 2

    version = sys.argv[1][1:]
    errors = (
        validate_docs.validate_markdown()
        + validate_docs.validate_release_metadata()
        + validate_docs.validate_skill()
    )

    props = (validate_docs.ROOT / "Directory.Build.props").read_text(encoding="utf-8")
    if f"<Version>{version}</Version>" not in props:
        errors.append(f"release tag {sys.argv[1]} does not match Directory.Build.props")

    changelog = (validate_docs.ROOT / "CHANGELOG.md").read_text(encoding="utf-8")
    heading = re.search(rf"^##\s+{re.escape(version)}\s+—\s+(.+)$", changelog, re.MULTILINE)
    if not heading:
        errors.append(f"CHANGELOG has no dated {version} heading")
    elif heading.group(1).strip().lower() == "unreleased":
        errors.append(f"CHANGELOG {version} is still marked unreleased")

    if errors:
        for error in errors:
            print(f"ERROR: {error}")
        return 1
    print(f"Release metadata is synchronized for {sys.argv[1]}.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
