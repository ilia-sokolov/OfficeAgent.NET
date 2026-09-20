#!/usr/bin/env python3
"""Build deterministic standalone sample archives for a release."""

from __future__ import annotations

import argparse
import re
import sys
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
VERSION_RE = re.compile(r"<Version>([^<]+)</Version>")
ZIP_TIMESTAMP = (1980, 1, 1, 0, 0, 0)


def current_version() -> str:
    match = VERSION_RE.search((ROOT / "Directory.Build.props").read_text(encoding="utf-8"))
    if not match:
        raise ValueError("Directory.Build.props has no Version")
    return match.group(1)


def add_bytes(archive: zipfile.ZipFile, name: str, content: bytes) -> None:
    info = zipfile.ZipInfo(name, ZIP_TIMESTAMP)
    info.compress_type = zipfile.ZIP_DEFLATED
    info.external_attr = 0o100644 << 16
    archive.writestr(info, content)


def package_quickedit(output: Path, version: str) -> Path:
    source = ROOT / "samples" / "QuickEdit"
    project = (source / "QuickEdit.Package.csproj.template").read_text(encoding="utf-8")
    project = project.replace("__OFFICEAGENT_VERSION__", version)
    if "__OFFICEAGENT_VERSION__" in project:
        raise ValueError("QuickEdit package project still contains a version placeholder")

    files = {
        "quickedit-sample/QuickEdit.csproj": project.encode(),
        "quickedit-sample/Program.cs": (source / "Program.cs").read_bytes(),
        "quickedit-sample/README.md": (source / "README.package.md").read_bytes(),
        "quickedit-sample/services-agreement.docx": (
            ROOT / "samples" / "documents" / "services-agreement.docx"
        ).read_bytes(),
    }
    output.mkdir(parents=True, exist_ok=True)
    destination = output / "quickedit-sample.zip"
    with zipfile.ZipFile(destination, "w") as archive:
        for name, content in sorted(files.items()):
            add_bytes(archive, name, content)
    return destination


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--version", default=current_version())
    args = parser.parse_args()
    try:
        archive = package_quickedit(args.output.resolve(), args.version)
    except (OSError, ValueError, zipfile.BadZipFile) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    print(f"Packaged {archive}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
