#!/usr/bin/env python3
"""Run the packaged QuickEdit sample outside the repository against local packages."""

from __future__ import annotations

import argparse
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def run(command: list[str], cwd: Path) -> None:
    result = subprocess.run(command, cwd=cwd, check=False)
    if result.returncode:
        raise RuntimeError(f"command failed ({result.returncode}): {' '.join(command)}")


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


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--artifacts", type=Path, required=True)
    parser.add_argument("--dotnet", default="dotnet")
    args = parser.parse_args()
    artifacts = args.artifacts.resolve()
    archive = artifacts / "quickedit-sample.zip"
    if not archive.is_file():
        print(f"ERROR: missing {archive}", file=sys.stderr)
        return 1

    try:
        with tempfile.TemporaryDirectory(prefix="officeagent-quickedit-") as temporary:
            root = Path(temporary)
            with zipfile.ZipFile(archive) as package:
                package.extractall(root)
            sample = root / "quickedit-sample"
            config = sample / "NuGet.Config"
            config.write_text(
                """<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <packageSources>
    <clear />
    <add key="officeagent-local" value="LOCAL_FEED" />
    <add key="nuget.org" value="https://api.nuget.org/v3/index.json" protocolVersion="3" />
  </packageSources>
</configuration>
""".replace("LOCAL_FEED", artifacts.as_posix()),
                encoding="utf-8",
            )
            run([args.dotnet, "restore", "QuickEdit.csproj", "--configfile", str(config)], sample)
            output = sample / "quickedit-output.docx"
            run(
                [
                    args.dotnet,
                    "run",
                    "--project",
                    "QuickEdit.csproj",
                    "--no-restore",
                    "--",
                    "services-agreement.docx",
                    str(output),
                ],
                sample,
            )
            verify_output(output)
            missing_output = sample / "must-not-exist.docx"
            missing = subprocess.run(
                [
                    args.dotnet,
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
                check=False,
            )
            if missing.returncode != 4 or missing_output.exists():
                raise RuntimeError(
                    "QuickEdit missing-source path did not fail closed without an output"
                )
    except (OSError, RuntimeError, zipfile.BadZipFile, ET.ParseError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    print("Packaged QuickEdit sample passed outside-repository verification.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
