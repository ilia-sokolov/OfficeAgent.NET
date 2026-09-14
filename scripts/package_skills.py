#!/usr/bin/env python3
"""Create deterministic skill archives declared by the release inventory."""

from __future__ import annotations

import argparse
import os
import re
import sys
import zipfile
from pathlib import Path

import release_evidence


ROOT = Path(__file__).resolve().parents[1]
SKILL_NAME_RE = re.compile(r"^[a-z0-9]+(?:-[a-z0-9]+)*$")
ARCHIVE_TIMESTAMP = (1980, 1, 1, 0, 0, 0)


def package_skills(inventory: dict, output: Path) -> list[Path]:
    output.mkdir(parents=True, exist_ok=True)
    archives: list[Path] = []
    for skill in inventory["skills"]:
        name = skill["name"]
        archive_name = skill["archive"]
        if not SKILL_NAME_RE.fullmatch(name):
            raise release_evidence.EvidenceError(f"invalid skill name: {name}")
        if archive_name != f"{name}.zip" or Path(archive_name).name != archive_name:
            raise release_evidence.EvidenceError(f"invalid skill archive name: {archive_name}")
        source = (ROOT / skill["source"]).resolve()
        expected = (ROOT / "skills" / name).resolve()
        if source != expected or not source.is_dir():
            raise release_evidence.EvidenceError(f"invalid or missing skill source: {skill['source']}")
        if not (source / "SKILL.md").is_file():
            raise release_evidence.EvidenceError(f"{name} has no SKILL.md")
        files = sorted(path for path in source.rglob("*") if path.is_file())
        if not files:
            raise release_evidence.EvidenceError(f"{name} has no package files")
        for path in files:
            if path.is_symlink():
                raise release_evidence.EvidenceError(f"skill packages cannot contain symlinks: {path}")

        destination = output / archive_name
        temporary = output / f".{archive_name}.tmp"
        try:
            with zipfile.ZipFile(temporary, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
                for path in files:
                    relative = path.relative_to(source).as_posix()
                    info = zipfile.ZipInfo(f"{name}/{relative}", ARCHIVE_TIMESTAMP)
                    info.compress_type = zipfile.ZIP_DEFLATED
                    info.external_attr = (0o100644 & 0xFFFF) << 16
                    archive.writestr(info, path.read_bytes(), compresslevel=9)
            os.replace(temporary, destination)
        finally:
            if temporary.exists():
                temporary.unlink()
        archives.append(destination)
    return archives


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--inventory", type=Path, default=release_evidence.DEFAULT_INVENTORY)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    try:
        inventory = release_evidence.load_inventory(args.inventory.resolve())
        archives = package_skills(inventory, args.output.resolve())
    except (OSError, release_evidence.EvidenceError, zipfile.BadZipFile) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    for archive in archives:
        print(f"Created {archive.name}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
