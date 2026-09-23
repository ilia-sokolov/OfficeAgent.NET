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


def documented_release_version() -> str:
    changelog = (ROOT / "CHANGELOG.md").read_text(encoding="utf-8")
    match = re.search(r"^## (\d+\.\d+\.\d+) —", changelog, re.MULTILINE)
    if not match:
        raise release_evidence.EvidenceError("CHANGELOG.md has no release heading")
    return match.group(1)


def validate_skill_semantics(name: str, files: list[Path]) -> None:
    text_files = [path for path in files if path.suffix.lower() in {".md", ".csproj"}]
    combined = "\n".join(path.read_text(encoding="utf-8") for path in text_files)
    if name == "officeagent-integration":
        version = documented_release_version()
        if version not in combined:
            raise release_evidence.EvidenceError(
                f"{name} does not name documented release {version}"
            )
        stale_links = re.findall(r"blob/v(\d+\.\d+\.\d+)/", combined)
        if any(linked != version for linked in stale_links):
            raise release_evidence.EvidenceError(
                f"{name} contains a documentation link for a release other than {version}"
            )
        default = re.search(r"<OfficeAgentPackageVersion[^>]*>([^<]+)</OfficeAgentPackageVersion>", combined)
        if not default or default.group(1) != version:
            raise release_evidence.EvidenceError(
                f"{name} package default does not equal documented release {version}"
            )
    elif name == "word-document-review":
        required = ("writeOutcome", "possibleOutput", "do not retry blindly")
        missing = [phrase for phrase in required if phrase not in combined]
        if missing or "A plan that fails wrote nothing" in combined:
            raise release_evidence.EvidenceError(
                f"{name} does not carry the uncertain-write recovery contract"
            )


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
        validate_skill_semantics(name, files)

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
