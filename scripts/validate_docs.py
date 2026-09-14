#!/usr/bin/env python3
"""Validate local documentation links, fences, and release metadata."""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path
from urllib.parse import unquote


ROOT = Path(__file__).resolve().parents[1]
MARKDOWN = [
    ROOT / "README.md",
    ROOT / "CONTRIBUTING.md",
    ROOT / "CHANGELOG.md",
    ROOT / "SECURITY.md",
    ROOT / "SUPPORT.md",
    *sorted((ROOT / "docs").glob("*.md")),
    *sorted((ROOT / "security").glob("*.md")),
    *sorted((ROOT / "samples").glob("**/README.md")),
    *sorted((ROOT / "skills").glob("**/*.md")),
]


def github_anchor(title: str) -> str:
    title = re.sub(r"<[^>]+>", "", title.strip().lower())
    title = re.sub(r"[^\w\- ]", "", title, flags=re.UNICODE)
    return title.replace(" ", "-")


def anchors(path: Path) -> set[str]:
    found: set[str] = set()
    duplicates: dict[str, int] = {}
    fenced = False
    for line in path.read_text(encoding="utf-8").splitlines():
        if line.lstrip().startswith("```"):
            fenced = not fenced
            continue
        if fenced:
            continue
        match = re.match(r"^#{1,6}\s+(.+?)\s*#*$", line)
        if not match:
            continue
        base = github_anchor(match.group(1))
        count = duplicates.get(base, 0)
        duplicates[base] = count + 1
        found.add(base if count == 0 else f"{base}-{count}")
    return found


def validate_markdown() -> list[str]:
    errors: list[str] = []
    anchor_cache: dict[Path, set[str]] = {}
    link_pattern = re.compile(r"!?\[[^\]]*\]\(([^\s)]+)(?:\s+[^)]*)?\)")

    for path in MARKDOWN:
        if not path.exists():
            errors.append(f"missing documentation file: {path.relative_to(ROOT)}")
            continue
        text = path.read_text(encoding="utf-8")
        if sum(1 for line in text.splitlines() if line.lstrip().startswith("```")) % 2:
            errors.append(f"unclosed code fence: {path.relative_to(ROOT)}")

        for line_number, line in enumerate(text.splitlines(), 1):
            for match in link_pattern.finditer(line):
                href = unquote(match.group(1).strip("<>"))
                if re.match(r"^[a-z][a-z0-9+.-]*:", href, re.IGNORECASE):
                    continue
                relative, _, fragment = href.partition("#")
                target = (path.parent / relative).resolve() if relative else path.resolve()
                try:
                    target.relative_to(ROOT)
                except ValueError:
                    errors.append(
                        f"link escapes repository: {path.relative_to(ROOT)}:{line_number}: {href}"
                    )
                    continue
                if not target.exists():
                    errors.append(
                        f"missing link target: {path.relative_to(ROOT)}:{line_number}: {href}"
                    )
                    continue
                if fragment and target.suffix.lower() == ".md":
                    available = anchor_cache.setdefault(target, anchors(target))
                    if fragment not in available:
                        errors.append(
                            f"missing anchor: {path.relative_to(ROOT)}:{line_number}: {href}"
                        )
    return errors


def validate_release_metadata() -> list[str]:
    errors: list[str] = []
    props = (ROOT / "Directory.Build.props").read_text(encoding="utf-8")
    match = re.search(r"<Version>([^<]+)</Version>", props)
    if not match:
        return ["Directory.Build.props has no Version"]
    version = match.group(1)

    manifest = json.loads((ROOT / "server.json").read_text(encoding="utf-8"))
    package = manifest.get("packages", [{}])[0]
    if manifest.get("version") != version:
        errors.append(f"server.json version {manifest.get('version')} != {version}")
    if package.get("version") != version:
        errors.append(f"server.json package version {package.get('version')} != {version}")
    if package.get("identifier") != "OfficeAgent.Mcp":
        errors.append("server.json package identifier is not OfficeAgent.Mcp")
    if len(manifest.get("description", "")) > 100:
        errors.append("server.json description exceeds the registry's 100-character limit")

    changelog = (ROOT / "CHANGELOG.md").read_text(encoding="utf-8")
    headings = re.finditer(
        r"^##\s+([0-9]+\.[0-9]+\.[0-9]+)\s+—\s+(.+)$", changelog, re.MULTILINE
    )
    first_released = next(
        (match for match in headings if match.group(2).strip().lower() != "unreleased"),
        None,
    )
    if not first_released or first_released.group(1) != version:
        errors.append(f"first released CHANGELOG version does not match {version}")
    return errors


def validate_skills() -> list[str]:
    errors: list[str] = []
    paths = sorted((ROOT / "skills").glob("*/SKILL.md"))
    if not paths:
        return ["repository has no skill definitions"]
    for path in paths:
        name = path.parent.name
        text = path.read_text(encoding="utf-8")
        match = re.match(r"^---\n(.*?)\n---\n", text, re.DOTALL)
        if not match:
            errors.append(f"{name} SKILL.md has no YAML frontmatter")
            continue
        fields: dict[str, str] = {}
        for line in match.group(1).splitlines():
            key, separator, value = line.partition(":")
            if separator:
                fields[key.strip()] = value.strip()
        if fields.get("name") != name:
            errors.append(f"{name} skill name does not match its directory")
        description = fields.get("description", "")
        if not description or len(description) > 1024:
            errors.append(f"{name} skill description is empty or exceeds 1024 characters")
    return errors


def main() -> int:
    errors = validate_markdown() + validate_release_metadata() + validate_skills()
    if errors:
        for error in errors:
            print(f"ERROR: {error}")
        return 1
    print(f"Documentation validation passed for {len(MARKDOWN)} Markdown files.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
