#!/usr/bin/env python3
"""Keep the Claude plugin in plugins/officeagent in step with the repository.

The plugin folder is what Anthropic's directory installs, so it carries its own copies of the
agent skills, the licence and the icon. This script copies them from their sources, which stay
the single place to edit them, and checks that the server version the plugin pins matches its
manifest. Run it after changing a skill:

    python scripts/sync_claude_plugin.py          # copy, then check
    python scripts/sync_claude_plugin.py --check  # check only; exit 1 on any drift
"""
from __future__ import annotations

import argparse
import json
import re
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PLUGIN = ROOT / "plugins" / "officeagent"
COPIES = [
    (ROOT / "LICENSE", PLUGIN / "LICENSE"),
    (ROOT / "icon.png", PLUGIN / "icon.png"),
]


def skill_files() -> list[tuple[Path, Path]]:
    pairs = []
    for source in sorted((ROOT / "skills").rglob("*")):
        if source.is_file():
            pairs.append((source, PLUGIN / "skills" / source.relative_to(ROOT / "skills")))
    return pairs


def pinned_version() -> str:
    servers = json.loads((PLUGIN / ".mcp.json").read_text(encoding="utf-8"))["mcpServers"]
    args = servers["officeagent"]["args"]
    match = re.fullmatch(r"officeagent-mcp==(\S+)", args[0])
    if not match:
        raise SystemExit(f"the plugin must pin officeagent-mcp to an exact version, not {args[0]!r}")
    return match.group(1)


def problems() -> list[str]:
    found = []
    for source, target in COPIES + skill_files():
        if not target.is_file() or target.read_bytes() != source.read_bytes():
            found.append(f"{target.relative_to(ROOT)} differs from {source.relative_to(ROOT)}")
    expected = {target for _, target in skill_files()}
    for extra in sorted((PLUGIN / "skills").rglob("*")):
        if extra.is_file() and extra not in expected:
            found.append(f"{extra.relative_to(ROOT)} has no source in skills/")
    manifest = json.loads((PLUGIN / ".claude-plugin" / "plugin.json").read_text(encoding="utf-8"))
    if manifest["version"] != pinned_version():
        found.append(f"plugin.json version {manifest['version']} != pinned officeagent-mcp {pinned_version()}")
    return found


def sync() -> None:
    target_skills = PLUGIN / "skills"
    if target_skills.exists():
        shutil.rmtree(target_skills)
    for source, target in COPIES + skill_files():
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(source, target)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--check", action="store_true", help="report drift without copying")
    args = parser.parse_args()
    if not args.check:
        sync()
    found = problems()
    for problem in found:
        print(f"ERROR: {problem}", file=sys.stderr)
    if not found:
        print(f"Claude plugin is in step with the repository (officeagent-mcp=={pinned_version()}).")
    return 1 if found else 0


if __name__ == "__main__":
    sys.exit(main())
