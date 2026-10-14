from __future__ import annotations

import json
import re
import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import sync_claude_plugin as plugin_sync  # noqa: E402

PLUGIN = ROOT / "plugins" / "officeagent"
IMAGES = {".png", ".jpg", ".jpeg", ".gif", ".webp"}


def manifest() -> dict:
    return json.loads((PLUGIN / ".claude-plugin" / "plugin.json").read_text(encoding="utf-8"))


def server() -> dict:
    return json.loads((PLUGIN / ".mcp.json").read_text(encoding="utf-8"))["mcpServers"]["officeagent"]


class PluginContentTests(unittest.TestCase):
    def test_the_plugin_carries_the_repository_skills_licence_and_icon_unchanged(self) -> None:
        self.assertEqual([], plugin_sync.problems(),
                         "run python scripts/sync_claude_plugin.py after changing a skill")

    def test_the_readme_says_what_it_runs_and_sends_in_enough_words(self) -> None:
        readme = (PLUGIN / "README.md").read_text(encoding="utf-8")
        prose = re.sub(r"```.*?```", "", readme, flags=re.DOTALL)
        self.assertGreaterEqual(len(prose.split()), 40)
        for section in ("## What it runs and what data it sends", "## Privacy Policy", "## Requirements"):
            self.assertIn(section, readme)
        self.assertIn(f"version {plugin_sync.pinned_version()}", readme)


class PluginDirectoryRuleTests(unittest.TestCase):
    """The automated checks Anthropic's directory runs, as its pre-submission checklist states them."""

    def test_the_name_is_a_distinctive_kebab_case_identifier(self) -> None:
        name = manifest()["name"]
        self.assertRegex(name, r"^[a-z0-9](?:[a-z0-9-]{0,62}[a-z0-9])?$")
        self.assertNotIn(name, {"claude", "anthropic", "official", "plugin", "mcp", "test"})
        for field in ("description", "author", "version", "license"):
            self.assertIn(field, manifest())

    def test_files_stay_within_what_the_directory_inspects(self) -> None:
        files = [path for path in PLUGIN.rglob("*") if path.is_file()]
        self.assertLessEqual(len(files), 512)
        for path in files:
            relative = path.relative_to(PLUGIN).as_posix()
            self.assertFalse(path.is_symlink(), relative)
            self.assertNotIn(path.name, {".DS_Store", "Thumbs.db", "desktop.ini"}, relative)
            self.assertRegex(relative, r"^[A-Za-z0-9._/-]+$")
            if path.suffix.lower() in IMAGES:
                continue
            self.assertLess(path.stat().st_size, 256 * 1024, relative)
            path.read_text(encoding="utf-8")  # text only: any other binary is held for review
        self.assertFalse((PLUGIN / "bin").exists(), "a top-level bin/ stops claude.ai and Cowork installing it")

    def test_the_server_is_a_launcher_pinned_to_the_manifest_version(self) -> None:
        entry = server()
        self.assertEqual("uvx", entry["command"])
        self.assertEqual(f"officeagent-mcp=={manifest()['version']}", entry["args"][0])
        self.assertEqual("--stdio", entry["args"][1])
        self.assertFalse(any(name in {".npmrc", "uv.toml", "pip.conf"} for name in
                             (path.name for path in PLUGIN.rglob("*"))), "no custom package source")

    def test_every_option_the_server_reads_is_declared_with_a_default(self) -> None:
        options = manifest().get("userConfig", {})
        used = set(re.findall(r"\$\{user_config\.([A-Za-z_][A-Za-z0-9_]*)\}", json.dumps(server())))
        self.assertEqual(set(options), used)
        for key, option in options.items():
            # Cowork skips a server whose options have no default, because it does not prompt.
            self.assertIn("default", option, key)
            self.assertIs(option["default"], False, f"{key} must stay off until the user turns it on")

    def test_environment_names_bind_to_server_options(self) -> None:
        source = (ROOT / "src" / "OfficeAgent.Mcp" / "OfficeAgentMcpOptions.cs").read_text(encoding="utf-8")
        for name in server()["env"]:
            for part in (part for part in name.split("__")[1:] if not part.isdigit()):
                self.assertRegex(source, rf"\bpublic [^\n]+ {part} \{{ get; set; \}}", name)

    def test_the_server_reaches_only_the_project_folder(self) -> None:
        env = server()["env"]
        self.assertEqual("${CLAUDE_PROJECT_DIR}", env["OfficeAgent__FileSystemConnections__0__RootPath"])
        self.assertEqual({".docx", ".pptx", ".xlsx"},
                         {value for name, value in env.items() if "AllowedExtensions" in name})


if __name__ == "__main__":
    unittest.main()
