from __future__ import annotations

import base64
import io
import json
import subprocess
import sys
import tempfile
import textwrap
import unittest
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import smoke_packaged_artifacts as smoke  # noqa: E402


def fake_server(script: str) -> list[str]:
    """A stand-in stdio peer, so the harness boundaries are tested without packing."""
    return [sys.executable, "-c", textwrap.dedent(script)]


class NuGetConfigTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)

    def test_offline_config_clears_every_inherited_source(self) -> None:
        path = self.root / "NuGet.Config"
        smoke.write_nuget_config(path, self.root / "feed")
        text = path.read_text(encoding="utf-8")
        self.assertIn("<clear />", text)
        self.assertIn("officeagent-local", text)
        self.assertNotIn("nuget.org", text)

    def test_consumer_config_keeps_the_local_feed_first(self) -> None:
        path = self.root / "NuGet.Config"
        smoke.write_nuget_config(path, self.root / "feed", include_nuget_org=True)
        text = path.read_text(encoding="utf-8")
        self.assertIn("<clear />", text)
        self.assertLess(text.index("officeagent-local"), text.index("nuget.org"))


class ExportedBytesTests(unittest.TestCase):
    def docx(self, text: str, parts: dict[str, str] | None = None) -> bytes:
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w") as package:
            for name, content in (parts if parts is not None else {
                "[Content_Types].xml": "<Types />",
                "word/document.xml": f"<w:document><w:t>{text}</w:t></w:document>",
            }).items():
                package.writestr(name, content)
        return buffer.getvalue()

    def test_a_valid_package_containing_the_text_is_accepted(self) -> None:
        smoke.verify_docx_bytes(self.docx("Contoso Research"), "Contoso Research")

    def test_non_package_bytes_are_rejected(self) -> None:
        with self.assertRaises(smoke.SmokeError):
            smoke.verify_docx_bytes(b"not a package at all", "Contoso Research")

    def test_a_package_missing_the_main_part_is_rejected(self) -> None:
        content = self.docx("", parts={"[Content_Types].xml": "<Types />"})
        with self.assertRaises(smoke.SmokeError):
            smoke.verify_docx_bytes(content, "Contoso Research")

    def test_a_package_without_the_expected_text_is_rejected(self) -> None:
        with self.assertRaises(smoke.SmokeError):
            smoke.verify_docx_bytes(self.docx("Something else"), "Contoso Research")


class StdioBoundaryTests(unittest.TestCase):
    """The harness must fail on its own schedule instead of waiting forever."""

    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.cwd = Path(self.temporary.name)
        self.env = {"PATH": "", "SYSTEMROOT": ""}

    def start(self, script: str) -> smoke.StdioServer:
        server = smoke.StdioServer(fake_server(script), self.cwd, {**self.env})
        self.addCleanup(server.terminate)
        return server

    def test_a_silent_peer_fails_within_the_timeout(self) -> None:
        server = self.start(
            """
            import sys, time
            sys.stdin.readline()
            time.sleep(30)
            """
        )
        server.send({"jsonrpc": "2.0", "id": 1, "method": "initialize", "params": {}})
        with self.assertRaises(smoke.SmokeError) as caught:
            server.read_message(timeout=2)
        self.assertIn("no response", str(caught.exception))

    def test_a_peer_that_closes_stdout_is_reported_not_awaited(self) -> None:
        server = self.start(
            """
            import os, sys, time
            sys.stdin.readline()
            sys.stdout.flush()
            os.close(1)
            time.sleep(30)
            """
        )
        server.send({"jsonrpc": "2.0", "id": 1, "method": "initialize", "params": {}})
        with self.assertRaises(smoke.SmokeError) as caught:
            server.read_message(timeout=10)
        self.assertIn("closed stdout", str(caught.exception))

    def test_a_peer_that_ignores_stdin_eof_is_reported(self) -> None:
        server = self.start(
            """
            import time
            time.sleep(30)
            """
        )
        server.close_stdin()
        with self.assertRaises(smoke.SmokeError) as caught:
            server.wait(timeout=2)
        self.assertIn("did not exit", str(caught.exception))

    def test_a_peer_that_exits_on_stdin_eof_reports_its_code(self) -> None:
        server = self.start(
            """
            import sys
            sys.stdin.read()
            sys.exit(0)
            """
        )
        server.close_stdin()
        self.assertEqual(0, server.wait(timeout=20))

    def test_heavy_stderr_does_not_block_the_response(self) -> None:
        """The exact deadlock the drain thread exists to prevent."""
        server = self.start(
            """
            import sys, json
            sys.stdin.readline()
            sys.stderr.write("noise " * 20000 + "\\n")
            sys.stderr.flush()
            sys.stdout.write(json.dumps({"jsonrpc": "2.0", "id": 1, "result": {"ok": True}}) + "\\n")
            sys.stdout.flush()
            """
        )
        result = server.request("initialize", timeout=20)
        self.assertEqual({"ok": True}, result)

    def test_an_unrelated_frame_does_not_satisfy_the_request(self) -> None:
        server = self.start(
            """
            import sys, json
            sys.stdin.readline()
            sys.stdout.write(json.dumps({"jsonrpc": "2.0", "method": "notifications/message"}) + "\\n")
            sys.stdout.write(json.dumps({"jsonrpc": "2.0", "id": 1, "result": {"matched": True}}) + "\\n")
            sys.stdout.flush()
            """
        )
        self.assertEqual({"matched": True}, server.request("initialize", timeout=20))

    def test_a_tool_error_is_raised_not_returned(self) -> None:
        server = self.start(
            """
            import sys, json
            sys.stdin.readline()
            sys.stdout.write(json.dumps({
                "jsonrpc": "2.0", "id": 1,
                "result": {"isError": True, "content": [{"type": "text", "text": "boom"}]}
            }) + "\\n")
            sys.stdout.flush()
            """
        )
        with self.assertRaises(smoke.SmokeError):
            server.call_tool("apply_plan", {})

    def test_a_double_encoded_tool_payload_is_unwrapped(self) -> None:
        inner = json.dumps({"committed": True})
        server = self.start(
            f"""
            import sys, json
            sys.stdin.readline()
            sys.stdout.write(json.dumps({{
                "jsonrpc": "2.0", "id": 1,
                "result": {{"content": [{{"type": "text", "text": {json.dumps(inner)}}}]}}
            }}) + "\\n")
            sys.stdout.flush()
            """
        )
        self.assertEqual({"committed": True}, server.call_tool("apply_plan", {}))


class RequiredToolsTests(unittest.TestCase):
    def test_the_required_tool_set_matches_the_shipped_tool_names(self) -> None:
        """A renamed tool must break this list rather than silently weaken the smoke."""
        source = (ROOT / "src" / "OfficeAgent.AgentFramework" / "OfficeAgentTools.cs").read_text(
            encoding="utf-8"
        )
        for name in sorted(smoke.REQUIRED_TOOLS):
            self.assertIn(f'"{name}"', source, f"{name} is no longer defined by the adapter")


class PackageSetTests(unittest.TestCase):
    def test_missing_packages_are_reported_before_anything_is_installed(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            with self.assertRaises(smoke.SmokeError) as caught:
                smoke.smoke(Path(temporary), "9.9.9", "dotnet")
        self.assertIn("missing packages", str(caught.exception))

    def test_repository_version_is_read_from_the_build_props(self) -> None:
        self.assertRegex(smoke.repository_version(), r"^\d+\.\d+\.\d+")


if __name__ == "__main__":
    unittest.main()
