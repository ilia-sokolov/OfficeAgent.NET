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


class TemplateFixtureTests(unittest.TestCase):
    """The template the smoke builds must be a real, stable package."""

    def test_the_fixture_is_a_package_carrying_a_tagged_content_control(self) -> None:
        content = smoke.build_template_docx()
        self.assertEqual(b"PK", content[:2])
        with zipfile.ZipFile(io.BytesIO(content)) as package:
            self.assertEqual(
                {"[Content_Types].xml", "_rels/.rels", "word/document.xml"},
                set(package.namelist()),
            )
            document = package.read("word/document.xml").decode("utf-8")
        self.assertIn('w:val="CustomerName"', document)

    def test_the_fixture_is_byte_identical_on_every_build(self) -> None:
        """
        The server reports a hash of these bytes and the token binds to it, so a fixture
        that varied between runs would make the token meaningless and the failure obscure.
        """
        self.assertEqual(smoke.build_template_docx(), smoke.build_template_docx())


class TemplateWorkflowTests(unittest.TestCase):
    """
    The template checks must fail when the server misbehaves.

    A smoke step that cannot fail is worse than no step: it reports success for a surface
    that was never exercised, which is the exact defect this workflow was added to catch.
    """

    class FakeServer:
        def __init__(self, overrides: dict) -> None:
            self.overrides = overrides
            self.calls: list[str] = []

        def call_tool(self, name: str, arguments: dict) -> dict:
            self.calls.append(name)
            if name in self.overrides:
                value = self.overrides[name]
                return value(self.calls) if callable(value) else value
            return self.defaults(name)

        def defaults(self, name: str) -> dict:
            if name == "populate_template_batch":
                # The first commit is the stale one and must be refused; the second carries
                # the matching token and proceeds. A fake that committed both would make the
                # happy path pass for the wrong reason.
                if self.calls.count("populate_template_batch") == 1:
                    return {
                        "committed": False,
                        "items": [{"diagnostics": [{"code": "stale-batch-preview"}]}],
                    }
                return {"committed": True, "items": [{"document": {"itemId": "output-1"}}]}
            return self.fixed(name)

        @staticmethod
        def fixed(name: str) -> dict:
            if name == "import_document_content":
                return {"documentId": "template-1"}
            if name == "discover_template":
                return {"slots": [{"name": "CustomerName"}], "templateSha256": "abc"}
            if name == "preview_template_batch":
                return {
                    "isValid": True,
                    "token": {"templateSha256": "abc", "batchSha256": "def"},
                }
            if name == "export_document_content":
                return {
                    "contentBase64": base64.b64encode(
                        _package_with("Contoso Research")
                    ).decode("ascii")
                }
            raise AssertionError(f"unexpected tool {name}")

    def run_with(self, overrides: dict) -> None:
        smoke.verify_template_workflow(self.FakeServer(overrides))

    def test_the_happy_path_passes(self) -> None:
        self.run_with({})

    def test_a_missing_slot_is_reported(self) -> None:
        with self.assertRaises(smoke.SmokeError):
            self.run_with({"discover_template": {"slots": [], "templateSha256": "abc"}})

    def test_a_preview_without_a_token_is_reported(self) -> None:
        with self.assertRaises(smoke.SmokeError):
            self.run_with({"preview_template_batch": {"isValid": True, "token": {}}})

    def test_a_server_that_commits_a_stale_batch_is_reported(self) -> None:
        """The refusal is the point of the token; accepting it must fail the smoke."""
        with self.assertRaises(smoke.SmokeError):
            self.run_with({"populate_template_batch": {"committed": True, "items": []}})

    def test_a_refusal_that_does_not_name_the_reason_is_reported(self) -> None:
        def populate(calls: list[str]) -> dict:
            first = calls.count("populate_template_batch") == 1
            if first:
                return {"committed": False, "items": []}
            return {"committed": True, "items": [{"document": {"itemId": "output-1"}}]}

        with self.assertRaises(smoke.SmokeError):
            self.run_with({"populate_template_batch": populate})

    def test_an_output_that_lacks_the_bound_value_is_reported(self) -> None:
        """The produced bytes are read, not the status field that claimed to produce them."""
        with self.assertRaises(smoke.SmokeError):
            self.run_with(
                {
                    "export_document_content": {
                        "contentBase64": base64.b64encode(
                            _package_with("Someone Else")
                        ).decode("ascii")
                    }
                }
            )


def _package_with(text: str) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as package:
        package.writestr("[Content_Types].xml", "<Types/>")
        package.writestr("word/document.xml", f"<w:document>{text}</w:document>")
    return buffer.getvalue()


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
