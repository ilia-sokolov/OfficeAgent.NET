#!/usr/bin/env python3
"""Smoke-test locally built OfficeAgent packages on the current platform.

Installs the packed MCP tool into an empty tool directory backed by a fresh
package cache, drives a complete document workflow over stdio, checks the
bounded failure modes that would otherwise hang CI, and builds a direct .NET
consumer that loads every format module from package references.

This exercises packaged artifacts offline. It is deliberately separate from any
post-publish smoke against the public registry: a local feed proves packaging,
not publication.
"""

from __future__ import annotations

import argparse
import base64
import io
import json
import os
import platform
import socket
import re
import subprocess
import sys
import tempfile
import urllib.error
import urllib.request
import threading
import time
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]

# The tools a zero-configuration server must offer. It mints a session
# connection, so connection-addressed and session tools are both expected.
REQUIRED_TOOLS = {
    "inspect_document",
    "find_in_document",
    "preview_plan",
    "apply_plan",
    "create_document",
    "import_document_content",
    "export_document_content",
    "describe_capabilities",
    # The template workflow an agent is meant to run: look before binding, validate the
    # whole batch before committing any of it, and bind the commit to what was reviewed.
    # All three landed on the direct client first and reached the adapters in V09-14, so
    # the installed package is where their absence would otherwise go unnoticed.
    "discover_template",
    "preview_template_batch",
    "populate_template_batch",
}

STARTUP_TIMEOUT_SECONDS = 90
CALL_TIMEOUT_SECONDS = 60
SHUTDOWN_TIMEOUT_SECONDS = 30
HTTP_START_SECONDS = 90
PROBE_TIMEOUT_SECONDS = 30


class SmokeError(RuntimeError):
    pass


def repository_version() -> str:
    text = (ROOT / "Directory.Build.props").read_text(encoding="utf-8")
    match = re.search(r"<Version>([^<]+)</Version>", text)
    if not match:
        raise SmokeError("Directory.Build.props has no Version")
    return match.group(1)


def write_nuget_config(path: Path, package_source: Path, include_nuget_org: bool = False) -> None:
    """Write an explicit config with every inherited source cleared.

    The packed tool carries its own dependencies, so installing it proves
    packaging only when no other feed is reachable. A library consumer
    legitimately needs third-party transitive packages, so that config keeps
    nuget.org as a second source behind the local feed.
    """
    upstream = (
        '    <add key="nuget.org" value="https://api.nuget.org/v3/index.json" protocolVersion="3" />\n'
        if include_nuget_org
        else ""
    )
    path.write_text(
        """<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <packageSources>
    <clear />
    <add key="officeagent-local" value="{source}" />
{upstream}  </packageSources>
</configuration>
""".format(source=package_source.as_posix(), upstream=upstream),
        encoding="utf-8",
        newline="\n",
    )


def run(command: list[str], cwd: Path, env: dict[str, str], timeout: int = 600) -> subprocess.CompletedProcess[str]:
    result = subprocess.run(
        command, cwd=cwd, env=env, text=True, capture_output=True, check=False, timeout=timeout
    )
    if result.returncode:
        raise SmokeError(
            f"command failed ({result.returncode}): {' '.join(command)}\n{result.stdout}\n{result.stderr}"
        )
    return result


def resolved_package_problems(assets: dict, version: str, packages: tuple[str, ...]) -> list[str]:
    """Every package that restore did not resolve as a package at exactly the version.

    Reads a restored project's obj/project.assets.json. Package ids compare case-insensitively, as
    NuGet does; versions compare exactly. A missing package, another version, or a project reference
    in place of the package is a problem.
    """
    libraries = assets.get("libraries") or {}
    problems = []
    for package in packages:
        resolved = sorted(
            key for key in libraries if key.split("/", 1)[0].lower() == package.lower()
        )
        expected = f"{package}/{version}"
        exact = [key for key in resolved if key.split("/", 1)[1] == version]
        if not exact or len(resolved) != 1:
            problems.append(f"{package}: resolved {resolved or 'nothing'}, expected {expected!r}")
        elif libraries[exact[0]].get("type") != "package":
            problems.append(f"{package}: resolved {exact[0]!r} as {libraries[exact[0]].get('type')!r}, not a package")
    return problems


def isolated_env(cache: Path, nuget_config: Path) -> dict[str, str]:
    env = dict(os.environ)
    env["NUGET_PACKAGES"] = str(cache)
    # A fallback folder would satisfy a restore from outside the fresh cache.
    env.pop("NUGET_FALLBACK_PACKAGES", None)
    env["DOTNET_NOLOGO"] = "1"
    env["DOTNET_CLI_TELEMETRY_OPTOUT"] = "1"
    env["DOTNET_SKIP_FIRST_TIME_EXPERIENCE"] = "1"
    env["RESTORE_CONFIG_FILE"] = str(nuget_config)
    return env


class StdioServer:
    """A packaged MCP server driven over stdio with every wait bounded.

    stderr is drained on a background thread. A server that logs more than the
    pipe buffer would otherwise block forever while we wait on stdout, which is
    exactly the hang this task exists to prevent.
    """

    def __init__(self, command: list[str], cwd: Path, env: dict[str, str]) -> None:
        self._process = subprocess.Popen(
            command,
            cwd=cwd,
            env=env,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            bufsize=1,
        )
        self._stderr: list[str] = []
        self._drain = threading.Thread(target=self._drain_stderr, daemon=True)
        self._drain.start()
        self._next_id = 0

        # Protocol counters. Every frame in either direction passes through `send`,
        # `send_raw` or `read_message`, so counting there cannot be bypassed by a caller
        # that builds its own payload. Unlike wall-clock time these are deterministic
        # properties of the exchange, which is what makes them worth gating on.
        self.reset_counters()

    def reset_counters(self) -> None:
        """Zero the protocol counters, so one iteration's traffic is attributable to it."""
        self.frames_sent = 0
        self.frames_received = 0
        self.bytes_sent = 0
        self.bytes_received = 0
        self.tool_calls = 0

    @property
    def protocol_counters(self) -> dict[str, int]:
        return {
            "FramesSent": self.frames_sent,
            "FramesReceived": self.frames_received,
            "BytesSent": self.bytes_sent,
            "BytesReceived": self.bytes_received,
            "ToolCalls": self.tool_calls,
        }

    def _drain_stderr(self) -> None:
        assert self._process.stderr is not None
        for line in self._process.stderr:
            self._stderr.append(line.rstrip("\n"))

    @property
    def stderr_text(self) -> str:
        return "\n".join(self._stderr)

    def send(self, payload: dict) -> None:
        self.send_raw(json.dumps(payload))

    def send_raw(self, line: str) -> None:
        assert self._process.stdin is not None
        self.frames_sent += 1
        self.bytes_sent += len(line.encode("utf-8")) + 1
        self._process.stdin.write(line + "\n")
        self._process.stdin.flush()

    def read_message(self, timeout: int = CALL_TIMEOUT_SECONDS) -> dict:
        """Read one JSON-RPC frame, failing rather than blocking forever."""
        assert self._process.stdout is not None
        result: list[str] = []

        def read_one() -> None:
            line = self._process.stdout.readline()
            if line:
                result.append(line)

        reader = threading.Thread(target=read_one, daemon=True)
        reader.start()
        reader.join(timeout)
        if reader.is_alive():
            raise SmokeError(
                f"server produced no response within {timeout}s; stderr:\n{self.stderr_text}"
            )
        if not result:
            raise SmokeError(f"server closed stdout unexpectedly; stderr:\n{self.stderr_text}")
        self.frames_received += 1
        self.bytes_received += len(result[0].encode("utf-8"))
        try:
            return json.loads(result[0])
        except json.JSONDecodeError as exc:
            raise SmokeError(f"server emitted a non-JSON frame: {result[0]!r}") from exc

    def request(self, method: str, params: dict | None = None, timeout: int = CALL_TIMEOUT_SECONDS) -> dict:
        self._next_id += 1
        request_id = self._next_id
        self.send({"jsonrpc": "2.0", "id": request_id, "method": method, "params": params or {}})
        while True:
            message = self.read_message(timeout)
            # Notifications and server-initiated requests are not our answer.
            if message.get("id") != request_id:
                continue
            if "error" in message:
                raise SmokeError(f"{method} failed: {json.dumps(message['error'])}")
            return message.get("result", {})

    def call_tool(self, name: str, arguments: dict) -> dict:
        payload = self.call_tool_json(name, arguments)
        if not isinstance(payload, dict):
            raise SmokeError(f"tool {name} returned {type(payload).__name__}, expected an object")
        return payload

    def call_tool_json(self, name: str, arguments: dict):
        self.tool_calls += 1
        result = self.request("tools/call", {"name": name, "arguments": arguments})
        if result.get("isError"):
            raise SmokeError(f"tool {name} returned an error: {json.dumps(result)[:800]}")
        blocks = [block for block in result.get("content", []) if block.get("type") == "text"]
        if not blocks:
            raise SmokeError(f"tool {name} returned no text content")
        text = blocks[0]["text"]
        try:
            payload = json.loads(text)
        except json.JSONDecodeError as exc:
            raise SmokeError(f"tool {name} returned non-JSON text: {text[:400]}") from exc
        # Some hosts hand back the JSON document as a quoted string rather than an
        # object; unwrap one level so callers always see the structured payload.
        if isinstance(payload, str):
            try:
                payload = json.loads(payload)
            except json.JSONDecodeError as exc:
                raise SmokeError(f"tool {name} returned a string, not a JSON object: {text[:400]}") from exc
        return payload

    def close_stdin(self) -> None:
        assert self._process.stdin is not None
        self._process.stdin.close()

    def wait(self, timeout: int = SHUTDOWN_TIMEOUT_SECONDS) -> int:
        try:
            return self._process.wait(timeout=timeout)
        except subprocess.TimeoutExpired as exc:
            self.terminate()
            raise SmokeError(f"server did not exit within {timeout}s of stdin EOF") from exc

    def terminate(self) -> None:
        """Only ever applied to the process this harness started."""
        if self._process.poll() is None:
            self._process.terminate()
            try:
                self._process.wait(timeout=10)
            except subprocess.TimeoutExpired:
                self._process.kill()
                self._process.wait(timeout=10)

    def __enter__(self) -> "StdioServer":
        return self

    def __exit__(self, *exc_info: object) -> None:
        self.terminate()


def initialize(server: StdioServer) -> dict:
    result = server.request(
        "initialize",
        {
            "protocolVersion": "2024-11-05",
            "capabilities": {},
            "clientInfo": {"name": "officeagent-packaged-smoke", "version": "1.0"},
        },
        timeout=STARTUP_TIMEOUT_SECONDS,
    )
    server.send({"jsonrpc": "2.0", "method": "notifications/initialized"})
    return result


def reported_version_matches(reported: str | None, expected: str) -> bool:
    """Exact after removing the source-control suffix: 1.0.0-rc.30 is not 1.0.0-rc.3."""
    return bool(reported) and reported.split("+", 1)[0] == expected


def verify_document_workflow(server: StdioServer, version: str) -> None:
    info = initialize(server)
    server_info = info.get("serverInfo", {})
    reported = server_info.get("version", "")
    if not reported_version_matches(reported, version):
        raise SmokeError(f"server reports version {reported!r}, expected the candidate {version!r}")

    tools = {tool["name"] for tool in server.request("tools/list").get("tools", [])}
    missing = sorted(REQUIRED_TOOLS - tools)
    if missing:
        raise SmokeError(f"packaged server is missing tools: {missing}; offered {sorted(tools)}")
    if not tools:
        raise SmokeError("packaged server offered no tools")

    # Discovery first, then act only on what it advertised. This is the whole point of
    # the capability tool: the workflow below is chosen from the server's own answer
    # rather than from assumptions baked into this script.
    capabilities = server.call_tool("describe_capabilities", {})
    # Property names are camelCase as of 1.0. Enum values such as "Word" and "Tracked" are
    # member names written by the enum converter and were never affected.
    contract = capabilities.get("contracts", {}).get("editPlan")
    if not contract:
        raise SmokeError(f"discovery reported no edit-plan contract: {json.dumps(capabilities)[:400]}")

    word = next(
        (fmt for fmt in capabilities.get("formats", []) if fmt.get("format") == "Word"), None)
    if word is None:
        raise SmokeError("discovery reported no Word format on a server that offers Word tools")
    if "changeText" not in word.get("operations", []):
        raise SmokeError(f"discovery does not advertise changeText for Word: {json.dumps(word)[:400]}")
    if "Tracked" not in word.get("changeModes", []):
        raise SmokeError(f"discovery does not advertise Tracked for Word: {json.dumps(word)[:400]}")
    if not capabilities.get("limits", {}).get("maximumCompressedBytes"):
        raise SmokeError("discovery reported no ingestion ceiling")
    print(f"discovery=passed contract={contract} word-operations={len(word['operations'])}")

    created = server.call_tool(
        "create_document",
        {
            "connectionId": "session",
            "name": "platform-smoke.docx",
            "planJson": json.dumps(
                {
                    "contractVersion": contract,
                    "operations": [
                        {
                            "op": "changeText",
                            "target": {"paraId": "auto-0000", "expect": ""},
                            "with": "Prepared for Northwind Labs.",
                        }
                    ]
                }
            ),
        },
    )
    if not created.get("committed"):
        raise SmokeError(f"create_document did not commit: {json.dumps(created)[:800]}")
    document_id = created.get("outputDocumentId")
    if not document_id:
        raise SmokeError("create_document returned no outputDocumentId")

    inspected = server.call_tool(
        "inspect_document", {"connectionId": "session", "documentId": document_id}
    )
    if "Northwind Labs" not in json.dumps(inspected):
        raise SmokeError("inspection does not contain the created text")

    hits = server.call_tool_json(
        "find_in_document",
        {"connectionId": "session", "documentId": document_id, "pattern": "Northwind Labs"},
    )
    if not isinstance(hits, list) or not hits:
        raise SmokeError(f"find_in_document returned no hits: {json.dumps(hits)[:400]}")
    hit = hits[0]

    plan_json = json.dumps(
        {
            "operations": [
                {
                    "op": "changeText",
                    "target": {
                        "paraId": hit["paraId"],
                        "expect": hit["expect"],
                        "occurrence": hit.get("occurrence", 0),
                    },
                    "with": "Contoso Research",
                }
            ]
        }
    )

    previewed = server.call_tool(
        "preview_plan",
        {"connectionId": "session", "documentId": document_id, "planJson": plan_json},
    )
    if not previewed.get("isValid"):
        raise SmokeError(f"preview_plan rejected the plan: {json.dumps(previewed)[:800]}")

    applied = server.call_tool(
        "apply_plan",
        {"connectionId": "session", "documentId": document_id, "planJson": plan_json},
    )
    if not applied.get("committed"):
        raise SmokeError(f"apply_plan did not commit: {json.dumps(applied)[:800]}")

    exported = server.call_tool(
        "export_document_content", {"connectionId": "session", "documentId": document_id}
    )
    payload = exported.get("contentBase64")
    if not payload:
        raise SmokeError("export_document_content returned no bytes")
    verify_docx_bytes(base64.b64decode(payload), "Contoso Research")
    print("mcp-workflow=passed")

    verify_template_workflow(server)


CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument'
    '.wordprocessingml.document.main+xml"/>'
    "</Types>"
)

PACKAGE_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    '/officeDocument" Target="word/document.xml"/>'
    "</Relationships>"
)

TEMPLATE_DOCUMENT = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
    "<w:p><w:r><w:t xml:space=\"preserve\">Quote for </w:t></w:r>"
    "<w:sdt><w:sdtPr><w:tag w:val=\"CustomerName\"/><w:id w:val=\"101\"/></w:sdtPr>"
    "<w:sdtContent><w:r><w:t>CUSTOMER</w:t></w:r></w:sdtContent></w:sdt></w:p>"
    "<w:p/>"
    "</w:body></w:document>"
)


def build_template_docx() -> bytes:
    """
    A minimal .docx carrying one content control, built here rather than committed.

    No plan operation creates a content control, so the packaged server cannot mint a
    template with create_document: fill only fills a control that already exists. A .docx is
    a zip of a few XML parts, so the fixture is assembled in-process with the standard
    library. That keeps the smoke free of Office, of the .NET test tree, and of a binary
    checked into the repository whose bytes nothing here could explain.
    """
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as package:
        for name, content in (
            ("[Content_Types].xml", CONTENT_TYPES),
            ("_rels/.rels", PACKAGE_RELS),
            ("word/document.xml", TEMPLATE_DOCUMENT),
        ):
            # A fixed timestamp keeps the bytes, and therefore the template hash the server
            # reports, identical on every run and every platform.
            info = zipfile.ZipInfo(name, date_time=(2026, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            package.writestr(info, content)
    return buffer.getvalue()


def verify_template_workflow(server: StdioServer) -> None:
    """
    The template preflight contract, driven through the installed package.

    Discovery, batch preflight and the hash-bound commit reached the adapters late: they
    existed on the direct .NET client while neither adapter exposed them, so an agent could
    commit a batch it had never previewed. Listing the three tools proves they are offered.
    Only running them proves they work from a package, which is the claim that matters.
    """
    imported = server.call_tool(
        "import_document_content",
        {
            "connectionId": "session",
            "name": "quote-template.docx",
            "contentBase64": base64.b64encode(build_template_docx()).decode("ascii"),
        },
    )
    template_id = imported.get("documentId")
    if not template_id:
        raise SmokeError(f"import_document_content returned no documentId: {json.dumps(imported)[:400]}")

    discovery = server.call_tool(
        "discover_template", {"connectionId": "session", "documentId": template_id}
    )
    slots = [slot.get("name") for slot in discovery.get("slots", [])]
    if "CustomerName" not in slots:
        raise SmokeError(f"discover_template did not report the CustomerName slot: {slots}")
    if not discovery.get("templateSha256"):
        raise SmokeError("discover_template reported no template hash")

    def batch(customer: str) -> str:
        return json.dumps(
            {
                "items": [
                    {
                        "outputName": "quote-smoke.docx",
                        "binding": {
                            "values": {"CustomerName": customer},
                            "missingValueBehavior": "Ignore",
                        },
                    }
                ]
            }
        )

    preview = server.call_tool(
        "preview_template_batch",
        {"connectionId": "session", "documentId": template_id, "requestJson": batch("Contoso Research")},
    )
    if not preview.get("isValid"):
        raise SmokeError(f"preview_template_batch refused a valid batch: {json.dumps(preview)[:600]}")
    token = preview.get("token")
    if not token or not token.get("templateSha256") or not token.get("batchSha256"):
        raise SmokeError(f"preview_template_batch issued no usable token: {json.dumps(preview)[:600]}")

    # The refusal first, so a commit cannot be credited to a token it never honoured. The
    # reviewed batch said Contoso Research; this one asks for something else.
    stale = server.call_tool(
        "populate_template_batch",
        {
            "connectionId": "session",
            "documentId": template_id,
            "requestJson": batch("Someone Else"),
            "expectedTokenJson": json.dumps(token),
        },
    )
    if stale.get("committed"):
        raise SmokeError("a commit whose batch changed after review was accepted")
    if "stale-batch-preview" not in json.dumps(stale):
        raise SmokeError(f"a stale commit was refused without naming why: {json.dumps(stale)[:600]}")

    committed = server.call_tool(
        "populate_template_batch",
        {
            "connectionId": "session",
            "documentId": template_id,
            "requestJson": batch("Contoso Research"),
            "expectedTokenJson": json.dumps(token),
        },
    )
    if not committed.get("committed"):
        raise SmokeError(f"a token-bound commit was refused: {json.dumps(committed)[:600]}")

    output_id = next(
        (item["document"]["itemId"]
         for item in committed.get("items", [])
         if isinstance(item.get("document"), dict) and item["document"].get("itemId")),
        None,
    )
    if not output_id:
        raise SmokeError(f"populate_template_batch reported no output document: {json.dumps(committed)[:600]}")

    exported = server.call_tool(
        "export_document_content", {"connectionId": "session", "documentId": output_id}
    )
    payload = exported.get("contentBase64")
    if not payload:
        raise SmokeError("the populated output exported no bytes")
    # Read the produced document, not the status field that claimed to have produced it.
    verify_docx_bytes(base64.b64decode(payload), "Contoso Research")
    print(f"template-workflow=passed slots={len(slots)}")


def verify_docx_bytes(content: bytes, expected_text: str) -> None:
    """The exported bytes must be a real package, not just a non-empty blob."""
    if content[:2] != b"PK":
        raise SmokeError("exported bytes are not an OPC package")
    with zipfile.ZipFile(io.BytesIO(content)) as package:
        names = set(package.namelist())
        required = {"[Content_Types].xml", "word/document.xml"}
        if not required.issubset(names):
            raise SmokeError(f"exported package is missing parts: {sorted(required - names)}")
        document = package.read("word/document.xml").decode("utf-8")
    if expected_text not in document:
        raise SmokeError(f"exported document does not contain {expected_text!r}")


def verify_stdin_eof_exit(command: list[str], cwd: Path, env: dict[str, str]) -> None:
    """Closing stdin must end the process, not leave an orphan holding CI open."""
    with StdioServer(command, cwd, env) as server:
        initialize(server)
        server.close_stdin()
        code = server.wait()
    print(f"eof-exit=passed code={code}")


def verify_malformed_request(command: list[str], cwd: Path, env: dict[str, str]) -> str:
    """A malformed frame must not wedge the server.

    JSON-RPC says a parse error should be answered with code -32700. The MCP
    stdio transport this server is built on logs the parse failure and drops the
    frame instead, so the property worth asserting is liveness, not the error
    response: after garbage input the server must still answer a valid request,
    or exit. Either is bounded. Silence while still running is the failure, and
    it is what would hang a CI job.
    """
    with StdioServer(command, cwd, env) as server:
        initialize(server)
        server.send_raw("{ this is not valid json")
        try:
            tools = server.request("tools/list", timeout=PROBE_TIMEOUT_SECONDS)
            if not tools.get("tools"):
                raise SmokeError("server answered after a malformed frame but offered no tools")
            outcome = "dropped-server-responsive"
        except SmokeError as exc:
            if "closed stdout" in str(exc):
                outcome = "stream-closed"
            elif "no response" in str(exc):
                raise SmokeError(
                    "server stopped responding after a malformed frame; a client would hang here"
                ) from exc
            else:
                raise
        server.terminate()
    print(f"malformed-request=passed outcome={outcome}")
    return outcome


def verify_missing_configuration(command: list[str], cwd: Path, env: dict[str, str]) -> None:
    """An explicitly named but absent config file must fail fast and say why."""
    missing = cwd / "definitely-absent.json"
    result = subprocess.run(
        command + ["--config", str(missing)],
        cwd=cwd,
        env=env,
        text=True,
        capture_output=True,
        check=False,
        timeout=STARTUP_TIMEOUT_SECONDS,
        stdin=subprocess.DEVNULL,
    )
    if result.returncode == 0:
        raise SmokeError("a missing --config file was accepted instead of failing")
    combined = f"{result.stdout}\n{result.stderr}"
    if "definitely-absent" not in combined:
        raise SmokeError(f"startup failure does not name the missing file:\n{combined[:800]}")
    # The failure is bounded and names the file, which is what this criterion asks
    # for. How gracefully it fails is recorded rather than asserted: the server
    # currently terminates on an unhandled exception instead of printing a single
    # diagnostic, so the exit code and first line are reported for the evidence.
    first = next((line.strip() for line in combined.splitlines() if line.strip()), "")
    print(f"missing-config=passed code={result.returncode} detail={first[:120]}")


def free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as probe:
        probe.bind(("127.0.0.1", 0))
        return probe.getsockname()[1]


def verify_http_health(executable: Path, workdir: Path, env: dict[str, str], version: str) -> None:
    """The installed server over HTTP, as the container runs it: /healthz names the exact version."""
    port = free_port()
    http_env = dict(env)
    http_env["ASPNETCORE_URLS"] = f"http://127.0.0.1:{port}"
    process = subprocess.Popen(
        [str(executable)], cwd=workdir, env=http_env,
        stdout=subprocess.DEVNULL, stderr=subprocess.PIPE, text=True)
    try:
        deadline = time.monotonic() + HTTP_START_SECONDS
        body = None
        while time.monotonic() < deadline:
            if process.poll() is not None:
                raise SmokeError(f"HTTP server exited with {process.returncode}: {process.stderr.read()[:800]}")
            try:
                with urllib.request.urlopen(f"http://127.0.0.1:{port}/healthz", timeout=5) as response:
                    body = json.loads(response.read().decode("utf-8"))
                    break
            except (urllib.error.URLError, ConnectionError, TimeoutError):
                time.sleep(0.5)
        if body is None:
            raise SmokeError(f"/healthz did not answer within {HTTP_START_SECONDS} s")
        if body.get("status") != "ok" or not reported_version_matches(body.get("version"), version):
            raise SmokeError(f"/healthz reports {body!r}, expected the candidate {version!r}")
        print(f"http-health=passed version={body.get('version')}")
    finally:
        process.kill()
        try:
            process.wait(timeout=SHUTDOWN_TIMEOUT_SECONDS)
        except subprocess.TimeoutExpired:
            pass


def verify_installed_tool_payload(tool_root: Path, version: str) -> None:
    """The installed tool must carry its dependencies, not just its own assembly."""
    assemblies = {path.name for path in tool_root.rglob("*.dll")}
    expected = {
        "OfficeAgent.Core.dll",
        "OfficeAgent.Word.dll",
        "OfficeAgent.PowerPoint.dll",
        "OfficeAgent.Excel.dll",
        "OfficeAgent.Abstractions.dll",
        "DocumentFormat.OpenXml.dll",
    }
    missing = sorted(expected - assemblies)
    if missing:
        raise SmokeError(f"installed tool is missing dependencies: {missing}")
    print(f"installed-payload=passed assemblies={len(assemblies)} version={version}")


DIRECT_CONSUMER_PROGRAM = """using DocumentFormat.OpenXml.Packaging;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Word;
using OfficeAgent.PowerPoint;
using OfficeAgent.Excel;

var client = new OfficeAgentClient(new WordModule(), new PowerPointModule(), new ExcelModule());

foreach (var (name, anchorText) in new[]
         {
             ("smoke.docx", "auto-0000"),
             ("smoke.pptx", "slide256/shape2/p0"),
             ("smoke.xlsx", (string?)null)
         })
{
    var bytes = client.CreateBlank(name);
    if (bytes.Length == 0) throw new InvalidOperationException($"{name} produced no bytes");

    var inspection = client.Inspect(bytes);
    if (inspection.Snapshot.ETag.Length == 0)
        throw new InvalidOperationException($"{name} produced no snapshot");
    Console.WriteLine($"format={inspection.Format} name={name} anchors={inspection.Anchors.Count}");
}

// One validated operation, proving the Word module is loaded and functioning rather
// than merely referenced.
var document = client.CreateBlank("edit.docx");
using var handle = new MemoryStream(document, writable: false);
var stream = new StreamHandle(handle, "edit.docx");
var inspected = client.Inspect(stream);
var plan = new DocumentPlan
{
    Snapshot = inspected.Snapshot,
    Operations = new PlanOperation[]
    {
        new ChangeTextOp
        {
            Target = new TextSpanAnchor { ParaId = "auto-0000", Expect = "" },
            With = "Prepared for Northwind Labs."
        }
    }
};
var preview = client.Preview(stream, plan);
if (!preview.IsValid)
    throw new InvalidOperationException(
        "preview failed: " + string.Join("; ", preview.Errors.Select(e => $"{e.Code}: {e.Message}")));

using var applied = client.Commit(stream, plan);
if (!applied.Committed) throw new InvalidOperationException("commit failed");
var edited = applied.ToBytes();
using (var package = WordprocessingDocument.Open(new MemoryStream(edited, writable: false), isEditable: false))
{
    var text = package.MainDocumentPart?.Document?.Body?.InnerText ?? "";
    if (!text.Contains("Northwind Labs", StringComparison.Ordinal))
        throw new InvalidOperationException("edited document does not contain the expected text");
}

Console.WriteLine("direct-consumer=passed");
"""


DIRECT_CONSUMER_PROJECT = """<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>Exe</OutputType>
    <TargetFramework>net8.0</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="OfficeAgent.Core" Version="{version}" />
    <PackageReference Include="OfficeAgent.Word" Version="{version}" />
    <PackageReference Include="OfficeAgent.PowerPoint" Version="{version}" />
    <PackageReference Include="OfficeAgent.Excel" Version="{version}" />
  </ItemGroup>
</Project>
"""


def verify_direct_consumer(fixture: Path, artifacts: Path, version: str, dotnet: str, cache: Path) -> None:
    consumer = fixture / "direct-consumer"
    consumer.mkdir()
    (consumer / "DirectConsumer.csproj").write_text(
        DIRECT_CONSUMER_PROJECT.format(version=version), encoding="utf-8", newline="\n"
    )
    (consumer / "Program.cs").write_text(DIRECT_CONSUMER_PROGRAM, encoding="utf-8", newline="\n")
    config = consumer / "NuGet.Config"
    write_nuget_config(config, artifacts, include_nuget_org=True)
    env = isolated_env(cache, config)

    run([dotnet, "restore", "--configfile", str(config)], consumer, env)
    result = run(
        [dotnet, "run", "--configuration", "Release", "--no-restore"], consumer, env, timeout=900
    )
    if "direct-consumer=passed" not in result.stdout:
        raise SmokeError(f"direct consumer did not pass:\n{result.stdout}\n{result.stderr}")

    assets = json.loads((consumer / "obj" / "project.assets.json").read_text(encoding="utf-8"))
    resolved = set(assets.get("libraries", {}))
    for package in ("OfficeAgent.Core", "OfficeAgent.Word", "OfficeAgent.PowerPoint", "OfficeAgent.Excel"):
        if f"{package}/{version}" not in resolved:
            raise SmokeError(f"direct consumer did not resolve {package} {version}")
    print(result.stdout.strip())


def smoke(artifacts: Path, version: str, dotnet: str) -> None:
    required = [artifacts / f"OfficeAgent.{name}.{version}.nupkg" for name in
                ("Abstractions", "Core", "Word", "PowerPoint", "Excel", "AgentFramework", "SharePoint", "Mcp")]
    missing = [path.name for path in required if not path.is_file()]
    if missing:
        raise SmokeError(f"missing packages: {', '.join(missing)}")

    with tempfile.TemporaryDirectory(prefix="officeagent-platform-smoke-") as temporary:
        fixture = Path(temporary)
        cache = fixture / "package-cache"
        cache.mkdir()
        tool_root = fixture / "tools"
        tool_root.mkdir()
        config = fixture / "NuGet.Config"
        write_nuget_config(config, artifacts)
        env = isolated_env(cache, config)

        print(f"platform={platform.system()} release={platform.release()} arch={platform.machine()}")
        print(f"package-version={version}")

        run(
            [
                dotnet, "tool", "install", "OfficeAgent.Mcp",
                "--version", version,
                "--tool-path", str(tool_root),
                "--configfile", str(config),
            ],
            fixture,
            env,
        )
        verify_installed_tool_payload(tool_root, version)

        executable = tool_root / ("officeagent-mcp.exe" if os.name == "nt" else "officeagent-mcp")
        if not executable.is_file():
            raise SmokeError(f"installed tool executable not found at {executable}")
        command = [str(executable), "--stdio"]

        # An empty working directory: no appsettings.json, no configuration at all.
        workdir = fixture / "empty"
        workdir.mkdir()

        with StdioServer(command, workdir, env) as server:
            verify_document_workflow(server, version)
        verify_http_health(executable, workdir, env, version)

        verify_stdin_eof_exit(command, workdir, env)
        verify_malformed_request(command, workdir, env)
        verify_missing_configuration(command, workdir, env)
        verify_direct_consumer(fixture, artifacts, version, dotnet, cache)

        run([dotnet, "tool", "uninstall", "OfficeAgent.Mcp", "--tool-path", str(tool_root)], fixture, env)
        if executable.exists():
            raise SmokeError("tool uninstall left the executable in place")
        print("uninstall=passed")

    print("packaged-smoke=passed")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--artifacts", type=Path, required=True)
    parser.add_argument("--version", default=None)
    parser.add_argument("--dotnet", default="dotnet")
    args = parser.parse_args()
    try:
        smoke(args.artifacts.resolve(), args.version or repository_version(), args.dotnet)
    except (OSError, SmokeError, subprocess.TimeoutExpired, zipfile.BadZipFile) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    print("Packaged artifact smoke completed successfully.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
