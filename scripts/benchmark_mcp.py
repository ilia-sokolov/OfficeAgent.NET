#!/usr/bin/env python3
"""
Measure the installed MCP path: protocol calls, bytes, and wall-clock time.

The direct .NET benchmark measures the engine. This measures what an agent actually pays to
use it through a packaged server, which is a different question and has a different answer.
It installs the packed tool into an empty tool directory with a fresh package cache, exactly
as the packaged smoke does, and drives the same stdio protocol an MCP client would.

Two of the three metrics here are deterministic. The number of JSON-RPC frames exchanged and
the number of bytes in them are properties of the protocol and the payloads, not of the
machine: the same workflow against the same server produces the same counts every time. That
makes them worth gating on, in a way wall-clock time on this path is not. Process startup,
stdio buffering and the tool host make MCP timing noisier than the direct path, which is
already too noisy to gate, so timing here stays reported rather than enforced.

Results use the same schema as the direct harness, with the protocol counters added, so one
reporter renders both.
"""

from __future__ import annotations

import argparse
import base64
import json
import os
import platform
import subprocess
import sys
import tempfile
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parent))

from smoke_packaged_artifacts import (  # noqa: E402
    SmokeError,
    StdioServer,
    build_template_docx,
    initialize,
    isolated_env,
    repository_version,
    run,
    verify_docx_bytes,
    write_nuget_config,
)

SCHEMA_VERSION = "1.1"


def git(*args: str) -> str:
    try:
        return subprocess.run(
            ["git", *args], capture_output=True, text=True, check=False
        ).stdout.strip()
    except OSError:
        return ""


def word_plan(text: str) -> str:
    return json.dumps(
        {
            "operations": [
                {
                    "op": "changeText",
                    "target": {"paraId": "auto-0000", "expect": ""},
                    "with": text,
                }
            ]
        }
    )


def document_workflow(server: StdioServer, tag: str) -> dict[str, Any]:
    """
    Create, inspect, find, preview, apply and export one document.

    Names carry the iteration tag because the session connection outlives an iteration: a
    fixed name makes every repetition after the first fail with `already-exists`, which is
    the server behaving correctly and the benchmark asking for the wrong thing.
    """
    created = server.call_tool(
        "create_document",
        {"connectionId": "session", "name": f"benchmark-{tag}.docx", "planJson": ""},
    )
    document_id = created.get("documentId") or created.get("outputDocumentId")
    if not document_id:
        raise SmokeError(f"create_document returned no id: {json.dumps(created)[:300]}")

    server.call_tool("inspect_document", {"connectionId": "session", "documentId": document_id})
    plan = word_plan("Prepared for Northwind Labs.")
    previewed = server.call_tool(
        "preview_plan",
        {"connectionId": "session", "documentId": document_id, "planJson": plan},
    )
    if not previewed.get("isValid"):
        raise SmokeError(f"preview_plan refused a valid plan: {json.dumps(previewed)[:400]}")

    applied = server.call_tool(
        "apply_plan",
        {"connectionId": "session", "documentId": document_id, "planJson": plan},
    )
    if not applied.get("committed"):
        raise SmokeError(f"apply_plan did not commit: {json.dumps(applied)[:400]}")

    # find_in_document answers with an array, not an object, so it needs the raw accessor.
    hits = server.call_tool_json(
        "find_in_document",
        {"connectionId": "session", "documentId": document_id, "pattern": "Northwind",
         "regex": False, "wholeWord": False, "caseSensitive": False,
         "spreadsheetValueView": "both"},
    )
    if not isinstance(hits, list) or not hits:
        raise SmokeError(f"find_in_document found nothing after the edit: {json.dumps(hits)[:300]}")

    exported = server.call_tool(
        "export_document_content", {"connectionId": "session", "documentId": document_id}
    )
    payload = exported.get("contentBase64")
    if not payload:
        raise SmokeError("export_document_content returned no bytes")
    content = base64.b64decode(payload)
    verify_docx_bytes(content, "Northwind Labs")
    return {"output_bytes": len(content), "work_units": 6}


def template_workflow(server: StdioServer, tag: str) -> dict[str, Any]:
    """Import a template, discover it, preflight a batch and commit it bound to that preview."""
    imported = server.call_tool(
        "import_document_content",
        {
            "connectionId": "session",
            "name": f"benchmark-template-{tag}.docx",
            "contentBase64": base64.b64encode(build_template_docx()).decode("ascii"),
        },
    )
    template_id = imported.get("documentId")
    if not template_id:
        raise SmokeError("import_document_content returned no documentId")

    discovery = server.call_tool(
        "discover_template", {"connectionId": "session", "documentId": template_id}
    )
    slots = [slot.get("name") for slot in discovery.get("slots", [])]
    if "CustomerName" not in slots:
        raise SmokeError(f"discover_template did not report the expected slot: {slots}")

    request = json.dumps(
        {
            "items": [
                {
                    "outputName": f"quote-{tag}-{index}.docx",
                    "binding": {
                        "values": {"CustomerName": f"Customer {index}"},
                        "missingValueBehavior": "Ignore",
                    },
                }
                for index in range(4)
            ]
        }
    )

    preview = server.call_tool(
        "preview_template_batch",
        {"connectionId": "session", "documentId": template_id, "requestJson": request},
    )
    if not preview.get("isValid"):
        raise SmokeError(f"preview_template_batch refused a valid batch: {json.dumps(preview)[:400]}")

    committed = server.call_tool(
        "populate_template_batch",
        {
            "connectionId": "session",
            "documentId": template_id,
            "requestJson": request,
            "expectedTokenJson": json.dumps(preview["token"]),
        },
    )
    if not committed.get("committed"):
        raise SmokeError(f"the token-bound batch commit was refused: {json.dumps(committed)[:400]}")

    return {"output_bytes": 0, "work_units": len(committed.get("items", []))}


SCENARIOS = {
    "mcp-document-workflow": (
        "Create, inspect, preview, apply, find and export one document over stdio.",
        document_workflow,
    ),
    "mcp-template-batch": (
        "Import a template, discover it, preflight a four-item batch and commit it bound to "
        "that preview, over stdio.",
        template_workflow,
    ),
}


def measure(server: StdioServer, repetitions: int, warmup: int) -> list[dict[str, Any]]:
    scenarios: list[dict[str, Any]] = []
    for name, (description, workflow) in SCENARIOS.items():
        iterations: list[dict[str, Any]] = []
        for index in range(warmup + repetitions):
            server.reset_counters()
            started = time.perf_counter()
            failure = None
            detail: dict[str, Any] = {"output_bytes": 0, "work_units": 0}
            try:
                detail = workflow(server, str(index))
            except Exception as error:  # noqa: BLE001 - recorded, not swallowed
                failure = f"{type(error).__name__}: {error}"
            elapsed = (time.perf_counter() - started) * 1000.0
            counters = server.protocol_counters

            if index < warmup:
                continue
            iterations.append(
                {
                    "Index": len(iterations),
                    "ElapsedMilliseconds": elapsed,
                    # The MCP path is driven from Python, so managed allocation is not
                    # observable here. It is recorded as zero rather than guessed at, and the
                    # reporter's allocation gate is inert for these scenarios by design.
                    "AllocatedBytes": 0,
                    "OutputBytes": detail["output_bytes"],
                    "WorkUnits": detail["work_units"],
                    "Succeeded": failure is None,
                    "Failure": failure,
                    "Protocol": counters,
                }
            )

        scenarios.append(
            {
                "Name": name,
                "Description": description,
                "Inputs": ["packaged officeagent-mcp"],
                "WarmupIterations": warmup,
                "Iterations": iterations,
                "FailedIterations": sum(1 for i in iterations if not i["Succeeded"]),
                "ProcessPeakWorkingSetBytes": 0,
            }
        )
    return scenarios


def benchmark(artifacts: Path, version: str, dotnet: str, repetitions: int, warmup: int) -> dict[str, Any]:
    required = [artifacts / f"OfficeAgent.{name}.{version}.nupkg" for name in
                ("Abstractions", "Core", "Word", "PowerPoint", "Excel", "AgentFramework",
                 "SharePoint", "Mcp")]
    missing = [path.name for path in required if not path.is_file()]
    if missing:
        raise SmokeError(f"missing packages: {', '.join(missing)}")

    with tempfile.TemporaryDirectory(prefix="officeagent-mcp-benchmark-") as temporary:
        fixture = Path(temporary)
        cache = fixture / "package-cache"
        cache.mkdir()
        tool_root = fixture / "tools"
        tool_root.mkdir()
        config = fixture / "NuGet.Config"
        write_nuget_config(config, artifacts)
        env = isolated_env(cache, config)

        run(
            [dotnet, "tool", "install", "OfficeAgent.Mcp", "--version", version,
             "--tool-path", str(tool_root), "--configfile", str(config)],
            fixture,
            env,
        )
        executable = tool_root / ("officeagent-mcp.exe" if os.name == "nt" else "officeagent-mcp")
        if not executable.is_file():
            raise SmokeError(f"installed tool executable not found at {executable}")

        workdir = fixture / "empty"
        workdir.mkdir()
        with StdioServer([str(executable), "--stdio"], workdir, env) as server:
            info = initialize(server)
            reported = info.get("serverInfo", {}).get("version", "")
            if reported.split("+")[0] != version:
                raise SmokeError(
                    f"server reports version {reported!r}, expected the candidate {version!r}")
            scenarios = measure(server, repetitions, warmup)

    return {
        "SchemaVersion": SCHEMA_VERSION,
        "Environment": {
            "Commit": git("rev-parse", "HEAD"),
            "CommitIsClean": not git("status", "--porcelain"),
            "Corpus": "packaged officeagent-mcp over stdio",
            "RepetitionsPerScenario": repetitions,
            "WarmupIterations": warmup,
            "Runtime": f"packaged tool {version}",
            "OperatingSystem": f"{platform.system()} {platform.release()}",
            "Architecture": platform.machine(),
            "ProcessorCount": os.cpu_count() or 0,
            "ServerGarbageCollection": False,
            "Configuration": "Release",
            "MeasurementMethod":
                "perf_counter wall clock around one stdio workflow against an installed tool. "
                "Protocol frames and bytes are counted at the stdio boundary and are exact. "
                "Managed allocation is not observable from this side and is recorded as zero.",
            "StartedUtc": datetime.now(timezone.utc).isoformat(),
        },
        "Scenarios": scenarios,
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Benchmark the installed MCP path.")
    parser.add_argument("--artifacts", required=True, type=Path, help="Directory of packed .nupkg files.")
    parser.add_argument("--output", required=True, type=Path, help="Raw results JSON path.")
    parser.add_argument("--version", help="Candidate version; defaults to the repository version.")
    parser.add_argument("--dotnet", default="dotnet", help="dotnet executable to use.")
    parser.add_argument("--repetitions", type=int, default=10)
    parser.add_argument("--warmup", type=int, default=2)
    args = parser.parse_args(argv)

    try:
        version = args.version or repository_version()
        # The local feed is written into a NuGet.Config inside a temporary fixture directory,
        # so a relative path here would resolve against that directory and find nothing.
        results = benchmark(
            args.artifacts.resolve(), version, args.dotnet, args.repetitions, args.warmup)
    except (SmokeError, OSError, subprocess.CalledProcessError) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2

    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(results, indent=2), encoding="utf-8", newline="\n")

    failed = sum(s["FailedIterations"] for s in results["Scenarios"])
    for scenario in results["Scenarios"]:
        first = scenario["Iterations"][0]["Protocol"] if scenario["Iterations"] else {}
        print(f"{scenario['Name']}: calls={first.get('ToolCalls', 0)} "
              f"frames={first.get('FramesSent', 0)}/{first.get('FramesReceived', 0)} "
              f"bytes={first.get('BytesSent', 0)}/{first.get('BytesReceived', 0)}")
    print(f"mcp-benchmark-scenarios={len(results['Scenarios'])} failed-iterations={failed}")
    print(f"raw-results={args.output}")
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
