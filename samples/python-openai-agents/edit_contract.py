"""Edit a Word contract with the OpenAI Agents SDK through OfficeAgent's MCP server.

The agent changes the initial term in clause 5 of services-agreement.docx from twelve months
to twenty-four months as a tracked change, and saves the result beside the input.

    python edit_contract.py               # run the agent (needs OPENAI_API_KEY)
    python edit_contract.py --list-tools  # connect to the server and list its tools, no model
"""

from __future__ import annotations

import argparse
import asyncio
import json
import os
import re
import shlex
import shutil
import sys
import zipfile
from pathlib import Path
from typing import Any

from agents import Agent, RunContextWrapper, RunHooks, Runner, Tool
from agents.mcp import MCPServerStdio

HERE = Path(__file__).resolve().parent
DEFAULT_DOCUMENT = HERE.parent / "documents" / "services-agreement.docx"
CONNECTION_ID = "contracts"
DEFAULT_MODEL = "gpt-5.6-luna"

TASK = (
    "In services-agreement.docx, change the initial term in clause 5 from twelve months to "
    "twenty-four months as a tracked change. Clause 4 also says twelve months; leave it as it "
    "is. Save the result as a new version beside the original."
)


def server_command(text: str) -> tuple[str, list[str]]:
    """Split OFFICEAGENT_MCP_COMMAND into an executable and its arguments."""
    parts = shlex.split(text, posix=os.name != "nt")
    parts = [part.strip('"') for part in parts]
    if not parts:
        raise SystemExit("The server command is empty.")
    return parts[0], parts[1:]


def server_environment(root: Path) -> dict[str, str]:
    """One filesystem connection, configured the way docs/mcp-server.md describes."""
    env = {
        "OfficeAgent__FileSystemConnections__0__ConnectionId": CONNECTION_ID,
        "OfficeAgent__FileSystemConnections__0__RootPath": str(root),
        # The server logs to stderr; keep it to warnings so the tool calls stay readable.
        "Logging__LogLevel__Default": "Warning",
        "Logging__LogLevel__ModelContextProtocol": "Error",
    }
    # The MCP client passes only a minimal environment to the server; keep the .NET ones.
    env.update({key: value for key, value in os.environ.items() if key.startswith("DOTNET_")})
    return env


def prepare_workdir(document: Path, workdir: Path) -> Path:
    """Copy a fresh contract into the connection root and clear earlier outputs."""
    workdir.mkdir(parents=True, exist_ok=True)
    for old in workdir.glob(f"{document.stem}.v*{document.suffix}"):
        old.unlink()
    target = workdir / document.name
    shutil.copyfile(document, target)
    return target


def short(text: str, limit: int = 160) -> str:
    text = " ".join(text.split())
    return text if len(text) <= limit else text[: limit - 3] + "..."


def tool_payload(result: Any) -> dict[str, Any] | None:
    """Read the JSON object an OfficeAgent tool returned, whatever wrapper it arrived in."""
    value: Any = result
    # The result can be a text-content object, JSON text, or JSON text encoded again as a string.
    for _ in range(3):
        if not isinstance(value, str):
            break
        try:
            value = json.loads(value)
        except json.JSONDecodeError:
            return None
    if isinstance(value, dict) and value.get("type") == "text" and "text" in value:
        return tool_payload(value["text"])
    if isinstance(value, list) and value and isinstance(value[0], dict) and "text" in value[0]:
        return tool_payload(value[0]["text"])
    return value if isinstance(value, dict) else None


class PrintToolCalls(RunHooks):
    """Prints every tool call and keeps the last committed write's change report."""

    def __init__(self) -> None:
        self.report: dict[str, Any] | None = None

    async def on_tool_start(self, context: RunContextWrapper, agent: Agent, tool: Tool) -> None:
        arguments = getattr(context, "tool_arguments", "") or ""
        print(f"-> {tool.name} {short(arguments)}")

    async def on_tool_end(
        self, context: RunContextWrapper, agent: Agent, tool: Tool, result: Any
    ) -> None:
        payload = tool_payload(result)
        if payload is None:
            print(f"   {short(str(result))}")
            return
        if "committed" in payload:
            print(
                f"   isValid={payload.get('isValid')} committed={payload.get('committed')} "
                f"writeOutcome={payload.get('writeOutcome')} outputName={payload.get('outputName')}"
            )
            for error in payload.get("errors") or []:
                print(f"   error: {short(json.dumps(error))}")
            if payload.get("committed"):
                self.report = payload


def verify(output: Path) -> None:
    """Check the saved package for the tracked change, without trusting the agent's account."""
    xml = zipfile.ZipFile(output).read("word/document.xml").decode("utf-8")
    inserted = "".join(re.findall(r"<w:ins\b.*?</w:ins>", xml, re.S))
    deleted = "".join(re.findall(r"<w:del\b.*?</w:del>", xml, re.S))
    current = "".join(re.findall(r"<w:t(?:\s[^>]*)?>([^<]*)<", re.sub(r"<w:del\b.*?</w:del>", "", xml, flags=re.S)))
    print(f"Tracked insertion of 'twenty-four months': {'twenty-four months' in inserted}")
    print(f"Tracked deletion of 'twelve months':       {'twelve months' in deleted}")
    print(f"Clause 4 still says twelve months:         {'in the twelve months preceding' in current}")


async def list_tools(server: MCPServerStdio) -> None:
    init = server.server_initialize_result
    info = init.server_info if init else None
    instructions = (init.instructions or "") if init else ""
    print(f"Server: {info.name if info else '?'} {info.version if info else ''}")
    print(f"Server instructions: {len(instructions)} characters")
    for tool in await server.list_tools():
        print(f"  {tool.name}")
    result = await server.call_tool("list_connections", {})
    print(f"list_connections: {short(result.content[0].text, 300)}")


async def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--document", type=Path, default=DEFAULT_DOCUMENT)
    parser.add_argument("--workdir", type=Path, default=HERE / "work")
    parser.add_argument(
        "--server-command",
        default=os.environ.get("OFFICEAGENT_MCP_COMMAND", "officeagent-mcp --stdio"),
        help="command that starts the server over stdio (env: OFFICEAGENT_MCP_COMMAND)",
    )
    parser.add_argument("--list-tools", action="store_true", help="list tools and exit; no model")
    args = parser.parse_args()

    workdir = args.workdir.resolve()
    source = prepare_workdir(args.document.resolve(), workdir)
    command, command_args = server_command(args.server_command)

    async with MCPServerStdio(
        name="officeagent",
        params={"command": command, "args": command_args, "env": server_environment(workdir)},
        cache_tools_list=True,
        client_session_timeout_seconds=120,
    ) as server:
        if args.list_tools:
            await list_tools(server)
            return 0

        if not os.environ.get("OPENAI_API_KEY"):
            print("Set OPENAI_API_KEY to run the agent, or pass --list-tools.", file=sys.stderr)
            return 2

        init = server.server_initialize_result
        # The server sends its own guidance - addressing, the inspect/find/preview/apply loop,
        # tracked changes, error codes - when the session starts. Pass it on as the instructions.
        instructions = (init.instructions or "") if init else ""
        agent = Agent(
            name="Contract editor",
            instructions=(
                f"{instructions}\n\nThe user's documents are in connection '{CONNECTION_ID}'."
            ),
            mcp_servers=[server],
            model=os.environ.get("OPENAI_MODEL", DEFAULT_MODEL),
        )

        hooks = PrintToolCalls()
        print(f"Task: {TASK}\n")
        result = await Runner.run(agent, TASK, hooks=hooks, max_turns=20)
        print(f"\nAgent: {result.final_output}\n")

    if hooks.report is None:
        print("No write was committed.", file=sys.stderr)
        return 1

    print("Change report:")
    print(json.dumps(hooks.report.get("changes"), indent=2))
    output = workdir / hooks.report["outputName"]
    print(f"\nSource: {source}")
    print(f"Output: {output}")
    verify(output)
    return 0


if __name__ == "__main__":
    sys.exit(asyncio.run(main()))
