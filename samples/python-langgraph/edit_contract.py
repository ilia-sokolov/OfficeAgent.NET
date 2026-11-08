"""Edit a Word contract with a LangGraph agent through OfficeAgent's MCP server.

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

from langchain_core.messages import AIMessage, HumanMessage, SystemMessage, ToolMessage
from langchain_mcp_adapters.sessions import create_session
from langchain_mcp_adapters.tools import load_mcp_tools
from langchain_openai import ChatOpenAI
from langgraph.graph import START, MessagesState, StateGraph
from langgraph.prebuilt import ToolNode, tools_condition

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
    # The result can be content blocks, JSON text, or JSON text encoded again as a string.
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


def verify(output: Path) -> None:
    """Check the saved package for the tracked change, without trusting the agent's account."""
    xml = zipfile.ZipFile(output).read("word/document.xml").decode("utf-8")
    inserted = "".join(re.findall(r"<w:ins\b.*?</w:ins>", xml, re.S))
    deleted = "".join(re.findall(r"<w:del\b.*?</w:del>", xml, re.S))
    current = "".join(re.findall(r"<w:t(?:\s[^>]*)?>([^<]*)<", re.sub(r"<w:del\b.*?</w:del>", "", xml, flags=re.S)))
    print(f"Tracked insertion of 'twenty-four months': {'twenty-four months' in inserted}")
    print(f"Tracked deletion of 'twelve months':       {'twelve months' in deleted}")
    print(f"Clause 4 still says twelve months:         {'in the twelve months preceding' in current}")


def build_graph(model: Any, tools: list, instructions: str):
    """A two-node agent: the model decides, the tool node calls OfficeAgent, until no tool call."""
    bound = model.bind_tools(tools)

    async def call_model(state: MessagesState) -> dict:
        return {"messages": [await bound.ainvoke([SystemMessage(instructions), *state["messages"]])]}

    graph = StateGraph(MessagesState)
    graph.add_node("agent", call_model)
    graph.add_node("tools", ToolNode(tools))
    graph.add_edge(START, "agent")
    graph.add_conditional_edges("agent", tools_condition)
    graph.add_edge("tools", "agent")
    return graph.compile()


async def run_agent(graph: Any) -> tuple[dict[str, Any] | None, str]:
    """Stream the graph, printing each tool call and keeping the last committed write's report."""
    report: dict[str, Any] | None = None
    final = ""
    async for update in graph.astream(
        {"messages": [HumanMessage(TASK)]}, stream_mode="updates", config={"recursion_limit": 40}
    ):
        for node in update.values():
            for message in node.get("messages", []):
                if isinstance(message, AIMessage):
                    for call in message.tool_calls:
                        print(f"-> {call['name']} {short(json.dumps(call['args']))}")
                    if not message.tool_calls:
                        final = message.text if hasattr(message, "text") else str(message.content)
                elif isinstance(message, ToolMessage):
                    payload = tool_payload(message.content)
                    if payload is None:
                        if message.status == "error":
                            print(f"   {short(str(message.content))}")
                        continue
                    if "committed" in payload:
                        print(
                            f"   isValid={payload.get('isValid')} committed={payload.get('committed')} "
                            f"writeOutcome={payload.get('writeOutcome')} "
                            f"outputName={payload.get('outputName')}"
                        )
                        for error in payload.get("errors") or []:
                            print(f"   error: {short(json.dumps(error))}")
                        if payload.get("committed"):
                            report = payload
    return report, final


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
    connection = {
        "transport": "stdio",
        "command": command,
        "args": command_args,
        "env": server_environment(workdir),
    }

    # One session for the whole run, so the server starts once and keeps its registrations.
    async with create_session(connection) as session:
        init = await session.initialize()
        tools = await load_mcp_tools(session, server_name="officeagent")

        if args.list_tools:
            print(f"Server: {init.serverInfo.name} {init.serverInfo.version}")
            print(f"Server instructions: {len(init.instructions or '')} characters")
            for tool in tools:
                print(f"  {tool.name}")
            result = await session.call_tool("list_connections", {})
            print(f"list_connections: {short(result.content[0].text, 300)}")
            return 0

        if not os.environ.get("OPENAI_API_KEY"):
            print("Set OPENAI_API_KEY to run the agent, or pass --list-tools.", file=sys.stderr)
            return 2

        # The server sends its own guidance - addressing, the inspect/find/preview/apply loop,
        # tracked changes, error codes - when the session starts. Pass it on as the instructions.
        instructions = (
            f"{init.instructions or ''}\n\nThe user's documents are in connection '{CONNECTION_ID}'."
        )
        model = ChatOpenAI(model=os.environ.get("OPENAI_MODEL", DEFAULT_MODEL))
        graph = build_graph(model, tools, instructions)

        print(f"Task: {TASK}\n")
        report, final = await run_agent(graph)
        print(f"\nAgent: {final}\n")

    if report is None:
        print("No write was committed.", file=sys.stderr)
        return 1

    print("Change report:")
    print(json.dumps(report.get("changes"), indent=2))
    output = workdir / report["outputName"]
    print(f"\nSource: {source}")
    print(f"Output: {output}")
    verify(output)
    return 0


if __name__ == "__main__":
    sys.exit(asyncio.run(main()))
