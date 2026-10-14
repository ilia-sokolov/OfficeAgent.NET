"""The OfficeAgent MCP server, packaged so pip and uv can install it without .NET.

The wheel carries the self-contained server executable from the matching GitHub release;
this package only finds it and starts it.
"""
import sys
from pathlib import Path

__version__ = "__VERSION__"


def executable() -> Path:
    """The path of the bundled server executable."""
    name = "officeagent-mcp.exe" if sys.platform == "win32" else "officeagent-mcp"
    return Path(__file__).resolve().parent / "bin" / name
