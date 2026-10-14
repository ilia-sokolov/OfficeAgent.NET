# officeagent-mcp

The [OfficeAgent.NET](https://github.com/ilia-sokolov/OfficeAgent.NET) MCP server, installable
with pip or uv. It lets an AI assistant inspect and edit existing Word, PowerPoint and Excel
files: the assistant previews a typed edit plan, and the server applies it all or nothing, with
Word edits written as real tracked changes you accept or reject in Word.

The wheel contains the self-contained server executable from the matching GitHub release, so
nothing else needs to be installed: no .NET, no Office.

## Use it from an MCP client

```json
{
  "mcpServers": {
    "officeagent": {
      "command": "uvx",
      "args": ["officeagent-mcp", "--stdio"],
      "env": {
        "OfficeAgent__FileSystemConnections__0__ConnectionId": "documents",
        "OfficeAgent__FileSystemConnections__0__RootPath": "/absolute/path/to/documents",
        "OfficeAgent__FileSystemConnections__0__AllowedExtensions__0": ".docx",
        "OfficeAgent__FileSystemConnections__0__AllowedExtensions__1": ".pptx",
        "OfficeAgent__FileSystemConnections__0__AllowedExtensions__2": ".xlsx"
      }
    }
  }
}
```

The server reads and writes only inside `RootPath`. With no connection configured, it keeps
documents in memory for the life of the process.

Or install it as a command: `pip install officeagent-mcp`, then run `officeagent-mcp --stdio`.

## Platforms

Windows (x64, Arm64), macOS 15 or later (Intel, Apple silicon), and Linux with glibc 2.27 or
later (x64, Arm64). Alpine and other musl-based distributions are not supported.

## More

- [MCP server guide: settings, tools, hosting](https://github.com/ilia-sokolov/OfficeAgent.NET/blob/main/docs/mcp-server.md)
- [Privacy policy](https://github.com/ilia-sokolov/OfficeAgent.NET/blob/main/PRIVACY.md)
- [Changelog](https://github.com/ilia-sokolov/OfficeAgent.NET/blob/main/CHANGELOG.md)

MIT licensed. Each wheel is built from a release archive whose checksum and GitHub build
attestation are verified first, and PyPI carries the wheel's own publish attestation.
