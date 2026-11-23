# Sample configuration

[`word-and-powerpoint.json`](word-and-powerpoint.json) configures the MCP server for one folder
of Word documents and PowerPoint decks. It turns on `create_document`, allows `.docx` and
`.pptx`, and makes `Direct` the default change mode, because PowerPoint cannot record tracked
changes.

Set `RootPath` to the folder the server may read and write, then start the server with it:

```bash
officeagent-mcp --stdio --config samples/config/word-and-powerpoint.json
```

Environment variables still override the file. The README's
[Configure broader workflows](../../README.md#configure-broader-workflows) section and the MCP
server guide's [configuration file](../../docs/mcp-server.md#a-configuration-file) section walk
through it.
