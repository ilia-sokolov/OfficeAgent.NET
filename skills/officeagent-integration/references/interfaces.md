# Interface configuration

Keep every OfficeAgent package on `0.9.0` for the released kit. Use the direct client as the application boundary, then add one optional model-facing interface if the host needs it.

## Direct .NET

For Word, install:

```bash
dotnet add package OfficeAgent.Core --version 0.9.0
dotnet add package OfficeAgent.Word --version 0.9.0
```

Create `OfficeAgentClient` with a `WordModule`. Use `StreamHandle` for caller-owned bytes, or configure a bounded provider when documents need opaque ids and saved versions. The installed [recipes](recipes.md) exercise both forms before any agent interface is added.

## Microsoft Agent Framework

Install the adapter and the required format module:

```bash
dotnet add package OfficeAgent.AgentFramework --version 0.9.0
dotnet add package OfficeAgent.Word --version 0.9.0
dotnet add package Microsoft.Extensions.DependencyInjection --version 8.0.0
```

Build the same client through dependency injection, then expose its tools:

```csharp
var services = new ServiceCollection()
    .AddWordFormat()
    .AddFileSystemDocumentProvider("workspace", storageRoot)
    .AddOfficeAgent()
    .BuildServiceProvider();

var client = services.GetRequiredService<OfficeAgentClient>();
var tools = new OfficeAgentTools(client).AsAIFunctions();
var instructions = OfficeAgentTools.SystemPromptGuidance;
```

Registration and creation tools are off by default. Enable them only when the host grants the corresponding storage capability. The host must deliver committed output bytes through its own attachment or download channel.

## MCP

Install the server separately from the skill:

```bash
dotnet tool install --global OfficeAgent.Mcp --version 0.9.0
officeagent-mcp --stdio
```

A local MCP client configuration uses the executable as a stdio child process:

```json
{
  "mcpServers": {
    "officeagent": {
      "command": "officeagent-mcp",
      "args": ["--stdio"]
    }
  }
}
```

Configure allowed filesystem roots or SharePoint connections in the server, not in model-supplied plan JSON. Restart the client and confirm that OfficeAgent tools are listed. Skill discovery and MCP connection discovery are separate checks.

Use the versioned [agent integration guide](https://github.com/ilia-sokolov/OfficeAgent.NET/blob/v0.9.0/docs/agent-integration.md) and [MCP server reference](https://github.com/ilia-sokolov/OfficeAgent.NET/blob/v0.9.0/docs/mcp-server.md) for the complete option and tool contracts.
