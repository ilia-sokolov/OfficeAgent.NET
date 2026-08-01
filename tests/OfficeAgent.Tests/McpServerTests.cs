using OfficeAgent.Mcp;
using System.Text.Json;

namespace OfficeAgent.Tests;

/// <summary>
/// MCP composition tests: configuration alone yields the full toolset, the
/// registration tools obey the host switch, and the advertised instructions match
/// what the in-process tool layer teaches.
/// </summary>
public class McpServerTests
{
    [Fact]
    public void Toolset_exposes_core_and_registration_tools_by_default()
    {
        using var root = new TemporaryRoot();
        var options = OptionsFor(root);

        var tools = OfficeAgentMcpServer.BuildToolset(options);

        var names = tools.Select(t => t.ProtocolTool.Name).ToArray();
        Assert.Equal(9, names.Length);
        Assert.Contains("open_document", names);
        Assert.Contains("edit_document", names);
        Assert.Contains("inspect_document", names);
        Assert.Contains("find_in_document", names);
        Assert.Contains("preview_plan", names);
        Assert.Contains("apply_plan", names);
        Assert.Contains("register_document", names);
        Assert.Contains("remove_document", names);
        Assert.Contains("list_connections", names);

        // Creation is the one capability an upgrade must not switch on by itself.
        Assert.DoesNotContain("create_document", names);
    }

    [Fact]
    public void Creation_is_opt_in_and_never_arrives_with_registration_alone()
    {
        using var root = new TemporaryRoot();
        var options = OptionsFor(root);

        Assert.False(options.AllowCreation);
        Assert.DoesNotContain("create_document", OfficeAgentMcpServer.InstructionsFor(options));

        options.AllowCreation = true;
        var names = OfficeAgentMcpServer.BuildToolset(options).Select(t => t.ProtocolTool.Name).ToArray();
        Assert.Equal(10, names.Length);
        Assert.Contains("create_document", names);
    }

    [Fact]
    public void Toolset_omits_registration_tools_when_disallowed()
    {
        using var root = new TemporaryRoot();
        var options = OptionsFor(root);
        options.AllowRegistration = false;

        var names = OfficeAgentMcpServer.BuildToolset(options).Select(t => t.ProtocolTool.Name).ToArray();

        Assert.Equal(4, names.Length);
        Assert.DoesNotContain("register_document", names);
        Assert.DoesNotContain("create_document", names);
        Assert.DoesNotContain("remove_document", names);
        Assert.DoesNotContain("list_connections", names);

        // The composites take a source, so they belong to the registration opt-in too.
        Assert.DoesNotContain("open_document", names);
        Assert.DoesNotContain("edit_document", names);

        options.AllowCreation = true;
        names = OfficeAgentMcpServer.BuildToolset(options).Select(t => t.ProtocolTool.Name).ToArray();
        Assert.Equal(6, names.Length);
        Assert.Contains("create_document", names);
        Assert.Contains("list_connections", names);
        Assert.DoesNotContain("register_document", names);
        Assert.DoesNotContain("remove_document", names);
        Assert.Contains("create_document", OfficeAgentMcpServer.InstructionsFor(options));
        Assert.DoesNotContain("register_document", OfficeAgentMcpServer.InstructionsFor(options));
    }

    [Fact]
    public void Instructions_follow_the_registration_switch()
    {
        using var root = new TemporaryRoot();
        var options = OptionsFor(root);

        Assert.Contains("register_document", OfficeAgentMcpServer.InstructionsFor(options));

        options.AllowRegistration = false;
        Assert.DoesNotContain("register_document", OfficeAgentMcpServer.InstructionsFor(options));
    }

    [Fact]
    public void Creation_guidance_and_tool_travel_together()
    {
        using var root = new TemporaryRoot();
        var options = OptionsFor(root);
        options.AllowCreation = true;

        Assert.Contains("create_document", OfficeAgentMcpServer.InstructionsFor(options));

        // With creation off the agent must be told nothing about a tool it will not get.
        options.AllowCreation = false;
        Assert.DoesNotContain("create_document", OfficeAgentMcpServer.InstructionsFor(options));
        var names = OfficeAgentMcpServer.BuildToolset(options).Select(t => t.ProtocolTool.Name).ToArray();
        Assert.DoesNotContain("create_document", names);
        Assert.Contains("register_document", names);
    }

    [Fact]
    public void Instructions_list_configured_connection_ids()
    {
        using var root = new TemporaryRoot();
        var options = OptionsFor(root);
        options.SharePointConnections.Add(new SharePointConnectionOptions { ConnectionId = "legal" });

        var instructions = OfficeAgentMcpServer.InstructionsFor(options);

        Assert.Contains("Configured connections", instructions);
        Assert.Contains("\"documents\" (filesystem)", instructions);
        Assert.Contains("\"legal\" (sharepoint)", instructions);
    }

    [Fact]
    public void Server_refuses_to_start_without_connections()
    {
        var ex = Assert.Throws<InvalidOperationException>(() =>
            OfficeAgentMcpServer.BuildToolset(new OfficeAgentMcpOptions()));
        Assert.Contains("at least one connection", ex.Message);
    }

    [Fact]
    public void On_behalf_of_sharepoint_connection_composes()
    {
        // Forces construction of the SharePoint provider and its OBO token provider
        // through the configured AuthMode path - validates the full wiring without a
        // tenant. Missing OBO credentials would throw here.
        var options = new OfficeAgentMcpOptions
        {
            AllowCreation = true,
            SharePointConnections =
            {
                new SharePointConnectionOptions
                {
                    ConnectionId = "legal",
                    AuthMode = "onBehalfOf",
                    TenantId = "00000000-0000-0000-0000-000000000000",
                    ClientId = "api-client-id",
                    ClientSecret = "api-secret"
                }
            }
        };

        var names = OfficeAgentMcpServer.BuildToolset(options).Select(t => t.ProtocolTool.Name).ToArray();
        Assert.Contains("apply_plan", names);
        Assert.Contains("register_document", names);
        Assert.DoesNotContain("create_document", names);
        Assert.DoesNotContain("create_document", OfficeAgentMcpServer.InstructionsFor(options));
    }

    [Fact]
    public void Sharepoint_creation_is_exposed_only_with_a_configured_destination()
    {
        var options = new OfficeAgentMcpOptions
        {
            AllowRegistration = false,
            AllowCreation = true,
            SharePointConnections =
            {
                new SharePointConnectionOptions
                {
                    ConnectionId = "legal",
                    AuthMode = "onBehalfOf",
                    TenantId = "00000000-0000-0000-0000-000000000000",
                    ClientId = "api-client-id",
                    ClientSecret = "api-secret",
                    CreationDriveId = "drive-legal",
                    CreationFolderItemId = "folder-drafts"
                }
            }
        };

        var names = OfficeAgentMcpServer.BuildToolset(options).Select(t => t.ProtocolTool.Name).ToArray();

        Assert.Contains("create_document", names);
        Assert.Contains("list_connections", names);
        Assert.DoesNotContain("register_document", names);
        Assert.Contains("configured folder", OfficeAgentMcpServer.InstructionsFor(options));
    }

    [Fact]
    public void Creation_requires_a_connection_that_allows_docx()
    {
        using var root = new TemporaryRoot();
        var options = OptionsFor(root);
        options.AllowRegistration = false;
        options.AllowCreation = true;
        options.FileSystemConnections[0].AllowedExtensions = new List<string> { ".xlsx" };

        var names = OfficeAgentMcpServer.BuildToolset(options).Select(t => t.ProtocolTool.Name).ToArray();

        Assert.DoesNotContain("create_document", names);
        Assert.DoesNotContain("list_connections", names);
    }

    [Fact]
    public void Connection_payload_reports_creation_per_connection()
    {
        using var root = new TemporaryRoot();
        var options = OptionsFor(root);
        options.AllowCreation = true;
        options.FileSystemConnections.Add(new FileSystemConnectionOptions
        {
            ConnectionId = "spreadsheets",
            RootPath = root.Path,
            AllowedExtensions = new List<string> { ".xlsx" }
        });
        options.SharePointConnections.Add(new SharePointConnectionOptions
        {
            ConnectionId = "legal",
            CreationDriveId = "drive-legal",
            CreationFolderItemId = "folder-drafts"
        });
        options.SharePointConnections.Add(new SharePointConnectionOptions
        {
            ConnectionId = "archive"
        });

        using var json = JsonDocument.Parse(OfficeAgentMcpServer.ConnectionsPayload(options));
        var capabilities = json.RootElement.EnumerateArray().ToDictionary(
            item => item.GetProperty("connectionId").GetString()!,
            item => item.GetProperty("canCreateDocuments").GetBoolean());

        Assert.True(capabilities["documents"]);
        Assert.False(capabilities["spreadsheets"]);
        Assert.True(capabilities["legal"]);
        Assert.False(capabilities["archive"]);
    }

    private static OfficeAgentMcpOptions OptionsFor(TemporaryRoot root) => new()
    {
        FileSystemConnections =
        {
            new FileSystemConnectionOptions { ConnectionId = "documents", RootPath = root.Path }
        }
    };

    private sealed class TemporaryRoot : IDisposable
    {
        public TemporaryRoot()
        {
            Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"officeagent-mcp-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Path);
        }

        public string Path { get; }

        public void Dispose()
        {
            if (Directory.Exists(Path)) Directory.Delete(Path, recursive: true);
        }
    }
}
