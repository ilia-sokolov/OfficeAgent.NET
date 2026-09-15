using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
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
        Assert.Equal(12, names.Length);
        Assert.Contains("open_document", names);
        Assert.Contains("edit_document", names);
        Assert.Contains("inspect_document", names);
        Assert.Contains("find_in_document", names);
        Assert.Contains("preview_plan", names);
        Assert.Contains("apply_plan", names);
        Assert.DoesNotContain("populate_template_batch", names);
        Assert.Contains("compare_documents", names);
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
        Assert.Equal(15, names.Length);
        Assert.Contains("create_document", names);
        Assert.Contains("populate_template_batch", names);
    }

    [Fact]
    public void Toolset_omits_registration_tools_when_disallowed()
    {
        using var root = new TemporaryRoot();
        var options = OptionsFor(root);
        options.AllowRegistration = false;

        var names = OfficeAgentMcpServer.BuildToolset(options).Select(t => t.ProtocolTool.Name).ToArray();

        Assert.Equal(7, names.Length);
        Assert.DoesNotContain("populate_template_batch", names);
        Assert.Contains("compare_documents", names);
        Assert.DoesNotContain("register_document", names);
        Assert.DoesNotContain("create_document", names);
        Assert.DoesNotContain("remove_document", names);
        Assert.DoesNotContain("list_connections", names);

        // The composites take a source, so they belong to the registration opt-in too.
        Assert.DoesNotContain("open_document", names);
        Assert.DoesNotContain("edit_document", names);

        options.AllowCreation = true;
        names = OfficeAgentMcpServer.BuildToolset(options).Select(t => t.ProtocolTool.Name).ToArray();
        Assert.Equal(11, names.Length);
        Assert.Contains("create_document", names);
        Assert.Contains("populate_template_batch", names);
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
    public void With_nothing_configured_the_server_starts_on_a_session_connection()
    {
        // Installing the server and configuring nothing is the commonest first contact with
        // it - from the MCP registry, where the filesystem settings are optional. Exiting
        // there taught the reader nothing; a session connection needs no storage, no
        // credentials and no filesystem reach, so there is something useful to offer.
        var options = new OfficeAgentMcpOptions();
        var names = OfficeAgentMcpServer.BuildToolset(options).Select(t => t.ProtocolTool.Name).ToArray();

        Assert.Contains("inspect_document", names);
        Assert.Contains("apply_plan", names);
        Assert.Contains("import_document_content", names);
        Assert.Contains("export_document_content", names);

        // Creation comes with it: the session starts empty, so without it the agent would
        // hold a connection it can do nothing with until something is handed to it.
        Assert.Contains("create_document", names);

        var connections = OfficeAgentMcpServer.ConnectionsPayload(options);
        Assert.Contains($"\"connectionId\":\"{OfficeAgentMcpServer.DefaultSessionConnectionId}\"", connections);
        Assert.Contains("\"provider\":\"session\"", connections);
    }

    [Fact]
    public void Starting_with_nothing_configured_says_so_on_stderr()
    {
        // Quietly doing something other than what the operator believes they configured is
        // the failure this has to avoid - a typo in a connection variable must not look
        // like a working server.
        var notice = OfficeAgentMcpServer.StartupNotice(new OfficeAgentMcpOptions());

        Assert.NotNull(notice);
        Assert.Contains("No storage is configured", notice);
        Assert.Contains(OfficeAgentMcpServer.DefaultSessionConnectionId, notice);
        Assert.Contains("last only while this server runs", notice);
        Assert.Contains("OfficeAgent__FileSystemConnections__0__RootPath", notice);
    }

    [Fact]
    public void A_configured_server_is_left_alone_and_says_nothing()
    {
        using var root = new TemporaryRoot();

        // The fallback applies only when nothing at all was configured; anything the host
        // did configure is what it gets.
        Assert.Null(OfficeAgentMcpServer.StartupNotice(OptionsFor(root)));
        Assert.Null(OfficeAgentMcpServer.StartupNotice(
            new OfficeAgentMcpOptions { AllowInlineContent = true }));
        Assert.Null(OfficeAgentMcpServer.StartupNotice(
            new OfficeAgentMcpOptions { EphemeralConnectionId = "scratch" }));

        // And a session the host asked for still respects the creation opt-in.
        var explicitSession = new OfficeAgentMcpOptions { EphemeralConnectionId = "scratch" };
        Assert.DoesNotContain("create_document",
            OfficeAgentMcpServer.BuildToolset(explicitSession).Select(t => t.ProtocolTool.Name));
    }

    [Fact]
    public void Inline_content_alone_is_enough_to_start()
    {
        var names = OfficeAgentMcpServer
            .BuildToolset(new OfficeAgentMcpOptions { AllowInlineContent = true })
            .Select(t => t.ProtocolTool.Name)
            .ToArray();

        Assert.Equal(
            new[] { "create_document_content", "inspect_document_content", "edit_document_content" }.OrderBy(n => n),
            names.OrderBy(n => n));
    }

    [Fact]
    public void With_no_connection_the_connection_addressed_tools_are_not_offered()
    {
        var names = OfficeAgentMcpServer
            .BuildToolset(new OfficeAgentMcpOptions { AllowInlineContent = true, AllowCreation = true })
            .Select(t => t.ProtocolTool.Name)
            .ToArray();

        // AllowRegistration defaults to true and AllowCreation is on, but there is no
        // connectionId in existence for either to name.
        Assert.DoesNotContain("inspect_document", names);
        Assert.DoesNotContain("apply_plan", names);
        Assert.DoesNotContain("register_document", names);
        Assert.DoesNotContain("create_document", names);
        Assert.DoesNotContain("list_connections", names);
    }

    [Fact]
    public void Inline_content_is_configured_from_the_file_or_the_environment_like_any_other_setting()
    {
        var file = Path.Combine(Path.GetTempPath(), $"officeagent-inline-{Guid.NewGuid():N}.json");
        File.WriteAllText(file, """{ "OfficeAgent": { "AllowInlineContent": true } }""");

        const string variable = "OfficeAgent__AllowInlineContent";
        try
        {
            Assert.True(Bind(file, environment: null).AllowInlineContent);
            Assert.True(Bind(configFile: null, environment: "true").AllowInlineContent);

            // The environment outranks the file, which is the precedence every other
            // setting follows - it is what lets one value be overridden per deployment.
            Assert.False(Bind(file, environment: "false").AllowInlineContent);
        }
        finally
        {
            Environment.SetEnvironmentVariable(variable, null);
            File.Delete(file);
        }

        static OfficeAgentMcpOptions Bind(string? configFile, string? environment)
        {
            Environment.SetEnvironmentVariable(variable, environment);

            var configuration = new ConfigurationBuilder();
            configuration.AddEnvironmentVariables();
            if (configFile is not null) OfficeAgentConfiguration.AddFile(configuration, configFile);

            return OfficeAgentConfiguration.Bind(configuration.Build());
        }
    }

    [Fact]
    public void Inline_content_sits_alongside_a_connection_when_both_are_configured()
    {
        var options = new OfficeAgentMcpOptions { AllowInlineContent = true };
        options.FileSystemConnections.Add(new FileSystemConnectionOptions
        {
            ConnectionId = "workspace",
            RootPath = Path.GetTempPath()
        });

        var names = OfficeAgentMcpServer.BuildToolset(options).Select(t => t.ProtocolTool.Name).ToArray();

        Assert.Contains("inspect_document", names);
        Assert.Contains("edit_document_content", names);
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

    [Theory]
    [InlineData("onBehalf0f")]
    [InlineData("shared")]
    public void Unknown_sharepoint_auth_mode_fails_closed(string authMode)
    {
        var options = new OfficeAgentMcpOptions
        {
            SharePointConnections =
            {
                new SharePointConnectionOptions
                {
                    ConnectionId = "legal",
                    AuthMode = authMode,
                    TenantId = "00000000-0000-0000-0000-000000000000",
                    ClientId = "api-client-id",
                    ClientSecret = "api-secret"
                }
            }
        };

        var error = Assert.Throws<InvalidOperationException>(() =>
            OfficeAgentMcpServer.BuildToolset(options));

        Assert.Contains("legal", error.Message);
        Assert.Contains(authMode, error.Message);
        Assert.Contains("appOnly or onBehalfOf", error.Message);
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
    public void Creation_accepts_an_excel_only_connection()
    {
        using var root = new TemporaryRoot();
        var options = OptionsFor(root);
        options.AllowRegistration = false;
        options.AllowCreation = true;
        options.FileSystemConnections[0].AllowedExtensions = new List<string> { ".xlsx" };

        var names = OfficeAgentMcpServer.BuildToolset(options).Select(t => t.ProtocolTool.Name).ToArray();

        Assert.Contains("create_document", names);
        Assert.Contains("list_connections", names);
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
        Assert.True(capabilities["spreadsheets"]);
        Assert.True(capabilities["legal"]);
        Assert.False(capabilities["archive"]);
    }

    [Fact]
    public void A_connection_change_mode_reaches_the_provider_and_a_bad_one_names_itself()
    {
        using var root = new TemporaryRoot();

        var services = new ServiceCollection();
        OfficeAgentMcpServer.AddOfficeAgentMcp(services, new OfficeAgentMcpOptions
        {
            FileSystemConnections =
            {
                new FileSystemConnectionOptions
                {
                    ConnectionId = "documents",
                    RootPath = root.Path,
                    DefaultChangeMode = "Direct"
                }
            }
        });

        using var provider = services.BuildServiceProvider();
        Assert.Equal(
            ChangeMode.Direct,
            provider.GetRequiredService<OfficeAgentClient>().DefaultChangeModeFor("documents"));

        // The binder drops a collection element it cannot bind, so this field is a string
        // and is validated here. Left as an enum, a typo would silently remove the whole
        // connection and report "no connections configured".
        var error = Assert.Throws<InvalidOperationException>(() =>
            OfficeAgentMcpServer.AddOfficeAgentMcp(new ServiceCollection(), new OfficeAgentMcpOptions
            {
                FileSystemConnections =
                {
                    new FileSystemConnectionOptions
                    {
                        ConnectionId = "documents",
                        RootPath = root.Path,
                        DefaultChangeMode = "Tracke"
                    }
                }
            }));

        Assert.Contains("documents", error.Message);
        Assert.Contains("Tracked or Direct", error.Message);
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
