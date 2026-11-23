using Microsoft.Extensions.Configuration;
using OfficeAgent.Mcp;

namespace OfficeAgent.Tests;

/// <summary>
/// The <c>appsettings.json</c> the container image and the HTTP server ship with. .NET merges
/// configuration lists by position, so a connection defined there survives under every
/// connection a deployment adds through the environment.
/// </summary>
public sealed class ShippedSettingsTests
{
    [Fact]
    public void A_server_configured_by_environment_lists_only_the_environments_connections()
    {
        var configuration = new ConfigurationBuilder()
            .AddJsonFile(ShippedSettings(), optional: false)
            .AddInMemoryCollection(new Dictionary<string, string?>
            {
                ["OfficeAgent:SharePointConnections:0:ConnectionId"] = "sharepoint",
                ["OfficeAgent:SharePointConnections:0:AuthMode"] = "onBehalfOf"
            })
            .Build();

        var options = OfficeAgentConfiguration.Bind(configuration);

        Assert.Empty(options.FileSystemConnections);
        Assert.Equal("sharepoint", Assert.Single(options.SharePointConnections).ConnectionId);
        Assert.True(string.IsNullOrWhiteSpace(options.EphemeralConnectionId));
    }

    [Fact]
    public void The_shipped_settings_define_no_connection()
    {
        var options = OfficeAgentConfiguration.Bind(
            new ConfigurationBuilder().AddJsonFile(ShippedSettings(), optional: false).Build());

        Assert.Empty(options.FileSystemConnections);
        Assert.Empty(options.SharePointConnections);
        Assert.Equal("http", options.Transport);
    }

    private static string ShippedSettings()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "OfficeAgent.NET.sln")))
            directory = directory.Parent;
        return Path.Combine(directory!.FullName, "src", "OfficeAgent.Mcp", "appsettings.json");
    }
}
