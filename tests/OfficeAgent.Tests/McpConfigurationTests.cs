using Microsoft.Extensions.Configuration;
using OfficeAgent.Mcp;

namespace OfficeAgent.Tests;

/// <summary>
/// Connections can be written as a file instead of as indexed environment variables. The
/// file has to sit below the environment so that existing deployments keep overriding it,
/// and a connection it cannot read has to say so: the binder's own behaviour is to drop
/// the entry, which arrives much later as "no connections configured".
/// </summary>
public class McpConfigurationTests
{
    [Fact]
    public void Published_mixed_document_configuration_binds_through_the_real_loader()
    {
        var path = RepositoryFile("samples", "config", "word-and-powerpoint.json");

        var options = OfficeAgentConfiguration.Bind(Configuration(path));

        Assert.True(options.AllowCreation);
        var connection = Assert.Single(options.FileSystemConnections);
        Assert.Equal("documents", connection.ConnectionId);
        Assert.Equal(new[] { ".docx", ".pptx" }, connection.AllowedExtensions);
        Assert.Equal("Direct", connection.DefaultChangeMode);
    }

    [Fact]
    public void A_configuration_file_supplies_connections_as_a_list()
    {
        using var root = new TemporaryFile("""
            {
              "OfficeAgent": {
                "AllowCreation": true,
                "FileSystemConnections": [
                  {
                    "ConnectionId": "documents",
                    "RootPath": "C:\\docs",
                    "AllowedExtensions": [ ".docx", ".pptx" ]
                  }
                ]
              }
            }
            """);

        var options = OfficeAgentConfiguration.Bind(Configuration(root.Path));

        Assert.True(options.AllowCreation);
        var connection = Assert.Single(options.FileSystemConnections);
        Assert.Equal("documents", connection.ConnectionId);
        Assert.Equal(new[] { ".docx", ".pptx" }, connection.AllowedExtensions);
    }

    [Fact]
    public void An_environment_variable_still_overrides_the_file()
    {
        using var root = new TemporaryFile("""
            {
              "OfficeAgent": {
                "FileSystemConnections": [ { "ConnectionId": "documents", "RootPath": "/from-file" } ]
              }
            }
            """);

        const string variable = "OfficeAgent__FileSystemConnections__0__RootPath";
        Environment.SetEnvironmentVariable(variable, "/from-environment");
        try
        {
            var configuration = new ConfigurationBuilder();
            configuration.AddEnvironmentVariables();
            OfficeAgentConfiguration.AddFile(configuration, root.Path);

            var options = OfficeAgentConfiguration.Bind(configuration.Build());

            // The environment wins, and the connection is still the file's - the two
            // sources merge into one entry rather than producing two.
            var connection = Assert.Single(options.FileSystemConnections);
            Assert.Equal("documents", connection.ConnectionId);
            Assert.Equal("/from-environment", connection.RootPath);
        }
        finally
        {
            Environment.SetEnvironmentVariable(variable, null);
        }
    }

    [Fact]
    public void A_connection_that_cannot_be_read_names_the_value_instead_of_disappearing()
    {
        using var root = new TemporaryFile("""
            {
              "OfficeAgent": {
                "FileSystemConnections": [
                  { "ConnectionId": "documents", "RootPath": "/docs", "MaximumBytes": "lots" }
                ]
              }
            }
            """);

        var error = Assert.Throws<InvalidOperationException>(
            () => OfficeAgentConfiguration.Bind(Configuration(root.Path)));

        Assert.Contains("MaximumBytes", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void The_config_option_survives_a_valueless_switch_in_front_of_it()
    {
        // The command-line configuration provider reads --stdio as the key of whatever
        // follows it, so --config has to be taken from the arguments directly.
        var path = OfficeAgentConfiguration.ResolvePath(
            new[] { "--stdio", "--config", "officeagent.json" }, _ => null);

        Assert.Equal(Path.GetFullPath("officeagent.json"), path);
    }

    [Fact]
    public void The_config_option_accepts_the_inline_form()
    {
        var path = OfficeAgentConfiguration.ResolvePath(
            new[] { "--config=officeagent.json", "--stdio" }, _ => null);

        Assert.Equal(Path.GetFullPath("officeagent.json"), path);
    }

    [Fact]
    public void A_config_option_without_a_path_is_reported()
    {
        Assert.Throws<InvalidOperationException>(
            () => OfficeAgentConfiguration.ResolvePath(new[] { "--config" }, _ => null));

        Assert.Throws<InvalidOperationException>(
            () => OfficeAgentConfiguration.ResolvePath(new[] { "--config", "--stdio" }, _ => null));
    }

    [Fact]
    public void The_environment_variable_names_a_file_when_no_option_is_given()
    {
        var path = OfficeAgentConfiguration.ResolvePath(
            Array.Empty<string>(),
            name => name == OfficeAgentConfiguration.PathVariable ? "from-environment.json" : null);

        Assert.Equal(Path.GetFullPath("from-environment.json"), path);
    }

    [Fact]
    public void A_named_file_that_does_not_exist_is_reported_rather_than_ignored()
    {
        var missing = Path.Combine(Path.GetTempPath(), $"officeagent-{Guid.NewGuid():N}.json");

        // Resolution returns it so the mistake surfaces, rather than starting a server
        // without the configuration the operator believes they supplied.
        Assert.Equal(missing, OfficeAgentConfiguration.ResolvePath(new[] { "--config", missing }, _ => null));

        var configuration = new ConfigurationBuilder();
        OfficeAgentConfiguration.AddFile(configuration, missing);
        Assert.Throws<FileNotFoundException>(() => configuration.Build());
    }

    [Fact]
    public void Declaring_the_allowed_extensions_replaces_the_default_rather_than_adding_to_it()
    {
        // The binder adds to a list that already holds .docx, so without this a connection
        // for decks alone cannot be expressed and .docx arrives twice.
        using var root = new TemporaryFile("""
            {
              "OfficeAgent": {
                "FileSystemConnections": [
                  { "ConnectionId": "decks", "RootPath": "/decks", "AllowedExtensions": [ ".pptx" ] }
                ]
              }
            }
            """);

        var options = OfficeAgentConfiguration.Bind(Configuration(root.Path));

        var connection = Assert.Single(options.FileSystemConnections);
        Assert.Equal(new[] { ".pptx" }, connection.AllowedExtensions);
    }

    [Fact]
    public void The_working_directory_is_never_probed_for_configuration()
    {
        // A connection root is a trust boundary and a stdio server starts in whatever
        // directory its client happened to be in, so the current directory is not a place
        // to pick configuration up from.
        var path = OfficeAgentConfiguration.ResolvePath(Array.Empty<string>(), _ => null);

        if (path is not null)
            Assert.NotEqual(Path.GetFullPath(Environment.CurrentDirectory), Path.GetDirectoryName(path));
    }

    private static IConfiguration Configuration(string path)
    {
        var configuration = new ConfigurationBuilder();
        OfficeAgentConfiguration.AddFile(configuration, path);
        return configuration.Build();
    }

    private static string RepositoryFile(params string[] parts)
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "OfficeAgent.NET.sln")))
            directory = directory.Parent;

        Assert.NotNull(directory);
        return Path.Combine(new[] { directory!.FullName }.Concat(parts).ToArray());
    }

    private sealed class TemporaryFile : IDisposable
    {
        public TemporaryFile(string content)
        {
            Path = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(), $"officeagent-config-{Guid.NewGuid():N}.json");
            File.WriteAllText(Path, content);
        }

        public string Path { get; }

        public void Dispose()
        {
            if (File.Exists(Path)) File.Delete(Path);
        }
    }
}
