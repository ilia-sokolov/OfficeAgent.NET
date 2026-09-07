using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Configuration.EnvironmentVariables;

namespace OfficeAgent.Mcp;

/// <summary>
/// Resolves and reads server configuration.
/// </summary>
/// <remarks>
/// Connections can be written as a JSON file instead of indexed environment variables.
/// The file carries the same <c>OfficeAgent</c> section shape as <c>appsettings.json</c>,
/// so a list of connections is written as a list rather than as
/// <c>FileSystemConnections__0__AllowedExtensions__1</c>. It layers below
/// <c>OfficeAgent__</c> environment variables and the command line, which leaves existing
/// deployments unchanged and keeps the environment as the place to override a single
/// value or supply a secret.
/// </remarks>
public static class OfficeAgentConfiguration
{
    /// <summary>The environment variable naming a configuration file.</summary>
    public const string PathVariable = "OFFICEAGENT_CONFIG";

    /// <summary>The command-line option naming a configuration file.</summary>
    public const string PathOption = "--config";

    /// <summary>
    /// Returns the configuration file to load, or <see langword="null"/> when none applies.
    /// </summary>
    /// <remarks>
    /// A path given on the command line or in the environment is returned whether or not it
    /// exists, so that a mistyped path reports itself instead of starting a server with
    /// configuration the operator believes they supplied. The per-user location is a probe
    /// rather than an instruction, so it is returned only when the file is really there.
    /// The working directory is deliberately not probed: a connection root is a trust
    /// boundary, and a stdio server is started in whatever directory its client happened to
    /// be in.
    /// </remarks>
    public static string? ResolvePath(
        IReadOnlyList<string> args,
        Func<string, string?>? environment = null)
    {
        if (args is null) throw new ArgumentNullException(nameof(args));
        environment ??= Environment.GetEnvironmentVariable;

        if (FromArguments(args) is { } fromArguments)
            return Path.GetFullPath(fromArguments);

        var configured = environment(PathVariable);
        if (!string.IsNullOrWhiteSpace(configured))
            return Path.GetFullPath(configured.Trim());

        var perUser = PerUserPath();
        return perUser is not null && File.Exists(perUser) ? perUser : null;
    }

    /// <summary>
    /// Adds <paramref name="path"/> to <paramref name="configuration"/> below the
    /// environment variables, so <c>OfficeAgent__</c> settings and the command line still
    /// override the file.
    /// </summary>
    public static void AddFile(IConfigurationBuilder configuration, string path)
    {
        if (configuration is null) throw new ArgumentNullException(nameof(configuration));
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("The configuration path is empty.", nameof(path));

        configuration.AddJsonFile(Path.GetFullPath(path), optional: false, reloadOnChange: false);

        // AddJsonFile appends, which would make the file outrank the environment. Move it
        // ahead of the first environment source instead; sources earlier in the list lose.
        var sources = configuration.Sources;
        var destination = FirstEnvironmentIndex(sources);
        if (destination < 0 || destination >= sources.Count - 1) return;

        var added = sources[sources.Count - 1];
        sources.RemoveAt(sources.Count - 1);
        sources.Insert(destination, added);
    }

    /// <summary>
    /// Binds the <c>OfficeAgent</c> section, failing when a connection could not be read.
    /// </summary>
    /// <remarks>
    /// The configuration binder skips a collection element it cannot bind rather than
    /// reporting it, so one bad value in one optional field removes a whole connection and
    /// surfaces later as "no connections configured" - a diagnostic pointing nowhere near
    /// the mistake. Comparing what was declared against what was bound catches that, and
    /// re-reading the offending element on its own names the value responsible.
    /// </remarks>
    public static OfficeAgentMcpOptions Bind(IConfiguration configuration)
    {
        if (configuration is null) throw new ArgumentNullException(nameof(configuration));

        var section = configuration.GetSection(OfficeAgentMcpOptions.SectionName);
        var options = section.Get<OfficeAgentMcpOptions>() ?? new OfficeAgentMcpOptions();

        var files = section.GetSection(nameof(OfficeAgentMcpOptions.FileSystemConnections));
        var sharePoint = section.GetSection(nameof(OfficeAgentMcpOptions.SharePointConnections));

        RequireEveryElementRead<FileSystemConnectionOptions>(files, options.FileSystemConnections.Count);
        RequireEveryElementRead<SharePointConnectionOptions>(sharePoint, options.SharePointConnections.Count);

        ReplaceDeclaredExtensions(files, index =>
            index < options.FileSystemConnections.Count
                ? options.FileSystemConnections[index].AllowedExtensions
                : null);
        ReplaceDeclaredExtensions(sharePoint, index =>
            index < options.SharePointConnections.Count
                ? options.SharePointConnections[index].AllowedExtensions
                : null);

        return options;
    }

    /// <summary>
    /// Replaces a connection's allowed extensions with exactly what configuration declared.
    /// </summary>
    /// <remarks>
    /// The binder adds to a list property rather than replacing it, and this one is created
    /// holding <c>.docx</c>. Configuring <c>[".docx", ".pptx"]</c> therefore yields
    /// <c>.docx</c> twice, and a connection meant to carry decks alone cannot be expressed
    /// at all, because the default cannot be taken back out. Declaring the list is taken to
    /// mean the list, so the default applies only when nothing was declared.
    /// </remarks>
    private static void ReplaceDeclaredExtensions(
        IConfigurationSection connections, Func<int, IList<string>?> target)
    {
        foreach (var element in connections.GetChildren())
        {
            if (!int.TryParse(element.Key, out var index)) continue;

            var declared = element.GetSection("AllowedExtensions").GetChildren()
                .Select(child => child.Value)
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Select(value => value!.Trim())
                .ToList();
            if (declared.Count == 0) continue;

            if (target(index) is not { } allowed) continue;
            allowed.Clear();
            foreach (var extension in declared) allowed.Add(extension);
        }
    }

    private static void RequireEveryElementRead<T>(IConfigurationSection section, int read)
    {
        var declared = section.GetChildren().ToList();
        if (declared.Count == read) return;

        foreach (var element in declared)
        {
            try
            {
                element.Get<T>();
            }
            catch (Exception error) when (error is InvalidOperationException or FormatException)
            {
                throw new InvalidOperationException(
                    $"Configuration '{element.Path}' could not be read: {error.Message} " +
                    "Correct the value; the entry is otherwise dropped and reported as a missing connection.",
                    error);
            }
        }

        throw new InvalidOperationException(
            $"Configuration '{section.Path}' declares {declared.Count} entries but only {read} could be read.");
    }

    private static string? FromArguments(IReadOnlyList<string> args)
    {
        // Read straight from the argument list rather than through the command-line
        // configuration provider: that provider treats a valueless switch such as --stdio
        // as the key of the argument that follows it, which would swallow --config.
        for (var index = 0; index < args.Count; index++)
        {
            var argument = args[index];
            if (argument is null) continue;

            if (argument.StartsWith(PathOption + "=", StringComparison.OrdinalIgnoreCase))
            {
                var inline = argument.Substring(PathOption.Length + 1).Trim();
                if (inline.Length == 0)
                    throw new InvalidOperationException($"{PathOption} was given without a path.");
                return inline;
            }

            if (!argument.Equals(PathOption, StringComparison.OrdinalIgnoreCase)) continue;

            var next = index + 1 < args.Count ? args[index + 1] : null;
            if (string.IsNullOrWhiteSpace(next) || next!.StartsWith("-", StringComparison.Ordinal))
                throw new InvalidOperationException($"{PathOption} was given without a path.");
            return next.Trim();
        }

        return null;
    }

    private static string? PerUserPath()
    {
        // ApplicationData is %APPDATA% on Windows and $XDG_CONFIG_HOME (or ~/.config)
        // elsewhere, which is the conventional per-user location on both.
        var root = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        return string.IsNullOrEmpty(root) ? null : Path.Combine(root, "OfficeAgent", "config.json");
    }

    private static int FirstEnvironmentIndex(IList<IConfigurationSource> sources)
    {
        for (var index = 0; index < sources.Count; index++)
            if (sources[index] is EnvironmentVariablesConfigurationSource) return index;
        return -1;
    }
}
