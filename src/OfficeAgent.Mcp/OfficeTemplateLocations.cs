using Microsoft.Win32;

namespace OfficeAgent.Mcp;

/// <summary>
/// What template discovery reads from the machine: the current user's registry, the special
/// folders and whether a directory exists. Tests supply their own, so discovery does not depend
/// on the machine running them.
/// </summary>
internal interface IOfficeTemplateEnvironment
{
    /// <summary>Whether the machine is Windows, where Office keeps its template folders in the registry.</summary>
    bool IsWindows { get; }

    /// <summary>Whether the machine is macOS, where Office keeps templates in its group container.</summary>
    bool IsMacOS { get; }

    /// <summary>The current user's Documents folder, or <see langword="null"/>.</summary>
    string? DocumentsFolder { get; }

    /// <summary>The current user's home folder, or <see langword="null"/>.</summary>
    string? HomeFolder { get; }

    /// <summary>
    /// A string value under <c>HKEY_CURRENT_USER</c>, with environment variables expanded, or
    /// <see langword="null"/> when it is not set.
    /// </summary>
    string? ReadUserRegistryString(string subKey, string valueName);

    /// <summary>Whether a directory exists.</summary>
    bool DirectoryExists(string path);
}

/// <summary>The real machine.</summary>
internal sealed class SystemOfficeTemplateEnvironment : IOfficeTemplateEnvironment
{
    public static readonly SystemOfficeTemplateEnvironment Instance = new();

    public bool IsWindows => OperatingSystem.IsWindows();

    public bool IsMacOS => OperatingSystem.IsMacOS();

    public string? DocumentsFolder => NonEmpty(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));

    public string? HomeFolder => NonEmpty(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile));

    public string? ReadUserRegistryString(string subKey, string valueName)
    {
        if (!OperatingSystem.IsWindows()) return null;
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(subKey);
            // REG_EXPAND_SZ comes back expanded; a REG_SZ that still carries %VARIABLES% is
            // expanded too, as Office itself does.
            return key?.GetValue(valueName) is string value && !string.IsNullOrWhiteSpace(value)
                ? Environment.ExpandEnvironmentVariables(value.Trim())
                : null;
        }
        catch (Exception ex) when (ex is System.Security.SecurityException or UnauthorizedAccessException or IOException)
        {
            return null;
        }
    }

    public bool DirectoryExists(string path) => Directory.Exists(path);

    private static string? NonEmpty(string value) => string.IsNullOrWhiteSpace(value) ? null : value;
}

/// <summary>
/// Turns the Office template folders people see in Word, PowerPoint and Excel into read-only
/// filesystem connections, so <c>find_templates</c> searches them.
/// </summary>
internal static class OfficeTemplateLocations
{
    /// <summary>The connection id of the personal templates folder.</summary>
    public const string PersonalConnectionId = "office-personal-templates";

    /// <summary>The connection id of the workgroup templates folder.</summary>
    public const string WorkgroupConnectionId = "office-workgroup-templates";

    private static readonly string[] Applications = { "Word", "PowerPoint", "Excel" };

    /// <summary>The documents and templates of the three formats; never a macro-enabled one.</summary>
    private static readonly string[] Extensions = { ".docx", ".dotx", ".pptx", ".potx", ".xlsx", ".xltx" };

    /// <summary>
    /// Adds a connection for each template folder that exists, after the configured ones. A
    /// folder that is missing, or whose id a configured connection already uses, is skipped with
    /// a line through <paramref name="log"/>; neither is an error.
    /// </summary>
    public static void AddTo(OfficeAgentMcpOptions options, IOfficeTemplateEnvironment environment, Action<string> log)
    {
        var taken = new HashSet<string>(
            options.FileSystemConnections.Select(c => c.ConnectionId)
                .Concat(options.SharePointConnections.Select(c => c.ConnectionId))
                .Concat(options.IManageConnections.Select(c => c.ConnectionId))
                .Concat(options.GoogleDriveConnections.Select(c => c.ConnectionId))
                .Concat(options.BoxConnections.Select(c => c.ConnectionId))
                .Concat(options.AzureBlobConnections.Select(c => c.ConnectionId))
                .Concat(options.AmazonS3Connections.Select(c => c.ConnectionId))
                .Append(options.EphemeralConnectionId?.Trim() ?? string.Empty),
            StringComparer.Ordinal);

        foreach (var connection in Discover(environment, log))
        {
            if (!taken.Add(connection.ConnectionId))
            {
                log($"OfficeAgent: the Office template folder was not added as '{connection.ConnectionId}' because a configured connection already uses that id.");
                continue;
            }
            options.FileSystemConnections.Add(connection);
        }
    }

    /// <summary>The template folders that exist on this machine, as read-only connections.</summary>
    public static IReadOnlyList<FileSystemConnectionOptions> Discover(IOfficeTemplateEnvironment environment, Action<string> log)
    {
        var personal = new List<string>();
        var workgroup = new List<string>();

        if (environment.IsWindows)
        {
            foreach (var application in Applications)
            {
                if (environment.ReadUserRegistryString(
                        $@"Software\Microsoft\Office\16.0\{application}\Options", "PersonalTemplates") is { } folder)
                    personal.Add(folder);
            }
            // Office's own default when no application names a folder.
            if (personal.Count == 0 && environment.DocumentsFolder is { } documents)
                personal.Add(Path.Combine(documents, "Custom Office Templates"));

            var shared =
                environment.ReadUserRegistryString(@"Software\Policies\Microsoft\Office\16.0\Common\General", "SharedTemplates")
                ?? environment.ReadUserRegistryString(@"Software\Microsoft\Office\16.0\Common\General", "SharedTemplates");
            if (shared is not null) workgroup.Add(shared);
        }
        else if (environment.IsMacOS && environment.HomeFolder is { } home)
        {
            personal.Add(Path.Combine(home, "Library", "Group Containers", "UBF8T346G9.Office", "User Content", "Templates"));
        }

        var connections = new List<FileSystemConnectionOptions>();
        connections.AddRange(Connections(Existing(personal, "personal", environment, log), PersonalConnectionId, "Personal"));
        connections.AddRange(Connections(Existing(workgroup, "workgroup", environment, log), WorkgroupConnectionId, "Workgroup"));
        return connections;
    }

    /// <summary>The distinct folders that exist, in the order found; the rest are logged.</summary>
    private static List<string> Existing(
        IEnumerable<string> folders, string kind, IOfficeTemplateEnvironment environment, Action<string> log)
    {
        var comparer = environment.IsWindows ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal;
        var seen = new HashSet<string>(comparer);
        var existing = new List<string>();
        foreach (var folder in folders)
        {
            string full;
            try
            {
                full = Path.GetFullPath(folder).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            }
            catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
            {
                log($"OfficeAgent: the Office {kind} templates folder setting is not a usable path; skipped.");
                continue;
            }

            if (!seen.Add(full)) continue;
            if (!environment.DirectoryExists(full))
            {
                log($"OfficeAgent: the Office {kind} templates folder '{full}' does not exist; skipped.");
                continue;
            }
            existing.Add(full);
        }
        return existing;
    }

    /// <summary>One connection per folder; several folders of one kind are numbered from 1.</summary>
    private static IEnumerable<FileSystemConnectionOptions> Connections(List<string> folders, string connectionId, string location) =>
        folders.Select((folder, index) => new FileSystemConnectionOptions
        {
            ConnectionId = folders.Count == 1 ? connectionId : $"{connectionId}-{index + 1}",
            RootPath = folder,
            AllowedExtensions = Extensions.ToList(),
            ReadOnly = true,
            TemplateLocation = location
        });
}
