using System.Text.Json;
using Microsoft.Extensions.AI;
using Microsoft.Extensions.DependencyInjection;
using ModelContextProtocol.Server;
using OfficeAgent.Abstractions;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.SharePoint;
using OfficeAgent.PowerPoint;
using OfficeAgent.Word;

namespace OfficeAgent.Mcp;

/// <summary>
/// Composition root shared by the stdio and HTTP hosts: turns
/// <see cref="OfficeAgentMcpOptions"/> into provider registrations and projects
/// <see cref="OfficeAgentTools"/> as MCP tools.
/// </summary>
public static class OfficeAgentMcpServer
{
    /// <summary>The MCP server name advertised during initialization.</summary>
    public const string ServerName = "officeagent";

    /// <summary>
    /// Instructions advertised to MCP clients: the same contract the in-process
    /// Microsoft.Extensions.AI tools teach, plus registration guidance when enabled, and
    /// an inventory of the configured connections when registration or creation is
    /// enabled, so the agent knows which connectionIds exist.
    /// </summary>
    public static string InstructionsFor(OfficeAgentMcpOptions options)
    {
        var creationEnabled = CreationEnabled(options);
        if (!options.AllowRegistration && !creationEnabled)
            return OfficeAgentTools.SystemPromptGuidance;

        return OfficeAgentTools.SystemPromptGuidance
            + (options.AllowRegistration ? OfficeAgentTools.RegistrationPromptGuidance : string.Empty)
            + (creationEnabled ? OfficeAgentTools.CreationPromptGuidance : string.Empty)
            + ConnectionInventory(options, options.AllowRegistration, creationEnabled);
    }

    /// <summary>
    /// Lists the host-configured connections and the enabled staging operations so the
    /// agent can address documents without guessing connectionIds.
    /// </summary>
    private static string ConnectionInventory(
        OfficeAgentMcpOptions options,
        bool registrationEnabled,
        bool creationEnabled)
    {
        var filesystemRegistration = registrationEnabled
            ? " a register_document source is a path under this connection's root."
            : string.Empty;
        var sharePointRegistration = registrationEnabled
            ? " a register_document source is a SharePoint/OneDrive URL or a \"driveId/itemId\" pair."
            : string.Empty;
        var lines = options.FileSystemConnections
            .Select(c => $"- \"{c.ConnectionId}\" (filesystem):" + filesystemRegistration
                + (creationEnabled && AllowsCreatableExtension(c.AllowedExtensions)
                    ? " create_document writes new documents into this connection's root; the name's extension picks the format."
                    : creationEnabled
                        ? " create_document is not available because this connection allows no creatable extension."
                        : string.Empty))
            .Concat(options.SharePointConnections
                .Select(c => $"- \"{c.ConnectionId}\" (sharepoint):" + sharePointRegistration
                    + (creationEnabled && HasSharePointCreationTarget(c) && AllowsCreatableExtension(c.AllowedExtensions)
                        ? " create_document writes new documents into this connection's configured folder; the name's extension picks the format."
                        : creationEnabled && !HasSharePointCreationTarget(c)
                            ? " create_document is not configured for this connection."
                            : creationEnabled
                                ? " create_document is not available because this connection allows no creatable extension."
                            : string.Empty)))
            .ToList();

        return lines.Count == 0
            ? string.Empty
            : "\n\nConfigured connections (use these connectionId values; never ask the user for them):\n"
                + string.Join("\n", lines);
    }

    /// <summary>
    /// Registers the OfficeAgent engine, the configured document providers, and the
    /// tool projection with the host's service collection.
    /// </summary>
    public static IServiceCollection AddOfficeAgentMcp(this IServiceCollection services, OfficeAgentMcpOptions options)
    {
        if (options is null) throw new ArgumentNullException(nameof(options));
        if (options.FileSystemConnections.Count == 0 && options.SharePointConnections.Count == 0)
            throw new InvalidOperationException(
                "The OfficeAgent MCP server requires at least one connection. Configure " +
                "OfficeAgent:FileSystemConnections or OfficeAgent:SharePointConnections.");

        AddFormats(services);
        services.AddOfficeAgent();

        // Parsed before anything is registered so a bad value names itself at startup
        // rather than silently leaving the connection on the global default.
        foreach (var connection in options.FileSystemConnections)
            ParseChangeMode(connection.DefaultChangeMode, connection.ConnectionId);
        foreach (var connection in options.SharePointConnections)
            ParseChangeMode(connection.DefaultChangeMode, connection.ConnectionId);

        foreach (var connection in options.FileSystemConnections)
        {
            services.AddFileSystemDocumentProvider(connection.ConnectionId, connection.RootPath, o =>
            {
                o.MaximumBytes = connection.MaximumBytes;
                o.AllowedExtensions = connection.AllowedExtensions.ToArray();
                o.DefaultChangeMode = ParseChangeMode(connection.DefaultChangeMode, connection.ConnectionId);
            });
        }

        foreach (var connection in options.SharePointConnections)
        {
            var captured = connection;
            services.AddSingleton<IDocumentProvider>(sp => CreateSharePointProvider(sp, captured));
        }

        services.AddSingleton(sp => new OfficeAgentTools(sp.GetRequiredService<OfficeAgentClient>()));
        return services;
    }

    /// <summary>Builds the MCP tool list, adding independently enabled staging and connection-discovery tools.</summary>
    public static IList<McpServerTool> CreateTools(OfficeAgentTools tools, OfficeAgentMcpOptions options)
    {
        if (tools is null) throw new ArgumentNullException(nameof(tools));
        if (options is null) throw new ArgumentNullException(nameof(options));
        var creationEnabled = CreationEnabled(options);
        var toolList = tools
            .AsAIFunctions(new OfficeAgentToolsOptions
            {
                AllowRegistration = options.AllowRegistration,
                AllowCreation = creationEnabled
            })
            .Select(function => McpServerTool.Create(function))
            .ToList();

        // A tool is the reliable discovery channel for either staging capability; unlike
        // server instructions, clients consistently surface it.
        if (options.AllowRegistration || creationEnabled)
            toolList.Add(ConnectionsTool(options, creationEnabled));

        return toolList;
    }

    /// <summary>
    /// Builds the <c>list_connections</c> tool from the configured connections, so an
    /// agent can enumerate the connectionIds it may address documents under.
    /// </summary>
    private static McpServerTool ConnectionsTool(OfficeAgentMcpOptions options, bool creationEnabled)
    {
        var payload = ConnectionsPayload(options);
        var addressedTools = options.AllowRegistration && creationEnabled
            ? "register_document, the document tools, and create_document (connections where canCreateDocuments is true)"
            : creationEnabled
                ? "the document tools and create_document (connections where canCreateDocuments is true)"
                : "register_document and the document tools";

        var function = AIFunctionFactory.Create(
            () => payload,
            new AIFunctionFactoryOptions
            {
                Name = "list_connections",
                Description =
                    "List the connections you can address documents under. Returns [{connectionId, provider, canCreateDocuments}] " +
                    "where provider is \"filesystem\" or \"sharepoint\". Use a connectionId as the connectionId " +
                    $"for {addressedTools}; never ask the user for it."
            });

        return McpServerTool.Create(function);
    }

    internal static string ConnectionsPayload(OfficeAgentMcpOptions options)
    {
        var creationEnabled = CreationEnabled(options);
        var connections = options.FileSystemConnections
            .Select(c => new
            {
                connectionId = c.ConnectionId,
                provider = "filesystem",
                canCreateDocuments = creationEnabled && AllowsCreatableExtension(c.AllowedExtensions)
            })
            .Concat(options.SharePointConnections
                .Select(c => new
                {
                    connectionId = c.ConnectionId,
                    provider = "sharepoint",
                    canCreateDocuments = creationEnabled &&
                        HasSharePointCreationTarget(c) &&
                        AllowsCreatableExtension(c.AllowedExtensions)
                }))
            .ToArray();
        return JsonSerializer.Serialize(connections);
    }

    private static bool CreationEnabled(OfficeAgentMcpOptions options) =>
        options.AllowCreation &&
        (options.FileSystemConnections.Any(c => AllowsCreatableExtension(c.AllowedExtensions)) ||
         options.SharePointConnections.Any(c =>
             HasSharePointCreationTarget(c) && AllowsCreatableExtension(c.AllowedExtensions)));

    private static bool HasSharePointCreationTarget(SharePointConnectionOptions connection) =>
        !string.IsNullOrWhiteSpace(connection.CreationDriveId) &&
        !string.IsNullOrWhiteSpace(connection.CreationFolderItemId);

    /// <summary>
    /// The format modules this server speaks. The single place they are named, so the
    /// toolset, the connection inventory, and the server instructions cannot drift apart.
    /// </summary>
    private static void AddFormats(IServiceCollection services)
    {
        services.AddWordFormat();
        services.AddPowerPointFormat();
    }

    /// <summary>
    /// The extensions this server can mint a blank document for, derived from the very
    /// modules <see cref="AddFormats"/> registers rather than a hard-coded list - adding
    /// a format module must not silently leave capability reporting behind.
    /// </summary>
    private static readonly Lazy<string[]> CreatableExtensions = new(() =>
    {
        var services = new ServiceCollection();
        AddFormats(services);
        using var provider = services.BuildServiceProvider();
        return provider.GetServices<IFormatModule>()
            .OfType<IBlankDocumentFactory>()
            .Select(factory => factory.Extension)
            .Where(extension => !string.IsNullOrWhiteSpace(extension))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
    });

    /// <summary>
    /// Whether a connection accepts at least one extension the engine can actually mint.
    /// A connection limited to formats no registered module creates cannot create
    /// anything, and advertising it would send the agent into a call that can only fail.
    /// </summary>
    private static bool AllowsCreatableExtension(IEnumerable<string> extensions) =>
        extensions.Any(extension => CreatableExtensions.Value.Contains(
            extension.StartsWith(".", StringComparison.Ordinal) ? extension : "." + extension,
            StringComparer.OrdinalIgnoreCase));

    /// <summary>
    /// Builds the full toolset from configuration alone: composes the engine and
    /// providers in a dedicated container (kept alive for the process lifetime,
    /// since the tools close over it) and projects them as MCP tools.
    /// </summary>
    public static IList<McpServerTool> BuildToolset(OfficeAgentMcpOptions options)
    {
        var services = new ServiceCollection();
        services.AddSingleton(new HttpClient());
        services.AddOfficeAgentMcp(options);
        var provider = services.BuildServiceProvider();
        return CreateTools(provider.GetRequiredService<OfficeAgentTools>(), options);
    }

    private static SharePointDocumentProvider CreateSharePointProvider(
        IServiceProvider services, SharePointConnectionOptions connection)
    {
        var http = services.GetRequiredService<HttpClient>();
        var tokens = CreateTokenProvider(connection, http);

        ISharePointRegistrationStore store = string.IsNullOrWhiteSpace(connection.RegistrationIndexPath)
            ? new InMemoryRegistrationStore()
            : new JsonFileRegistrationStore(connection.RegistrationIndexPath);

        return new SharePointDocumentProvider(new SharePointDocumentProviderOptions
        {
            ConnectionId = connection.ConnectionId,
            GraphBaseUrl = connection.GraphBaseUrl,
            MaximumBytes = connection.MaximumBytes,
            AllowedExtensions = connection.AllowedExtensions.ToArray(),
            CreationDriveId = connection.CreationDriveId,
            CreationFolderItemId = connection.CreationFolderItemId,
            DefaultChangeMode = ParseChangeMode(connection.DefaultChangeMode, connection.ConnectionId)
        }, http, tokens, store);
    }

    /// <summary>
    /// Reads a connection's <c>DefaultChangeMode</c>. It is bound as a string rather than
    /// as the enum on purpose: the configuration binder silently skips a collection element
    /// it cannot bind, so a typo in this one optional field would drop the whole connection
    /// and surface as "no connections configured" - a diagnostic pointing nowhere near the
    /// mistake. Parsing it here names the connection and the accepted values instead.
    /// </summary>
    private static ChangeMode ParseChangeMode(string value, string connectionId)
    {
        if (string.IsNullOrWhiteSpace(value)) return ChangeMode.Tracked;
        if (Enum.TryParse<ChangeMode>(value.Trim(), ignoreCase: true, out var mode)) return mode;

        throw new InvalidOperationException(
            $"Connection '{connectionId}' has DefaultChangeMode '{value}'. Expected Tracked or Direct.");
    }

    private static IAccessTokenProvider CreateTokenProvider(SharePointConnectionOptions connection, HttpClient http)
    {
        return connection.AuthMode?.Trim().ToLowerInvariant() switch
        {
            "onbehalfof" or "on-behalf-of" or "obo" => new OnBehalfOfAccessTokenProvider(new OnBehalfOfOptions
            {
                TenantId = connection.TenantId,
                ClientId = connection.ClientId,
                ClientSecret = connection.ClientSecret,
                Scope = connection.OnBehalfOfScope,
                Authority = connection.LoginAuthority
            }, http),
            // "appOnly" (the default): authenticate as the app's own identity, no user.
            _ => new AppOnlyAccessTokenProvider(new AppOnlyOptions
            {
                TenantId = connection.TenantId,
                ClientId = connection.ClientId,
                ClientSecret = connection.ClientSecret,
                Authority = connection.LoginAuthority
            }, http)
        };
    }
}
