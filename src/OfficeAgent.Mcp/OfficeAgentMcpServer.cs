using System.Text.Json;
using Microsoft.Extensions.AI;
using Microsoft.Extensions.DependencyInjection;
using ModelContextProtocol.Server;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.SharePoint;
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
                + (creationEnabled && AllowsWordCreation(c.AllowedExtensions)
                    ? " create_document writes new .docx files into this connection's root."
                    : creationEnabled
                        ? " create_document is not available because this connection does not allow .docx."
                        : string.Empty))
            .Concat(options.SharePointConnections
                .Select(c => $"- \"{c.ConnectionId}\" (sharepoint):" + sharePointRegistration
                    + (creationEnabled && HasSharePointCreationTarget(c) && AllowsWordCreation(c.AllowedExtensions)
                        ? " create_document writes new .docx files into this connection's configured folder."
                        : creationEnabled && !HasSharePointCreationTarget(c)
                            ? " create_document is not configured for this connection."
                            : creationEnabled
                                ? " create_document is not available because this connection does not allow .docx."
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

        services.AddWordFormat();
        services.AddOfficeAgent();

        foreach (var connection in options.FileSystemConnections)
        {
            services.AddFileSystemDocumentProvider(connection.ConnectionId, connection.RootPath, o =>
            {
                o.MaximumBytes = connection.MaximumBytes;
                o.AllowedExtensions = connection.AllowedExtensions.ToArray();
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
                canCreateDocuments = creationEnabled && AllowsWordCreation(c.AllowedExtensions)
            })
            .Concat(options.SharePointConnections
                .Select(c => new
                {
                    connectionId = c.ConnectionId,
                    provider = "sharepoint",
                    canCreateDocuments = creationEnabled &&
                        HasSharePointCreationTarget(c) &&
                        AllowsWordCreation(c.AllowedExtensions)
                }))
            .ToArray();
        return JsonSerializer.Serialize(connections);
    }

    private static bool CreationEnabled(OfficeAgentMcpOptions options) =>
        options.AllowCreation &&
        (options.FileSystemConnections.Any(c => AllowsWordCreation(c.AllowedExtensions)) ||
         options.SharePointConnections.Any(c =>
             HasSharePointCreationTarget(c) && AllowsWordCreation(c.AllowedExtensions)));

    private static bool HasSharePointCreationTarget(SharePointConnectionOptions connection) =>
        !string.IsNullOrWhiteSpace(connection.CreationDriveId) &&
        !string.IsNullOrWhiteSpace(connection.CreationFolderItemId);

    private static bool AllowsWordCreation(IEnumerable<string> extensions) =>
        extensions.Any(extension =>
            string.Equals(
                extension.StartsWith(".", StringComparison.Ordinal) ? extension : "." + extension,
                ".docx",
                StringComparison.OrdinalIgnoreCase));

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
            CreationFolderItemId = connection.CreationFolderItemId
        }, http, tokens, store);
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
