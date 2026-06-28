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
    /// an inventory of the configured connections so the agent knows which connectionIds
    /// exist (an MCP client has no other channel to discover them).
    /// </summary>
    public static string InstructionsFor(OfficeAgentMcpOptions options)
    {
        // The connection inventory teaches register_document usage, so it only applies
        // when the registration tools are exposed. Without registration the host pins the
        // agent to the (connectionId, documentId) ids it hands out by other means.
        if (!options.AllowRegistration)
            return OfficeAgentTools.SystemPromptGuidance;

        return OfficeAgentTools.SystemPromptGuidance
            + OfficeAgentTools.RegistrationPromptGuidance
            + ConnectionInventory(options);
    }

    /// <summary>
    /// Lists the host-configured connections - their ids and what each one's registration
    /// source looks like - so the agent can address documents and register new ones
    /// without guessing connectionIds or asking the user for them.
    /// </summary>
    private static string ConnectionInventory(OfficeAgentMcpOptions options)
    {
        var lines = options.FileSystemConnections
            .Select(c => $"- \"{c.ConnectionId}\" (filesystem): a register_document source is a path under this connection's root.")
            .Concat(options.SharePointConnections
                .Select(c => $"- \"{c.ConnectionId}\" (sharepoint): a register_document source is a SharePoint/OneDrive URL or a \"driveId/itemId\" pair."))
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

    /// <summary>Builds the MCP tool list: inspect/find/preview/apply, plus registration and connection-discovery tools when allowed.</summary>
    public static IList<McpServerTool> CreateTools(OfficeAgentTools tools, OfficeAgentMcpOptions options)
    {
        if (tools is null) throw new ArgumentNullException(nameof(tools));
        if (options is null) throw new ArgumentNullException(nameof(options));
        var toolList = tools
            .AsAIFunctions(new OfficeAgentToolsOptions { AllowRegistration = options.AllowRegistration })
            .Select(function => McpServerTool.Create(function))
            .ToList();

        // list_connections pairs with the registration tools: it lets the agent discover
        // which connectionIds exist so it can register documents. A tool is the reliable
        // discovery channel (unlike server instructions, which a client may not surface).
        if (options.AllowRegistration)
            toolList.Add(ConnectionsTool(options));

        return toolList;
    }

    /// <summary>
    /// Builds the <c>list_connections</c> tool from the configured connections, so an
    /// agent can enumerate the connectionIds it may address documents under.
    /// </summary>
    private static McpServerTool ConnectionsTool(OfficeAgentMcpOptions options)
    {
        var connections = options.FileSystemConnections
            .Select(c => new { connectionId = c.ConnectionId, provider = "filesystem" })
            .Concat(options.SharePointConnections
                .Select(c => new { connectionId = c.ConnectionId, provider = "sharepoint" }))
            .ToArray();
        var payload = JsonSerializer.Serialize(connections);

        var function = AIFunctionFactory.Create(
            () => payload,
            new AIFunctionFactoryOptions
            {
                Name = "list_connections",
                Description =
                    "List the connections you can address documents under. Returns [{connectionId, provider}] " +
                    "where provider is \"filesystem\" or \"sharepoint\". Use a connectionId as the connectionId " +
                    "for register_document and the document tools; never ask the user for it."
            });

        return McpServerTool.Create(function);
    }

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
            AllowedExtensions = connection.AllowedExtensions.ToArray()
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
