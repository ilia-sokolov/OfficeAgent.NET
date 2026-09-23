using System.Text.Json;
using Microsoft.Extensions.AI;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;
using ModelContextProtocol.Server;
using OfficeAgent.Abstractions;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.SharePoint;
using OfficeAgent.PowerPoint;
using OfficeAgent.Excel;
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
    /// The version the server reports in <c>serverInfo</c> and <c>/healthz</c>: the package version,
    /// prerelease label included, without the source-control suffix after <c>+</c>.
    /// </summary>
    /// <remarks>
    /// The four-part assembly version cannot carry a prerelease label, so a candidate built as
    /// 1.0.0-rc.1 once reported itself as the final 1.0.0. It is only the fallback now.
    /// </remarks>
    internal static string ReportedVersion(string? informationalVersion, Version? assemblyVersion)
    {
        var informational = informationalVersion?.Split('+')[0].Trim();
        return !string.IsNullOrEmpty(informational)
            ? informational!
            : assemblyVersion?.ToString(3) ?? "0.0.0";
    }

    /// <summary>
    /// Instructions advertised to MCP clients: the same contract the in-process
    /// Microsoft.Extensions.AI tools teach, plus registration guidance when enabled, and
    /// an inventory of the configured connections when registration or creation is
    /// enabled, so the agent knows which connectionIds exist.
    /// </summary>
    public static string InstructionsFor(OfficeAgentMcpOptions options, bool includeConnectionInventory = true)
    {
        var creationEnabled = CreationEnabled(options);
        var hasConnections = HasConnections(options);
        var registrationEnabled = RegistrationEnabled(options);

        return OfficeAgentTools.SystemPromptGuidance
            + (registrationEnabled ? OfficeAgentTools.RegistrationPromptGuidance : string.Empty)
            + (creationEnabled ? OfficeAgentTools.CreationPromptGuidance : string.Empty)
            + (HasEphemeral(options) ? OfficeAgentTools.EphemeralPromptGuidance : string.Empty)
            + (options.AllowInlineContent ? OfficeAgentTools.InlineContentPromptGuidance : string.Empty)
            + (includeConnectionInventory
                ? ConnectionInventory(options, registrationEnabled, creationEnabled)
                : string.Empty);
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
        var session = HasEphemeral(options)
            ? new[]
            {
                $"- \"{EphemeralId(options)}\" (session): documents live here for this run only, " +
                "held by the server rather than in storage; import_document_content puts one in and " +
                "export_document_content takes the result out." +
                (creationEnabled ? " create_document makes a new document in it; the name's extension picks the format." : string.Empty)
            }
            : Array.Empty<string>();

        var lines = session
            .Concat(options.FileSystemConnections
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
                            : string.Empty))))
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

        // No configuration at all is not an error: EphemeralId supplies a session
        // connection, so the server starts usable rather than exiting. StartupNotice tells
        // the operator that is what happened.
        AddFormats(services);
        services.AddOfficeAgent();
        services.TryAddSingleton<IConnectionAccessPolicy, AllowAllConnectionAccessPolicy>();
        services.TryAddSingleton<ITrustedPrincipalAccessor, AnonymousPrincipalAccessor>();

        // Parsed before anything is registered so a bad value names itself at startup
        // rather than silently leaving the connection on the global default.
        foreach (var connection in options.FileSystemConnections)
            ParseChangeMode(connection.DefaultChangeMode, connection.ConnectionId);
        foreach (var connection in options.SharePointConnections)
            ParseChangeMode(connection.DefaultChangeMode, connection.ConnectionId);

        if (HasEphemeral(options))
        {
            services.AddMemoryDocumentProvider(EphemeralId(options), o =>
            {
                o.MaximumTotalBytes = options.EphemeralMaximumTotalBytes;
                o.AllowedExtensions = CreatableExtensions.Value.ToList();
            });
        }

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

        services.AddSingleton(sp => new OfficeAgentTools(
            sp.GetRequiredService<OfficeAgentClient>(),
            sp.GetRequiredService<IConnectionAccessPolicy>(),
            sp.GetRequiredService<ITrustedPrincipalAccessor>()));
        return services;
    }

    /// <summary>Builds the MCP tool list, adding independently enabled staging and connection-discovery tools.</summary>
    public static IList<McpServerTool> CreateTools(OfficeAgentTools tools, OfficeAgentMcpOptions options)
    {
        if (tools is null) throw new ArgumentNullException(nameof(tools));
        if (options is null) throw new ArgumentNullException(nameof(options));
        var creationEnabled = CreationEnabled(options);
        var hasConnections = HasConnections(options);

        var toolList = tools
            .AsAIFunctions(new OfficeAgentToolsOptions
            {
                AllowRegistration = RegistrationEnabled(options),
                AllowCreation = creationEnabled,
                AllowInlineContent = options.AllowInlineContent,
                AllowEphemeralDocuments = HasEphemeral(options),
                // With no connection configured there is no connectionId any of the
                // document tools could be given, so they are left out rather than offered
                // and failed on first use.
                AllowConnectionAddressing = hasConnections
            })
            .Select(function => McpServerTool.Create(function))
            .ToList();

        // A tool is the reliable discovery channel for either staging capability; unlike
        // server instructions, clients consistently surface it.
        if (hasConnections && (options.AllowRegistration || creationEnabled))
            toolList.Add(ConnectionsTool(tools, options, creationEnabled));

        return toolList;
    }

    /// <summary>
    /// Whether any connection exists for a document tool to name - a session connection
    /// counts, because an agent addresses documents in it exactly as it would in storage.
    /// </summary>
    private static bool HasConnections(OfficeAgentMcpOptions options) =>
        options.FileSystemConnections.Count > 0 ||
        options.SharePointConnections.Count > 0 ||
        HasEphemeral(options);

    /// <summary>The session connection minted when the server is started with no configuration.</summary>
    public const string DefaultSessionConnectionId = "session";

    /// <summary>
    /// Whether the host configured nothing at all - no storage, no session connection, no
    /// inline content.
    /// </summary>
    private static bool IsZeroConfigured(OfficeAgentMcpOptions options) =>
        options.FileSystemConnections.Count == 0 &&
        options.SharePointConnections.Count == 0 &&
        !options.AllowInlineContent &&
        string.IsNullOrWhiteSpace(options.EphemeralConnectionId);

    /// <summary>
    /// The session connection the server will actually run with: the configured one, or the
    /// implied one when nothing was configured.
    /// </summary>
    /// <remarks>
    /// A server with no configuration used to refuse to start. Failing loudly was right
    /// while there was nothing it could usefully do, but a session connection needs no
    /// storage, no credentials and no filesystem reach - so there is now something, and
    /// exiting instead is a worse answer than starting. The fallback is exactly equivalent
    /// to setting <c>EphemeralConnectionId=session</c> and <c>AllowCreation=true</c>, and
    /// <see cref="StartupNotice"/> says so on stderr, because a server that quietly does
    /// something other than what the operator believes they configured is the failure this
    /// has to avoid.
    /// </remarks>
    private static string EphemeralId(OfficeAgentMcpOptions options) =>
        IsZeroConfigured(options)
            ? DefaultSessionConnectionId
            : options.EphemeralConnectionId?.Trim() ?? string.Empty;

    /// <summary>
    /// Whether the registration and source-addressed tools are worth offering.
    /// </summary>
    /// <remarks>
    /// They take a source - a path under a root, or a SharePoint URL - so they need a
    /// connection that has one. A session connection does not: its documents arrive through
    /// import_document_content, and there is no external item for register_document to name.
    /// Offering them on a session-only server would hand the agent three tools that can only
    /// fail, which costs it a turn to discover.
    /// </remarks>
    private static bool RegistrationEnabled(OfficeAgentMcpOptions options) =>
        options.AllowRegistration &&
        (options.FileSystemConnections.Count > 0 || options.SharePointConnections.Count > 0);

    private static bool HasEphemeral(OfficeAgentMcpOptions options) =>
        !string.IsNullOrWhiteSpace(EphemeralId(options));

    /// <summary>
    /// A line for the operator when the server's effective configuration is not the one they
    /// wrote, or <see langword="null"/> when it is. Logged to stderr at startup.
    /// </summary>
    public static string? StartupNotice(OfficeAgentMcpOptions options) =>
        IsZeroConfigured(options)
            ? "No storage is configured, so OfficeAgent started with an in-memory session " +
              $"connection called '{DefaultSessionConnectionId}' and document creation enabled. " +
              "Documents created there last only while this server runs and are written nowhere. " +
              "To edit documents on disk instead, set OfficeAgent__FileSystemConnections__0__ConnectionId " +
              "and OfficeAgent__FileSystemConnections__0__RootPath - see " +
              "https://github.com/ilia-sokolov/OfficeAgent.NET/blob/main/docs/mcp-server.md"
            : null;

    /// <summary>
    /// Builds the <c>list_connections</c> tool from the configured connections, so an
    /// agent can enumerate the connectionIds it may address documents under.
    /// </summary>
    private static McpServerTool ConnectionsTool(
        OfficeAgentTools tools,
        OfficeAgentMcpOptions options,
        bool creationEnabled)
    {
        var addressedTools = options.AllowRegistration && creationEnabled
            ? "register_document, the document tools, and create_document (connections where canCreateDocuments is true)"
            : creationEnabled
                ? "the document tools and create_document (connections where canCreateDocuments is true)"
                : "register_document and the document tools";

        var function = AIFunctionFactory.Create(
            (CancellationToken cancellationToken) =>
                ConnectionsPayloadAsync(tools, options, cancellationToken),
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

    internal static async Task<string> ConnectionsPayloadAsync(
        OfficeAgentTools tools,
        OfficeAgentMcpOptions options,
        CancellationToken cancellationToken = default)
    {
        var descriptors = new List<(string Id, string Provider, bool Creatable)>();
        if (HasEphemeral(options))
            descriptors.Add((EphemeralId(options), "session", CreationEnabled(options)));
        descriptors.AddRange(options.FileSystemConnections.Select(connection =>
            (connection.ConnectionId, "filesystem", CreationEnabled(options) &&
                AllowsCreatableExtension(connection.AllowedExtensions))));
        descriptors.AddRange(options.SharePointConnections.Select(connection =>
            (connection.ConnectionId, "sharepoint", CreationEnabled(options) &&
                HasSharePointCreationTarget(connection) &&
                AllowsCreatableExtension(connection.AllowedExtensions))));

        var visible = new List<object>();
        foreach (var descriptor in descriptors)
        {
            var read = await tools.CanAccessConnectionAsync(
                descriptor.Id, ConnectionCapability.Read, cancellationToken).ConfigureAwait(false);
            var register = await tools.CanAccessConnectionAsync(
                descriptor.Id, ConnectionCapability.Register, cancellationToken).ConfigureAwait(false);
            var create = descriptor.Creatable && await tools.CanAccessConnectionAsync(
                descriptor.Id, ConnectionCapability.Create, cancellationToken).ConfigureAwait(false);
            var edit = await tools.CanAccessConnectionAsync(
                descriptor.Id, ConnectionCapability.Edit, cancellationToken).ConfigureAwait(false);
            var delete = await tools.CanAccessConnectionAsync(
                descriptor.Id, ConnectionCapability.Delete, cancellationToken).ConfigureAwait(false);
            if (!read && !register && !create && !edit && !delete) continue;

            visible.Add(new
            {
                connectionId = descriptor.Id,
                provider = descriptor.Provider,
                canCreateDocuments = create
            });
        }

        return JsonSerializer.Serialize(visible);
    }

    internal static string ConnectionsPayload(OfficeAgentMcpOptions options)
    {
        var creationEnabled = CreationEnabled(options);
        var ephemeral = HasEphemeral(options)
            ? new[]
            {
                new
                {
                    connectionId = EphemeralId(options),
                    provider = "session",
                    canCreateDocuments = creationEnabled
                }
            }
            : Array.Empty<object>().Select(_ => new { connectionId = "", provider = "", canCreateDocuments = false });

        var connections = ephemeral
            .Concat(options.FileSystemConnections
            .Select(c => new
            {
                connectionId = c.ConnectionId,
                provider = "filesystem",
                canCreateDocuments = creationEnabled && AllowsCreatableExtension(c.AllowedExtensions)
            }))
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

    /// <remarks>
    /// <see cref="OfficeAgentMcpOptions.AllowCreation"/> guards authoring agent-named files
    /// under a host's storage root, which is a decision the host should make. A
    /// zero-configuration server has no root: its session connection starts empty and holds
    /// nothing that outlives the process, so withholding creation there would leave an
    /// agent with a connection it can do nothing with until a document is handed to it.
    /// That is why the implied session comes with creation, and a session the host
    /// configured deliberately still respects the flag.
    /// </remarks>
    private static bool CreationEnabled(OfficeAgentMcpOptions options) =>
        IsZeroConfigured(options) ||
        (options.AllowCreation &&
         (HasEphemeral(options) ||
          options.FileSystemConnections.Any(c => AllowsCreatableExtension(c.AllowedExtensions)) ||
          options.SharePointConnections.Any(c =>
              HasSharePointCreationTarget(c) && AllowsCreatableExtension(c.AllowedExtensions))));

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
        services.AddExcelFormat();
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
    public static IList<McpServerTool> BuildToolset(
        OfficeAgentMcpOptions options,
        IConnectionAccessPolicy? connectionAccess = null,
        ITrustedPrincipalAccessor? principalAccessor = null,
        IAuditActorProvider? auditActorProvider = null)
    {
        var services = new ServiceCollection();
        services.AddSingleton(new HttpClient());
        if (connectionAccess is not null) services.AddSingleton(connectionAccess);
        if (principalAccessor is not null) services.AddSingleton(principalAccessor);
        if (auditActorProvider is not null) services.AddSingleton(auditActorProvider);
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
        var authMode = connection.AuthMode?.Trim().ToLowerInvariant();
        return authMode switch
        {
            "onbehalfof" or "on-behalf-of" or "obo" => new OnBehalfOfAccessTokenProvider(new OnBehalfOfOptions
            {
                TenantId = connection.TenantId,
                ClientId = connection.ClientId,
                ClientSecret = connection.ClientSecret,
                Scope = connection.OnBehalfOfScope,
                Authority = connection.LoginAuthority
            }, http),
            "apponly" or "app-only" or "application" => new AppOnlyAccessTokenProvider(new AppOnlyOptions
            {
                TenantId = connection.TenantId,
                ClientId = connection.ClientId,
                ClientSecret = connection.ClientSecret,
                Scope = connection.AppOnlyScope,
                Authority = connection.LoginAuthority
            }, http),
            _ => throw new InvalidOperationException(
                $"SharePoint connection '{connection.ConnectionId}' has AuthMode '{connection.AuthMode}'. " +
                "Expected appOnly or onBehalfOf.")
        };
    }
}
