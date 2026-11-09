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
using OfficeAgent.Excel.Calculation;
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
            + (options.RequirePlanSnapshot ? OfficeAgentTools.SnapshotRequiredPromptGuidance : string.Empty)
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
            .Concat(StorageConnections.Describe(options).Select(c =>
                $"- \"{c.ConnectionId}\" ({c.Kind.Provider}):"
                + (registrationEnabled ? $" a register_document source is {c.Kind.Source}." : string.Empty)
                + (creationEnabled && c.HasCreationTarget && AllowsCreatableExtension(c.AllowedExtensions)
                    ? $" create_document {c.Kind.Creation}; the name's extension picks the format."
                    : creationEnabled && !c.HasCreationTarget
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

        // No configuration at all is not an error: EphemeralId supplies a session
        // connection, so the server starts usable rather than exiting. StartupNotice tells
        // the operator that is what happened.
        AddFormats(services);
        services.AddOfficeAgent();
        services.TryAddSingleton(new OfficeAgentClientOptions
        {
            RevisionAuthorFromActor = options.RevisionAuthorFromActor
        });
        services.TryAddSingleton<IAuditActorProvider, OnBehalfOfActorProvider>();
        services.TryAddSingleton<IConnectionAccessPolicy, AllowAllConnectionAccessPolicy>();
        services.TryAddSingleton<ITrustedPrincipalAccessor, AnonymousPrincipalAccessor>();

        // Parsed before anything is registered so a bad value names itself at startup
        // rather than silently leaving the connection on the global default.
        for (var index = 0; index < options.FileSystemConnections.Count; index++)
        {
            var connection = options.FileSystemConnections[index];
            ParseChangeMode(connection.DefaultChangeMode, connection.ConnectionId);
            ParseChangeGranularity(connection.DefaultChangeGranularity, connection.ConnectionId);
            ParseTemplateLocation(connection.TemplateLocation, connection.ConnectionId);
            ValidateFileSystemTemplatesFolder(connection, index);
        }
        for (var index = 0; index < options.SharePointConnections.Count; index++)
        {
            var connection = options.SharePointConnections[index];
            ParseChangeMode(connection.DefaultChangeMode, connection.ConnectionId);
            ParseChangeGranularity(connection.DefaultChangeGranularity, connection.ConnectionId);
            ParseTemplateLocation(connection.TemplateLocation, connection.ConnectionId);
            ValidateSharePointTemplatesFolder(connection, index);
        }
        StorageConnections.Validate(options);

        if (HasEphemeral(options))
        {
            services.AddMemoryDocumentProvider(EphemeralId(options), o =>
            {
                o.MaximumTotalBytes = options.EphemeralMaximumTotalBytes;
                o.AllowedExtensions = CreatableExtensions.Value.ToList();
                o.CalculateFormulas = options.EphemeralCalculateFormulas;
                o.RejectFormulaErrors = options.EphemeralRejectFormulaErrors;
            });
        }

        foreach (var connection in options.FileSystemConnections)
        {
            services.AddFileSystemDocumentProvider(connection.ConnectionId, connection.RootPath, o =>
            {
                o.MaximumBytes = connection.MaximumBytes;
                o.AllowedExtensions = connection.AllowedExtensions.ToArray();
                o.DefaultChangeMode = ParseChangeMode(connection.DefaultChangeMode, connection.ConnectionId);
                o.DefaultChangeGranularity = ParseChangeGranularity(connection.DefaultChangeGranularity, connection.ConnectionId);
                o.CalculateFormulas = connection.CalculateFormulas;
                o.RejectFormulaErrors = connection.RejectFormulaErrors;
                o.TemplatesFolder = string.IsNullOrWhiteSpace(connection.TemplatesFolder) ? null : connection.TemplatesFolder.Trim();
                o.TemplateLocation = ParseTemplateLocation(connection.TemplateLocation, connection.ConnectionId);
                o.ReadOnly = connection.ReadOnly;
            });
        }

        foreach (var connection in options.SharePointConnections)
        {
            var captured = connection;
            services.AddSingleton<IDocumentProvider>(sp => CreateSharePointProvider(sp, captured));
        }

        StorageConnections.Register(services, options);

        services.AddSingleton(sp => new OfficeAgentTools(
            sp.GetRequiredService<OfficeAgentClient>(),
            sp.GetRequiredService<IConnectionAccessPolicy>(),
            sp.GetRequiredService<ITrustedPrincipalAccessor>())
        {
            RequirePlanSnapshot = options.RequirePlanSnapshot
        });
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
                AllowConnectionAddressing = hasConnections,
                StrictToolSchemas = options.StrictToolSchemas,
                // Names only the kinds of storage this server has, so an agent is never told
                // to pass a source no connection accepts.
                SourceDescription = HasStorage(options) ? StorageConnections.SourceDescription(options) : null
            })
            .Select(function => McpServerTool.Create(function, ToolAnnotations.For(function.Name)))
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
        HasStorage(options) || HasEphemeral(options);

    /// <summary>Whether any connection other than the session one is configured.</summary>
    private static bool HasStorage(OfficeAgentMcpOptions options) =>
        StorageConnections.Describe(options).Count > 0;

    /// <summary>The session connection minted when the server is started with no configuration.</summary>
    public const string DefaultSessionConnectionId = "session";

    /// <summary>
    /// Whether the host configured nothing at all - no storage, no session connection, no
    /// inline content.
    /// </summary>
    private static bool IsZeroConfigured(OfficeAgentMcpOptions options) =>
        !HasStorage(options) &&
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
    /// They take a source - a path under a root, a SharePoint URL, a file id or link - so
    /// they need a connection that has one. A session connection does not: its documents arrive through
    /// import_document_content, and there is no external item for register_document to name.
    /// Offering them on a session-only server would hand the agent three tools that can only
    /// fail, which costs it a turn to discover.
    /// </remarks>
    private static bool RegistrationEnabled(OfficeAgentMcpOptions options) =>
        options.AllowRegistration && HasStorage(options);

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
                    "List the connections you can address documents under. Returns [{connectionId, provider, canCreateDocuments, " +
                    "allowedExtensions, defaultChangeMode, defaultChangeGranularity, calculatesFormulas, listable, templatesFolder, " +
                    "templateLocation, readOnly}] where provider names the storage, such as \"filesystem\" or \"sharepoint\"; " +
                    "listable says list_documents works on it, templateLocation (organization, workgroup, personal, connection, or null) " +
                    "says it holds templates, and readOnly true means it is never written to. Use a connectionId as the connectionId " +
                    $"for {addressedTools}; never ask the user for it."
            });

        return McpServerTool.Create(function, ToolAnnotations.For(function.Name));
    }

    /// <summary>
    /// One <c>list_connections</c> entry before access is checked. Built from configuration
    /// alone, so it carries ids and capabilities but never a credential, URL, container,
    /// bucket or prefix.
    /// </summary>
    private sealed record ConnectionEntry(
        string ConnectionId,
        string Provider,
        bool Creatable,
        IReadOnlyList<string> AllowedExtensions,
        string DefaultChangeMode,
        string DefaultChangeGranularity,
        bool CalculatesFormulas)
    {
        /// <summary>Whether list_documents works on the connection.</summary>
        public bool Listable { get; init; }

        public string? TemplatesFolder { get; init; }

        public TemplateLocationKind TemplateLocation { get; init; }

        /// <summary>Whether the connection refuses every write, where configuration says; otherwise null.</summary>
        public bool? ReadOnly { get; init; }
    }

    private static List<ConnectionEntry> ConnectionEntries(OfficeAgentMcpOptions options)
    {
        var creationEnabled = CreationEnabled(options);
        var entries = new List<ConnectionEntry>();
        if (HasEphemeral(options))
            entries.Add(new ConnectionEntry(
                EphemeralId(options), "session", creationEnabled, CreatableExtensions.Value,
                nameof(ChangeMode.Tracked), nameof(ChangeGranularity.Span), options.EphemeralCalculateFormulas)
            {
                Listable = true,
                ReadOnly = false
            });

        entries.AddRange(StorageConnections.Describe(options).Select(c =>
        {
            var location = ParseTemplateLocation(c.TemplateLocation ?? string.Empty, c.ConnectionId);
            return new ConnectionEntry(
                c.ConnectionId,
                c.Kind.Provider,
                creationEnabled && c.HasCreationTarget && AllowsCreatableExtension(c.AllowedExtensions),
                c.AllowedExtensions
                    .Where(extension => !string.IsNullOrWhiteSpace(extension))
                    .Select(extension => NormalizeExtension(extension))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList(),
                ParseChangeMode(c.DefaultChangeMode, c.ConnectionId).ToString(),
                ParseChangeGranularity(c.DefaultChangeGranularity, c.ConnectionId).ToString(),
                c.CalculatesFormulas)
            {
                // Every storage kind this server configures can list; the tool itself asks the
                // live providers, which is what a custom provider is judged by.
                Listable = true,
                TemplatesFolder = c.TemplatesFolder,
                TemplateLocation = location == TemplateLocationKind.None && c.TemplatesFolder is not null
                    ? TemplateLocationKind.Connection
                    : location,
                ReadOnly = c.ReadOnly
            };
        }));
        return entries;
    }

    /// <summary>
    /// Replaces what configuration says about listing and templates with what the connection's
    /// provider reports, which also covers providers this server cannot judge from configuration.
    /// </summary>
    private static ConnectionEntry WithListing(ConnectionEntry entry, IReadOnlyList<ConnectionListing> listings)
    {
        var listing = listings.FirstOrDefault(l => string.Equals(l.ConnectionId, entry.ConnectionId, StringComparison.Ordinal));
        return listing is null
            ? entry
            : entry with
            {
                Listable = listing.Listable,
                TemplatesFolder = listing.TemplatesFolder,
                TemplateLocation = listing.TemplateLocation
            };
    }

    private static string NormalizeExtension(string extension)
    {
        var trimmed = extension.Trim().ToLowerInvariant();
        return trimmed.StartsWith(".", StringComparison.Ordinal) ? trimmed : "." + trimmed;
    }

    private static object Payload(ConnectionEntry entry, bool canCreate) => new
    {
        connectionId = entry.ConnectionId,
        provider = entry.Provider,
        canCreateDocuments = canCreate,
        allowedExtensions = entry.AllowedExtensions,
        defaultChangeMode = entry.DefaultChangeMode,
        defaultChangeGranularity = entry.DefaultChangeGranularity,
        calculatesFormulas = entry.CalculatesFormulas,
        listable = entry.Listable,
        templatesFolder = entry.TemplatesFolder,
        templateLocation = entry.TemplateLocation == TemplateLocationKind.None
            ? null
            : entry.TemplateLocation.ToString().ToLowerInvariant(),
        readOnly = entry.ReadOnly
    };

    internal static async Task<string> ConnectionsPayloadAsync(
        OfficeAgentTools tools,
        OfficeAgentMcpOptions options,
        CancellationToken cancellationToken = default)
    {
        var visible = new List<object>();
        var listings = tools.DescribeListings();
        foreach (var configured in ConnectionEntries(options))
        {
            var entry = WithListing(configured, listings);
            var read = await tools.CanAccessConnectionAsync(
                entry.ConnectionId, ConnectionCapability.Read, cancellationToken).ConfigureAwait(false);
            var register = await tools.CanAccessConnectionAsync(
                entry.ConnectionId, ConnectionCapability.Register, cancellationToken).ConfigureAwait(false);
            var create = entry.Creatable && await tools.CanAccessConnectionAsync(
                entry.ConnectionId, ConnectionCapability.Create, cancellationToken).ConfigureAwait(false);
            var edit = await tools.CanAccessConnectionAsync(
                entry.ConnectionId, ConnectionCapability.Edit, cancellationToken).ConfigureAwait(false);
            var delete = await tools.CanAccessConnectionAsync(
                entry.ConnectionId, ConnectionCapability.Delete, cancellationToken).ConfigureAwait(false);
            if (!read && !register && !create && !edit && !delete) continue;

            visible.Add(Payload(entry, create));
        }

        return JsonSerializer.Serialize(visible);
    }

    internal static string ConnectionsPayload(OfficeAgentMcpOptions options) =>
        JsonSerializer.Serialize(ConnectionEntries(options).Select(entry => Payload(entry, entry.Creatable)).ToArray());

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
          StorageConnections.Describe(options).Any(c =>
              c.HasCreationTarget && AllowsCreatableExtension(c.AllowedExtensions))));

    /// <summary>
    /// The format modules this server speaks. The single place they are named, so the
    /// toolset, the connection inventory, and the server instructions cannot drift apart.
    /// </summary>
    private static void AddFormats(IServiceCollection services)
    {
        services.AddWordFormat();
        services.AddPowerPointFormat();
        services.AddExcelFormat();
        // Registered unconditionally: it runs only for a connection that sets CalculateFormulas.
        services.AddExcelFormulaCalculation();
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
            DefaultChangeMode = ParseChangeMode(connection.DefaultChangeMode, connection.ConnectionId),
            DefaultChangeGranularity = ParseChangeGranularity(connection.DefaultChangeGranularity, connection.ConnectionId),
            CalculateFormulas = connection.CalculateFormulas,
            RejectFormulaErrors = connection.RejectFormulaErrors,
            TemplatesFolder = string.IsNullOrWhiteSpace(connection.TemplatesFolder) ? null : connection.TemplatesFolder.Trim(),
            TemplateLocation = ParseTemplateLocation(connection.TemplateLocation, connection.ConnectionId)
        }, http, tokens, store);
    }

    /// <summary>
    /// Reads a connection's <c>TemplateLocation</c>, bound as a string for the same reason as
    /// <see cref="ParseChangeMode"/>. Empty means none, or <c>Connection</c> once a templates
    /// folder is set.
    /// </summary>
    internal static TemplateLocationKind ParseTemplateLocation(string value, string connectionId)
    {
        if (string.IsNullOrWhiteSpace(value)) return TemplateLocationKind.None;
        var trimmed = value.Trim();
        if (!char.IsDigit(trimmed[0]) && trimmed[0] != '-' &&
            Enum.TryParse<TemplateLocationKind>(trimmed, ignoreCase: true, out var location) &&
            Enum.IsDefined(typeof(TemplateLocationKind), location))
            return location;

        throw new InvalidOperationException(
            $"Connection '{connectionId}' has TemplateLocation '{value}'. Expected Connection, Personal, Workgroup or Organization.");
    }

    /// <summary>
    /// A filesystem templates folder is a folder below the root. Checked here so the mistake
    /// names the connection and its setting rather than surfacing as the provider's own refusal.
    /// </summary>
    private static void ValidateFileSystemTemplatesFolder(FileSystemConnectionOptions connection, int index)
    {
        if (string.IsNullOrWhiteSpace(connection.TemplatesFolder)) return;
        var folder = connection.TemplatesFolder.Trim();
        var segments = folder.Split(new[] { '/', '\\' }, StringSplitOptions.RemoveEmptyEntries);
        if (Path.IsPathRooted(folder) || segments.Length == 0 || segments.Any(s => s.Trim() is "." or ".."))
            throw new InvalidOperationException(
                $"Connection '{connection.ConnectionId}' has a TemplatesFolder that is not a folder below its root. " +
                $"Set OfficeAgent__FileSystemConnections__{index}__TemplatesFolder to a relative folder such as Templates.");
    }

    /// <summary>A SharePoint templates folder is named the way a folder is registered: by URL or by id pair.</summary>
    private static void ValidateSharePointTemplatesFolder(SharePointConnectionOptions connection, int index)
    {
        if (string.IsNullOrWhiteSpace(connection.TemplatesFolder)) return;
        var folder = connection.TemplatesFolder.Trim();
        var isUrl = Uri.TryCreate(folder, UriKind.Absolute, out var uri) &&
                    (uri.Scheme == Uri.UriSchemeHttps || uri.Scheme == Uri.UriSchemeHttp);
        var trimmed = folder.Trim('/');
        var slash = trimmed.IndexOf('/');
        var isPair = (trimmed.StartsWith("drives/", StringComparison.OrdinalIgnoreCase) &&
                      trimmed.IndexOf("/items/", StringComparison.OrdinalIgnoreCase) > 0) ||
                     (slash > 0 && slash == trimmed.LastIndexOf('/') && slash < trimmed.Length - 1);
        if (!isUrl && !isPair)
            throw new InvalidOperationException(
                $"SharePoint connection '{connection.ConnectionId}' has a TemplatesFolder that is neither a folder URL nor a " +
                $"'driveId/itemId' pair. Set OfficeAgent__SharePointConnections__{index}__TemplatesFolder to one of them.");
    }

    /// <summary>
    /// Reads a connection's <c>DefaultChangeMode</c>. It is bound as a string rather than
    /// as the enum on purpose: the configuration binder silently skips a collection element
    /// it cannot bind, so a typo in this one optional field would drop the whole connection
    /// and surface as "no connections configured" - a diagnostic pointing nowhere near the
    /// mistake. Parsing it here names the connection and the accepted values instead.
    /// </summary>
    internal static ChangeMode ParseChangeMode(string value, string connectionId)
    {
        if (string.IsNullOrWhiteSpace(value)) return ChangeMode.Tracked;
        if (Enum.TryParse<ChangeMode>(value.Trim(), ignoreCase: true, out var mode)) return mode;

        throw new InvalidOperationException(
            $"Connection '{connectionId}' has DefaultChangeMode '{value}'. Expected Tracked or Direct.");
    }

    /// <summary>
    /// Reads a connection's <c>DefaultChangeGranularity</c>, bound as a string for the same
    /// reason as <see cref="ParseChangeMode"/>. Empty means <c>Span</c>, the 1.x default.
    /// </summary>
    internal static ChangeGranularity ParseChangeGranularity(string value, string connectionId)
    {
        if (string.IsNullOrWhiteSpace(value)) return ChangeGranularity.Span;
        var trimmed = value.Trim();
        // Names only: Enum.TryParse also accepts "0" and "1", which would read as a typo fixed.
        if (!char.IsDigit(trimmed[0]) && trimmed[0] != '-' &&
            Enum.TryParse<ChangeGranularity>(trimmed, ignoreCase: true, out var granularity) &&
            Enum.IsDefined(typeof(ChangeGranularity), granularity))
            return granularity;

        throw new InvalidOperationException(
            $"Connection '{connectionId}' has DefaultChangeGranularity '{value}'. Expected Span or Minimal.");
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
