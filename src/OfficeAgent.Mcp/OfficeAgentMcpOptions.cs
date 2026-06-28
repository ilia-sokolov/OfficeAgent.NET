namespace OfficeAgent.Mcp;

/// <summary>
/// Configuration for the OfficeAgent MCP server, bound from the <c>OfficeAgent</c>
/// section (appsettings, environment variables with the <c>OfficeAgent__</c> prefix,
/// or command line). Connections are host configuration; the agent only ever sees
/// connection ids and opaque document ids.
/// </summary>
public sealed class OfficeAgentMcpOptions
{
    /// <summary>The configuration section the server binds.</summary>
    public const string SectionName = "OfficeAgent";

    /// <summary>
    /// Gets or sets the transport: <c>http</c> (default; streamable HTTP for cloud or
    /// shared hosting) or <c>stdio</c> (local child-process hosting under an MCP
    /// client). The <c>--stdio</c> command-line flag forces stdio.
    /// </summary>
    public string Transport { get; set; } = "http";

    /// <summary>
    /// Gets or sets whether the <c>register_document</c> / <c>remove_document</c> /
    /// <c>list_connections</c> tools are exposed, letting agents discover the configured
    /// connections, register documents with them, and remove registrations themselves.
    /// Defaults to <see langword="true"/>: an MCP client usually has no other channel to
    /// stage ids, unlike an in-process host. Set to <see langword="false"/> to pin agents
    /// to ids the host hands out by other means.
    /// </summary>
    public bool AllowRegistration { get; set; } = true;

    /// <summary>Gets or sets the filesystem connections to expose.</summary>
    public IList<FileSystemConnectionOptions> FileSystemConnections { get; set; } =
        new List<FileSystemConnectionOptions>();

    /// <summary>Gets or sets the SharePoint connections to expose.</summary>
    public IList<SharePointConnectionOptions> SharePointConnections { get; set; } =
        new List<SharePointConnectionOptions>();
}

/// <summary>One rooted filesystem connection.</summary>
public sealed class FileSystemConnectionOptions
{
    /// <summary>Gets or sets the connection id agents address documents under.</summary>
    public string ConnectionId { get; set; } = string.Empty;

    /// <summary>Gets or sets the root directory; registrations must stay under it.</summary>
    public string RootPath { get; set; } = string.Empty;

    /// <summary>Gets or sets the maximum accepted document size in bytes.</summary>
    public long MaximumBytes { get; set; } = 100 * 1024 * 1024;

    /// <summary>Gets or sets the allowed document extensions.</summary>
    public IList<string> AllowedExtensions { get; set; } = new List<string> { ".docx" };
}

/// <summary>One SharePoint document-library connection.</summary>
public sealed class SharePointConnectionOptions
{
    /// <summary>Gets or sets the connection id agents address documents under.</summary>
    public string ConnectionId { get; set; } = string.Empty;

    /// <summary>Gets or sets the Graph endpoint base. Override for sovereign clouds.</summary>
    public string GraphBaseUrl { get; set; } = "https://graph.microsoft.com/v1.0";

    /// <summary>Gets or sets the Entra login authority base. Override for sovereign clouds.</summary>
    public string LoginAuthority { get; set; } = "https://login.microsoftonline.com";

    /// <summary>
    /// Gets or sets how the connection authenticates to Graph:
    /// <c>onBehalfOf</c> exchanges the signed-in user's inbound token so the provider
    /// acts as that user (hosted HTTP only; respects the user's SharePoint
    /// permissions); <c>appOnly</c> (the default) authenticates as the app itself - a
    /// shared app identity, no user - for unattended scenarios.
    /// </summary>
    public string AuthMode { get; set; } = "appOnly";

    /// <summary>Gets or sets the Entra tenant id for the client-credentials or On-Behalf-Of flow.</summary>
    public string TenantId { get; set; } = string.Empty;

    /// <summary>Gets or sets the app registration's client id.</summary>
    public string ClientId { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the app registration's client secret. For <c>onBehalfOf</c> this is
    /// the middle-tier API app's secret. Prefer supplying it via the
    /// <c>OfficeAgent__SharePointConnections__N__ClientSecret</c> environment variable
    /// or a secret store over committing it to appsettings.
    /// </summary>
    public string ClientSecret { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the downstream Graph scope requested by the On-Behalf-Of exchange.
    /// The default is the Graph application scope, which resolves the user's consented
    /// delegated permissions.
    /// </summary>
    public string OnBehalfOfScope { get; set; } = "https://graph.microsoft.com/.default";

    /// <summary>
    /// Gets or sets the path of a JSON registration index. When set, registrations
    /// survive restarts; when empty, they live in process memory.
    /// </summary>
    public string RegistrationIndexPath { get; set; } = string.Empty;

    /// <summary>Gets or sets the maximum accepted document size in bytes.</summary>
    public long MaximumBytes { get; set; } = 100 * 1024 * 1024;

    /// <summary>Gets or sets the allowed document extensions.</summary>
    public IList<string> AllowedExtensions { get; set; } = new List<string> { ".docx" };
}
