namespace OfficeAgent.Mcp;

/// <summary>
/// Configuration for the OfficeAgent MCP server, bound from the <c>OfficeAgent</c>
/// section (appsettings, environment variables with the <c>OfficeAgent__</c> prefix,
/// or command line). Connections and credentials are host configuration. Editing
/// uses opaque ids; when registration is enabled, source-addressed tools also accept
/// provider-relative paths, SharePoint URLs, <c>driveId/itemId</c> pairs, iManage, Google
/// Drive or Box ids and links, or blob names and object keys.
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
    /// Gets or sets whether <c>register_document</c>, <c>remove_document</c>,
    /// <c>open_document</c>, <c>edit_document</c>, and <c>list_connections</c> are
    /// exposed, letting agents discover connections and stage existing documents.
    /// Defaults to <see langword="true"/>: an MCP client usually has no other channel to
    /// stage ids, unlike an in-process host. Set to <see langword="false"/> to pin agents
    /// to ids the host hands out by other means. Authoring brand-new documents is a
    /// separate opt-in - see <see cref="AllowCreation"/>.
    /// </summary>
    public bool AllowRegistration { get; set; } = true;

    /// <summary>
    /// Gets or sets whether <c>create_document</c> is exposed when at least one
    /// connection allows a format for which a blank-document factory is registered
    /// (currently <c>.docx</c>, <c>.pptx</c>, or <c>.xlsx</c>); SharePoint, iManage, Google Drive
    /// and Box connections additionally require an explicit creation folder. Defaults to <see langword="false"/>: letting
    /// an agent author new files under a connection root is a capability a host should
    /// choose, and an upgrade must not hand it to deployments that never asked for it.
    /// Turning it on adds <c>create_document</c> and, when registration is disabled,
    /// <c>list_connections</c> for creation discovery. Registration and composite tools
    /// remain unchanged.
    /// </summary>
    public bool AllowCreation { get; set; }

    /// <summary>
    /// Gets or sets whether the inline-content tools are exposed:
    /// <c>create_document_content</c>, <c>inspect_document_content</c>, and
    /// <c>edit_document_content</c>, which take the document as base64 and return the
    /// edited document the same way. They need no connection, so this is what makes the
    /// server useful with no storage configured at all - for a host whose documents arrive
    /// as attachments, or one that has nowhere to put a filesystem root.
    /// </summary>
    /// <remarks>
    /// Defaults to <see langword="false"/>. These tools move the whole package through the
    /// model's context in both directions, which costs tokens in proportion to file size
    /// and puts the complete file - not just the text the inspect tools already return - in
    /// front of the model provider. A deployment that configured a connection so documents
    /// would never travel that way should leave this off; one that turns it on should size
    /// its documents accordingly.
    /// </remarks>
    public bool AllowInlineContent { get; set; }

    /// <summary>
    /// Gets or sets the id of a session connection whose documents the server holds in
    /// memory, or empty for none. Setting it - <c>"session"</c> is the conventional value -
    /// gives the agent a connection it can create documents in and edit by opaque id with
    /// no storage configured, and adds <c>import_document_content</c> and
    /// <c>export_document_content</c> to move bytes in and out.
    /// </summary>
    /// <remarks>
    /// <para>
    /// This is the multi-step counterpart to <see cref="AllowInlineContent"/>. Passing a
    /// document inline costs the whole file in context each way and requires the model to
    /// reproduce it exactly to make a second edit - which it does not do reliably beyond a
    /// couple of kilobytes. Against a session connection the model passes a short id and
    /// the bytes never leave this process, so a ten-step edit costs no more context than a
    /// one-step edit.
    /// </para>
    /// <para>
    /// Nothing is persisted. The documents live for the life of the server process; a host
    /// that needs the result takes it with <c>export_document_content</c> or configures
    /// real storage. Creating documents in it still requires <see cref="AllowCreation"/>.
    /// </para>
    /// </remarks>
    public string EphemeralConnectionId { get; set; } = string.Empty;

    /// <summary>Gets or sets the total the session connection may hold at once. Defaults to 100 MB.</summary>
    public long EphemeralMaximumTotalBytes { get; set; } = 100L * 1024 * 1024;

    /// <summary>
    /// Gets or sets whether workbooks previewed or committed in the session connection have
    /// their formulas calculated, as <see cref="FileSystemConnectionOptions.CalculateFormulas"/>
    /// does for a folder. It applies to the zero-configuration session connection too. Defaults
    /// to <see langword="false"/>.
    /// </summary>
    public bool EphemeralCalculateFormulas { get; set; }

    /// <summary>
    /// Gets or sets whether a plan whose calculation produces a new formula error or a circular
    /// reference is refused in the session connection with <c>formula-error</c>. Applies only
    /// with <see cref="EphemeralCalculateFormulas"/>. Defaults to <see langword="false"/>.
    /// </summary>
    public bool EphemeralRejectFormulaErrors { get; set; }

    /// <summary>
    /// Gets or sets whether every tool parameter is listed as required, including the
    /// parameters that have a default. Defaults to <see langword="true"/>, which OpenAI and
    /// Azure OpenAI strict function calling require.
    /// </summary>
    /// <remarks>
    /// Set it to <see langword="false"/> for Microsoft Copilot Studio. Copilot Studio asks the
    /// user for every required input the model leaves out and cannot send an empty string,
    /// so a required <c>newName</c> whose default is empty stops <c>apply_plan</c> from ever
    /// running. With this off, parameters that have a default are optional. See
    /// <c>OfficeAgentToolsOptions.StrictToolSchemas</c>.
    /// </remarks>
    public bool StrictToolSchemas { get; set; } = true;

    /// <summary>
    /// Gets or sets whether tracked changes and comments carry the signed-in user's display
    /// name on a SharePoint connection using On-Behalf-Of. The plan's author stays on the
    /// receipt. Off by default, because <c>OfficeAgent</c> is the documented default author
    /// in 1.x; without a signed-in user nothing changes.
    /// </summary>
    public bool RevisionAuthorFromActor { get; set; }

    /// <summary>
    /// Gets or sets whether <c>preview_plan</c>, <c>apply_plan</c> and <c>edit_document</c>
    /// refuse a plan that does not carry the snapshot from inspection, with
    /// <c>snapshot-required</c>. Defaults to <see langword="false"/>.
    /// </summary>
    /// <remarks>
    /// A plan bound to its snapshot cannot be applied twice: once the first attempt is saved,
    /// the same plan is refused as stale, so a retried call cannot, for example, append a table
    /// row a second time. Models do not always copy the snapshot, so turn this on when an
    /// agent edits documents people rely on. See <c>OfficeAgentTools.RequirePlanSnapshot</c>.
    /// </remarks>
    public bool RequirePlanSnapshot { get; set; }

    /// <summary>
    /// Gets or sets the Office template locations: the personal and workgroup template folders
    /// Word, PowerPoint and Excel show under File > New. Off by default.
    /// </summary>
    public OfficeTemplateOptions OfficeTemplates { get; set; } = new();

    /// <summary>Gets or sets the filesystem connections to expose.</summary>
    public IList<FileSystemConnectionOptions> FileSystemConnections { get; set; } =
        new List<FileSystemConnectionOptions>();

    /// <summary>Gets or sets the SharePoint connections to expose.</summary>
    public IList<SharePointConnectionOptions> SharePointConnections { get; set; } =
        new List<SharePointConnectionOptions>();

    /// <summary>Gets or sets the iManage Work connections to expose.</summary>
    public IList<IManageConnectionOptions> IManageConnections { get; set; } =
        new List<IManageConnectionOptions>();

    /// <summary>Gets or sets the Google Drive connections to expose.</summary>
    public IList<GoogleDriveConnectionOptions> GoogleDriveConnections { get; set; } =
        new List<GoogleDriveConnectionOptions>();

    /// <summary>Gets or sets the Box connections to expose.</summary>
    public IList<BoxConnectionOptions> BoxConnections { get; set; } =
        new List<BoxConnectionOptions>();

    /// <summary>Gets or sets the Azure Blob Storage connections to expose.</summary>
    public IList<AzureBlobConnectionOptions> AzureBlobConnections { get; set; } =
        new List<AzureBlobConnectionOptions>();

    /// <summary>Gets or sets the Amazon S3 connections to expose.</summary>
    public IList<AmazonS3ConnectionOptions> AmazonS3Connections { get; set; } =
        new List<AmazonS3ConnectionOptions>();
}

/// <summary>One rooted filesystem connection.</summary>
public sealed class FileSystemConnectionOptions
{
    /// <summary>Gets or sets the connection id agents address documents under.</summary>
    public string ConnectionId { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the root directory; registrations must stay under it, and
    /// <c>create_document</c> writes new documents directly into it.
    /// </summary>
    public string RootPath { get; set; } = string.Empty;

    /// <summary>Gets or sets the maximum accepted document size in bytes.</summary>
    public long MaximumBytes { get; set; } = 100 * 1024 * 1024;

    /// <summary>Gets or sets the allowed document extensions.</summary>
    public IList<string> AllowedExtensions { get; set; } = new List<string> { ".docx" };

    /// <summary>
    /// Gets or sets the change mode used when a plan operation does not state one:
    /// <c>Tracked</c> (the default) or <c>Direct</c>. Set <c>Direct</c> for a connection
    /// over generated or machine-owned documents, or one serving <c>.pptx</c> - a deck has
    /// no redline vocabulary and refuses Tracked outright.
    /// </summary>
    public string DefaultChangeMode { get; set; } = "Tracked";

    /// <summary>
    /// Gets or sets how much of the matched text a tracked <c>changeText</c> that does not
    /// state a granularity marks: <c>Span</c> (the default in 1.x) marks the whole matched
    /// text, <c>Minimal</c> only the words that differ.
    /// </summary>
    public string DefaultChangeGranularity { get; set; } = "Span";

    /// <summary>
    /// Gets or sets whether workbooks previewed or committed through this connection have
    /// their formulas calculated with the bundled engine, so the saved file carries current
    /// values. Only allowlisted functions are calculated; <c>describe_capabilities</c> lists
    /// them, and the tool result names every formula left uncalculated and why. Defaults to
    /// <see langword="false"/>.
    /// </summary>
    public bool CalculateFormulas { get; set; }

    /// <summary>
    /// Gets or sets whether a plan whose calculation produces a new formula error or a
    /// circular reference is refused with <c>formula-error</c> rather than saved with the
    /// errors reported. Applies only with <see cref="CalculateFormulas"/>. Defaults to
    /// <see langword="false"/>.
    /// </summary>
    public bool RejectFormulaErrors { get; set; }

    /// <summary>
    /// Gets or sets the folder, relative to <see cref="RootPath"/>, that holds this connection's
    /// templates, such as <c>Templates</c>. <c>list_documents</c> with templates=true lists it and
    /// <c>find_templates</c> searches it. Empty for none.
    /// </summary>
    public string TemplatesFolder { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets what kind of template location the connection is: <c>Connection</c> (the
    /// default when <see cref="TemplatesFolder"/> is set), <c>Personal</c>, <c>Workgroup</c> or
    /// <c>Organization</c>. Without a <see cref="TemplatesFolder"/>, any of these makes the whole
    /// root a template location. Bound as a string so a typo names itself at startup.
    /// </summary>
    public string TemplateLocation { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets whether the connection refuses every save, creation and registration
    /// removal, and writes no registration index under the root. Use it for a templates folder
    /// people share. Defaults to <see langword="false"/>.
    /// </summary>
    public bool ReadOnly { get; set; }
}

/// <summary>
/// The Office template locations the server can add as read-only filesystem connections.
/// </summary>
/// <remarks>
/// On Windows the personal folder is the <c>PersonalTemplates</c> value Word, PowerPoint and
/// Excel keep under <c>HKCU\Software\Microsoft\Office\16.0\&lt;App&gt;\Options</c>, falling back to
/// <c>Documents\Custom Office Templates</c>; the workgroup folder is the <c>SharedTemplates</c>
/// value under <c>HKCU\Software\Policies\Microsoft\Office\16.0\Common\General</c>, then
/// <c>HKCU\Software\Microsoft\Office\16.0\Common\General</c>. On macOS the personal folder is
/// <c>~/Library/Group Containers/UBF8T346G9.Office/User Content/Templates</c>. A folder that does
/// not exist is skipped with a line on stderr. Microsoft's online templates, third-party template
/// managers and the SharePoint organization assets library are not discovered.
/// </remarks>
public sealed class OfficeTemplateOptions
{
    /// <summary>
    /// Gets or sets whether the server adds the <c>office-personal-templates</c> and
    /// <c>office-workgroup-templates</c> connections. Defaults to <see langword="false"/>.
    /// </summary>
    public bool Enabled { get; set; }
}

/// <summary>
/// The settings every storage-provider connection shares: its id, size and extension limits,
/// editing defaults, formula calculation, where registrations are kept, and its templates folder.
/// </summary>
/// <remarks>
/// The change mode and granularity are bound as strings, as on the filesystem and SharePoint
/// connections: the configuration binder drops a collection element it cannot bind, so a typo
/// bound to the enum would remove the whole connection instead of naming the bad value.
/// </remarks>
public abstract class StorageConnectionOptions
{
    /// <summary>Gets or sets the connection id agents address documents under.</summary>
    public string ConnectionId { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the path of a JSON registration index. When set, registrations
    /// survive restarts; when empty, they live in process memory.
    /// </summary>
    public string RegistrationIndexPath { get; set; } = string.Empty;

    /// <summary>Gets or sets the maximum accepted document size in bytes.</summary>
    public long MaximumBytes { get; set; } = 100 * 1024 * 1024;

    /// <summary>Gets or sets the allowed document extensions.</summary>
    public IList<string> AllowedExtensions { get; set; } = new List<string> { ".docx" };

    /// <summary>
    /// Gets or sets the change mode used when a plan operation does not state one:
    /// <c>Tracked</c> (the default) or <c>Direct</c>.
    /// </summary>
    public string DefaultChangeMode { get; set; } = "Tracked";

    /// <summary>
    /// Gets or sets how much of the matched text a tracked <c>changeText</c> that does not
    /// state a granularity marks: <c>Span</c> (the default in 1.x) or <c>Minimal</c>.
    /// </summary>
    public string DefaultChangeGranularity { get; set; } = "Span";

    /// <summary>
    /// Gets or sets whether workbooks previewed or committed through this connection have
    /// their formulas calculated, as <see cref="FileSystemConnectionOptions.CalculateFormulas"/>
    /// does for a folder. Defaults to <see langword="false"/>.
    /// </summary>
    public bool CalculateFormulas { get; set; }

    /// <summary>
    /// Gets or sets whether a plan whose calculation produces a new formula error or a
    /// circular reference is refused with <c>formula-error</c>. Applies only with
    /// <see cref="CalculateFormulas"/>. Defaults to <see langword="false"/>.
    /// </summary>
    public bool RejectFormulaErrors { get; set; }

    /// <summary>
    /// Gets or sets the connection's templates folder, which a templates listing reads: a Drive
    /// folder id, a numeric Box folder id, an iManage folder id (<c>LIBRARY!number</c>), or for
    /// Azure Blob and Amazon S3 a folder relative to the prefix, such as <c>templates</c>. Leave
    /// empty for a connection without templates.
    /// </summary>
    public string TemplatesFolder { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets where <see cref="TemplatesFolder"/> sits: <c>Connection</c>, <c>Personal</c>,
    /// <c>Workgroup</c> or <c>Organization</c>. Empty means <c>Connection</c> when a templates
    /// folder is set. Bound as a string for the reason the change mode is.
    /// </summary>
    public string TemplateLocation { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets whether the connection refuses every save and creation before the service
    /// is called. Registrations still work and persist as usual. Use it for a templates folder
    /// people share, beside a writable connection for the documents made from it; it cannot
    /// also name a creation folder. Defaults to <see langword="false"/>.
    /// </summary>
    public bool ReadOnly { get; set; }
}

/// <summary>
/// One iManage Work connection. Documents are registered by their iManage id
/// (<c>LIBRARY!number.version</c>) or an iManage Work link on <see cref="ServerUrl"/>.
/// </summary>
/// <remarks>
/// <see cref="ClientSecret"/>, <see cref="Password"/> and <see cref="RefreshToken"/> are
/// secrets. Supply them through <c>OfficeAgent__IManageConnections__N__*</c> environment
/// variables or a secret store, not appsettings. They are sent only to the iManage token
/// endpoint and never appear in tool results, connection descriptors or error messages.
/// </remarks>
public sealed class IManageConnectionOptions : StorageConnectionOptions
{
    /// <summary>
    /// Gets or sets the iManage server URL, an absolute https URL such as
    /// <c>https://cloudimanage.com</c> or your own Work server. Registration links must be on
    /// this host, and tokens are requested from it.
    /// </summary>
    public string ServerUrl { get; set; } = string.Empty;

    /// <summary>Gets or sets the Work REST API base path on <see cref="ServerUrl"/>.</summary>
    public string ApiBasePath { get; set; } = "/work/api/v2";

    /// <summary>
    /// Gets or sets the iManage customer id that every Work API v2 path names. An administrator
    /// can read it from the server's <c>/api</c> response.
    /// </summary>
    public string CustomerId { get; set; } = string.Empty;

    /// <summary>Gets or sets the OAuth2 token endpoint path on <see cref="ServerUrl"/>.</summary>
    public string TokenPath { get; set; } = "/auth/oauth2/token";

    /// <summary>
    /// Gets or sets the OAuth2 grant: <c>Password</c> (the default; a dedicated service
    /// account), <c>RefreshToken</c>, or <c>ClientCredentials</c> (only where your iManage
    /// environment enables it).
    /// </summary>
    public string GrantType { get; set; } = "Password";

    /// <summary>Gets or sets the application's client id from iManage Control Center.</summary>
    public string ClientId { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the application's client secret. Required for <c>ClientCredentials</c>;
    /// leave empty for a public client. Supply it through
    /// <c>OfficeAgent__IManageConnections__N__ClientSecret</c> or a secret store.
    /// </summary>
    public string ClientSecret { get; set; } = string.Empty;

    /// <summary>Gets or sets the service account's user id, for the <c>Password</c> grant.</summary>
    public string Username { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the service account's password, for the <c>Password</c> grant. Supply it
    /// through <c>OfficeAgent__IManageConnections__N__Password</c> or a secret store.
    /// </summary>
    public string Password { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the initial refresh token, for the <c>RefreshToken</c> grant. The rotated
    /// token is kept in memory only. Supply it through
    /// <c>OfficeAgent__IManageConnections__N__RefreshToken</c> or a secret store.
    /// </summary>
    public string RefreshToken { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the folder or workspace id, <c>LIBRARY!number</c>, that receives new
    /// documents. Leave empty to disable <c>create_document</c> for this connection.
    /// </summary>
    public string CreationFolderId { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the document class given to new documents, for libraries that require one.
    /// Leave empty to let the library apply its default.
    /// </summary>
    public string CreationDocumentClass { get; set; } = string.Empty;

    /// <summary>Gets or sets the comment recorded on each version and document OfficeAgent saves.</summary>
    public string SaveComment { get; set; } = "Saved by OfficeAgent.";

    /// <summary>
    /// Gets or sets whether a save locks the document while it checks the latest version and
    /// uploads. Defaults to <see langword="true"/>. Turn it off only for a server that does not
    /// accept document locks.
    /// </summary>
    public bool LockDuringSave { get; set; } = true;
}

/// <summary>
/// One Google Drive connection, authenticated as a Google service account. Documents are
/// registered by a Drive file id or a Drive or Docs URL.
/// </summary>
/// <remarks>
/// The service account key is a secret. Point <see cref="KeyFilePath"/> at the key file, or
/// supply its contents through <c>OfficeAgent__GoogleDriveConnections__N__KeyJson</c> from a
/// secret store; never write the key into appsettings or a manifest.
/// </remarks>
public sealed class GoogleDriveConnectionOptions : StorageConnectionOptions
{
    /// <summary>Gets or sets the path of the service account key file downloaded from Google Cloud.</summary>
    public string KeyFilePath { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the service account key file's JSON contents, the alternative to
    /// <see cref="KeyFilePath"/>. Supply it from the environment or a secret store.
    /// </summary>
    public string KeyJson { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the Google Workspace user to act as, through domain-wide delegation. Leave
    /// empty for the service account to act as itself, which is the narrower setup.
    /// </summary>
    public string Subject { get; set; } = string.Empty;

    /// <summary>Gets or sets the requested OAuth scope. Defaults to full Drive access.</summary>
    public string Scope { get; set; } = "https://www.googleapis.com/auth/drive";

    /// <summary>Gets or sets the Drive API v3 base URL.</summary>
    public string ApiBaseUrl { get; set; } = "https://www.googleapis.com/drive/v3";

    /// <summary>Gets or sets the Drive API v3 media upload base URL.</summary>
    public string UploadBaseUrl { get; set; } = "https://www.googleapis.com/upload/drive/v3";

    /// <summary>
    /// Gets or sets the Drive folder id that receives new documents. Leave empty to disable
    /// <c>create_document</c> for this connection.
    /// </summary>
    public string CreationFolderId { get; set; } = string.Empty;
}

/// <summary>
/// One Box connection, authenticated with the client credentials grant. Documents are
/// registered by a Box file id or a Box file URL.
/// </summary>
/// <remarks>
/// <see cref="ClientSecret"/> is a secret. Supply it through
/// <c>OfficeAgent__BoxConnections__N__ClientSecret</c> or a secret store, not appsettings.
/// </remarks>
public sealed class BoxConnectionOptions : StorageConnectionOptions
{
    /// <summary>Initializes the options with the 50 MB direct-upload ceiling Box recommends as the size limit.</summary>
    public BoxConnectionOptions() => MaximumBytes = 50 * 1024 * 1024;

    /// <summary>Gets or sets the Box app's client id.</summary>
    public string ClientId { get; set; } = string.Empty;

    /// <summary>Gets or sets the Box app's client secret.</summary>
    public string ClientSecret { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets who the token acts as: <c>Enterprise</c> (the default; the app's service
    /// account) or <c>User</c> (a managed user or app user).
    /// </summary>
    public string SubjectType { get; set; } = "Enterprise";

    /// <summary>Gets or sets the enterprise id, or the user id for <c>User</c>.</summary>
    public string SubjectId { get; set; } = string.Empty;

    /// <summary>Gets or sets the Box API base URL.</summary>
    public string ApiBaseUrl { get; set; } = "https://api.box.com/2.0";

    /// <summary>Gets or sets the Box upload API base URL.</summary>
    public string UploadBaseUrl { get; set; } = "https://upload.box.com/api/2.0";

    /// <summary>
    /// Gets or sets the numeric Box folder id that receives new documents. Leave empty to
    /// disable <c>create_document</c> for this connection.
    /// </summary>
    public string CreationFolderId { get; set; } = string.Empty;
}

/// <summary>
/// One Azure Blob Storage connection over one container and an optional prefix. Documents are
/// registered by blob names relative to the prefix, and new documents are created under it.
/// </summary>
/// <remarks>
/// With <see cref="ServiceUri"/> the connection authenticates with <c>DefaultAzureCredential</c>,
/// such as a managed identity. <see cref="ConnectionString"/> carries the account key: supply it
/// through <c>OfficeAgent__AzureBlobConnections__N__ConnectionString</c> or a secret store.
/// </remarks>
public sealed class AzureBlobConnectionOptions : StorageConnectionOptions
{
    /// <summary>Gets or sets the Blob service endpoint, such as <c>https://contoso.blob.core.windows.net</c>.</summary>
    public string ServiceUri { get; set; } = string.Empty;

    /// <summary>Gets or sets a storage connection string, the alternative to <see cref="ServiceUri"/>.</summary>
    public string ConnectionString { get; set; } = string.Empty;

    /// <summary>Gets or sets the container this connection is confined to.</summary>
    public string ContainerName { get; set; } = string.Empty;

    /// <summary>Gets or sets an optional blob-name prefix, such as <c>contracts/</c>.</summary>
    public string Prefix { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the client id of a user-assigned managed identity for
    /// <c>DefaultAzureCredential</c>. Leave empty for a system-assigned identity.
    /// </summary>
    public string ManagedIdentityClientId { get; set; } = string.Empty;
}

/// <summary>
/// One Amazon S3 connection over one bucket and an optional key prefix. Credentials come from
/// the AWS SDK's default chain, such as an instance or task role; none are configured here.
/// </summary>
public sealed class AmazonS3ConnectionOptions : StorageConnectionOptions
{
    /// <summary>Gets or sets the bucket this connection is confined to.</summary>
    public string BucketName { get; set; } = string.Empty;

    /// <summary>Gets or sets an optional key prefix, such as <c>contracts/</c>.</summary>
    public string KeyPrefix { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the AWS region, such as <c>eu-west-1</c>. Empty takes the region from the AWS
    /// SDK's default chain.
    /// </summary>
    public string Region { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets an endpoint for an S3-compatible store, such as MinIO. Empty uses Amazon S3.
    /// The store must support conditional writes.
    /// </summary>
    public string ServiceUrl { get; set; } = string.Empty;

    /// <summary>Gets or sets whether requests use path-style addressing, which most S3-compatible stores need.</summary>
    public bool ForcePathStyle { get; set; }
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
    /// Gets or sets the Graph application scope requested by the app-only
    /// client-credentials flow. Change this together with <see cref="GraphBaseUrl"/>
    /// and <see cref="LoginAuthority"/> for a sovereign cloud.
    /// </summary>
    public string AppOnlyScope { get; set; } = "https://graph.microsoft.com/.default";

    /// <summary>
    /// Gets or sets the path of a JSON registration index. When set, registrations
    /// survive restarts; when empty, they live in process memory.
    /// </summary>
    public string RegistrationIndexPath { get; set; } = string.Empty;

    /// <summary>Gets or sets the maximum accepted document size in bytes.</summary>
    public long MaximumBytes { get; set; } = 100 * 1024 * 1024;

    /// <summary>Gets or sets the allowed document extensions.</summary>
    public IList<string> AllowedExtensions { get; set; } = new List<string> { ".docx" };

    /// <summary>
    /// Gets or sets the Graph drive id used for <c>create_document</c>. Set together
    /// with <see cref="CreationFolderItemId"/>; registration is still allowed across
    /// every drive reachable by this connection.
    /// </summary>
    public string CreationDriveId { get; set; } = string.Empty;

    /// <summary>Gets or sets the destination folder's Graph drive-item id for new documents.</summary>
    public string CreationFolderItemId { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the change mode used when a plan operation does not state one:
    /// <c>Tracked</c> (the default) or <c>Direct</c>. Set <c>Direct</c> for a library of
    /// generated or machine-owned documents, or one serving <c>.pptx</c> - a deck has no
    /// redline vocabulary and refuses Tracked outright.
    /// </summary>
    public string DefaultChangeMode { get; set; } = "Tracked";

    /// <summary>
    /// Gets or sets how much of the matched text a tracked <c>changeText</c> that does not
    /// state a granularity marks: <c>Span</c> (the default in 1.x) marks the whole matched
    /// text, <c>Minimal</c> only the words that differ.
    /// </summary>
    public string DefaultChangeGranularity { get; set; } = "Span";

    /// <summary>
    /// Gets or sets whether workbooks previewed or committed through this connection have
    /// their formulas calculated with the bundled engine, so the saved file carries current
    /// values. Only allowlisted functions are calculated; <c>describe_capabilities</c> lists
    /// them, and the tool result names every formula left uncalculated and why. Defaults to
    /// <see langword="false"/>.
    /// </summary>
    public bool CalculateFormulas { get; set; }

    /// <summary>
    /// Gets or sets whether a plan whose calculation produces a new formula error or a
    /// circular reference is refused with <c>formula-error</c> rather than saved with the
    /// errors reported. Applies only with <see cref="CalculateFormulas"/>. Defaults to
    /// <see langword="false"/>.
    /// </summary>
    public bool RejectFormulaErrors { get; set; }

    /// <summary>
    /// Gets or sets the folder that holds this connection's templates: a SharePoint/OneDrive
    /// folder URL or a <c>driveId/itemId</c> pair. An organization assets library is not found
    /// automatically; name its folder here and set <see cref="TemplateLocation"/> to
    /// <c>Organization</c>. Empty for none.
    /// </summary>
    public string TemplatesFolder { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets what kind of template location <see cref="TemplatesFolder"/> is:
    /// <c>Connection</c> (the default when it is set), <c>Personal</c>, <c>Workgroup</c> or
    /// <c>Organization</c>.
    /// </summary>
    public string TemplateLocation { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets whether the connection refuses every save and creation before Graph is
    /// called, whatever the configured identity may write. Registrations still work and persist
    /// as usual. Use it for a templates library people share, beside a writable connection for
    /// the documents made from it; it cannot also name <see cref="CreationDriveId"/> or
    /// <see cref="CreationFolderItemId"/>. Defaults to <see langword="false"/>.
    /// </summary>
    public bool ReadOnly { get; set; }
}
