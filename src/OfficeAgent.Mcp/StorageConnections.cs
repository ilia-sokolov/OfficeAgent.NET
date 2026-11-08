using System.Text.RegularExpressions;
using Amazon.S3;
using Azure.Storage.Blobs;
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.AmazonS3;
using OfficeAgent.AzureBlob;
using OfficeAgent.Box;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.GoogleDrive;
using OfficeAgent.IManage;

namespace OfficeAgent.Mcp;

/// <summary>
/// What the agent is told about one kind of storage: the provider name, how a
/// <c>register_document</c> source is written, and where <c>create_document</c> puts a new file.
/// </summary>
/// <param name="Provider">The provider name <c>list_connections</c> reports.</param>
/// <param name="Label">The name used in tool descriptions, such as <c>Google Drive</c>.</param>
/// <param name="Source">The source, as the connection inventory in the server instructions says it.</param>
/// <param name="ToolSource">The source, as the <c>register_document</c> description says it.</param>
/// <param name="Creation">Where <c>create_document</c> writes, completing "create_document ...".</param>
internal sealed record StorageKind(string Provider, string Label, string Source, string ToolSource, string Creation)
{
    public static readonly StorageKind FileSystem = new(
        "filesystem", "filesystem",
        "a path under this connection's root",
        "a path under its root",
        "writes new documents into this connection's root");

    public static readonly StorageKind SharePoint = new(
        "sharepoint", "SharePoint",
        "a SharePoint/OneDrive URL or a \"driveId/itemId\" pair",
        "the document's SharePoint/OneDrive URL (e.g. 'https://contoso.sharepoint.com/:w:/s/…') or a 'driveId/itemId' pair (e.g. 'b!9a3f…/01ABCDEF')",
        "writes new documents into this connection's configured folder");

    public static readonly StorageKind IManage = new(
        IManageDocumentProvider.ProviderName, "iManage",
        "an iManage document id such as \"ACTIVE!12345.1\" or an iManage Work link to the document",
        "the document id 'LIBRARY!number.version' (e.g. 'ACTIVE!12345.1') or an iManage Work link to it",
        "files new documents in this connection's configured folder");

    public static readonly StorageKind GoogleDrive = new(
        GoogleDriveDocumentProvider.ProviderName, "Google Drive",
        "a Google Drive file id or the file's Drive or Docs URL",
        "the file id or its Drive URL (e.g. 'https://drive.google.com/file/d/…/view')",
        "writes new documents into this connection's configured folder");

    public static readonly StorageKind Box = new(
        BoxDocumentProvider.ProviderName, "Box",
        "a Box file id or the file's Box URL (https://app.box.com/file/{id})",
        "the file id or its Box URL (e.g. 'https://app.box.com/file/123456')",
        "writes new documents into this connection's configured folder");

    public static readonly StorageKind AzureBlob = new(
        AzureBlobDocumentProvider.ProviderName, "Azure Blob",
        "a blob name relative to this connection's folder, such as \"acme.docx\" or \"2026/acme.docx\"",
        "the blob name relative to the connection's folder (e.g. 'acme.docx')",
        "writes new documents into this connection's folder");

    public static readonly StorageKind AmazonS3 = new(
        AmazonS3DocumentProvider.ProviderName, "Amazon S3",
        "an object key relative to this connection's folder, such as \"acme.docx\" or \"2026/acme.docx\"",
        "the object key relative to the connection's folder (e.g. 'acme.docx')",
        "writes new documents into this connection's folder");
}

/// <summary>
/// One configured storage connection as agents see it. It carries no credential, server URL,
/// container, bucket or prefix, so nothing in it can leak through a tool result.
/// </summary>
internal sealed record StorageConnection(
    StorageKind Kind,
    string ConnectionId,
    IReadOnlyCollection<string> AllowedExtensions,
    bool HasCreationTarget,
    string DefaultChangeMode,
    string DefaultChangeGranularity,
    bool CalculatesFormulas);

/// <summary>
/// Supplies the transports the provider connections would otherwise build themselves, so tests
/// can drive the real providers against in-memory fakes. Production never registers it.
/// </summary>
internal sealed class StorageTransports
{
    /// <summary>The HTTP client for an iManage, Google Drive or Box connection, by provider name.</summary>
    public Func<string, HttpClient>? Http { get; init; }

    /// <summary>The container client for an Azure Blob connection.</summary>
    public Func<AzureBlobConnectionOptions, BlobContainerClient>? AzureBlob { get; init; }

    /// <summary>The S3 client for an Amazon S3 connection.</summary>
    public Func<AmazonS3ConnectionOptions, IAmazonS3>? AmazonS3 { get; init; }
}

/// <summary>
/// The configured storage connections: what agents are told about them, the startup checks that
/// stop the server on bad configuration, and the provider each one becomes.
/// </summary>
/// <remarks>
/// Every message raised here names the connection and the setting, never a secret's value: a
/// configuration error is printed to the console and may be collected into logs.
/// </remarks>
internal static class StorageConnections
{
    private static readonly Regex IManageContainerId = new("^[A-Za-z0-9_-]+![0-9]{1,18}$", RegexOptions.CultureInvariant);

    /// <summary>Every configured connection except the session one, in configuration order.</summary>
    public static IReadOnlyList<StorageConnection> Describe(OfficeAgentMcpOptions options)
    {
        var connections = new List<StorageConnection>();
        connections.AddRange(options.FileSystemConnections.Select(c => new StorageConnection(
            StorageKind.FileSystem, c.ConnectionId, c.AllowedExtensions.ToArray(), HasCreationTarget: true,
            c.DefaultChangeMode, c.DefaultChangeGranularity, c.CalculateFormulas)));
        connections.AddRange(options.SharePointConnections.Select(c => new StorageConnection(
            StorageKind.SharePoint, c.ConnectionId, c.AllowedExtensions.ToArray(),
            !string.IsNullOrWhiteSpace(c.CreationDriveId) && !string.IsNullOrWhiteSpace(c.CreationFolderItemId),
            c.DefaultChangeMode, c.DefaultChangeGranularity, c.CalculateFormulas)));
        connections.AddRange(options.IManageConnections.Select(c => Of(StorageKind.IManage, c, HasFolder(c.CreationFolderId))));
        connections.AddRange(options.GoogleDriveConnections.Select(c => Of(StorageKind.GoogleDrive, c, HasFolder(c.CreationFolderId))));
        connections.AddRange(options.BoxConnections.Select(c => Of(StorageKind.Box, c, HasFolder(c.CreationFolderId))));
        // Object stores create under the connection's own container and prefix, so they always
        // have a destination.
        connections.AddRange(options.AzureBlobConnections.Select(c => Of(StorageKind.AzureBlob, c, true)));
        connections.AddRange(options.AmazonS3Connections.Select(c => Of(StorageKind.AmazonS3, c, true)));
        return connections;
    }

    /// <summary>
    /// The <c>register_document</c> source sentence for the kinds of storage actually configured,
    /// so an agent is never told about a source no connection accepts.
    /// </summary>
    public static string SourceDescription(OfficeAgentMcpOptions options)
    {
        var kinds = Describe(options).Select(c => c.Kind).Distinct().ToList();
        return "source is connection-specific: " +
               string.Join("; ", kinds.Select(k => $"for {Article(k.Label)} {k.Label} connection, {k.ToolSource}")) + ".";
    }

    private static string Article(string label) =>
        "aeiouAEIOU".IndexOf(label[0]) >= 0 ? "an" : "a";

    private static StorageConnection Of(StorageKind kind, StorageConnectionOptions c, bool hasCreationTarget) =>
        new(kind, c.ConnectionId, c.AllowedExtensions.ToArray(), hasCreationTarget,
            c.DefaultChangeMode, c.DefaultChangeGranularity, c.CalculateFormulas);

    private static bool HasFolder(string? folder) => !string.IsNullOrWhiteSpace(folder);

    // ── Startup checks ────────────────────────────────────────────────────────

    /// <summary>
    /// Checks every iManage, Google Drive, Box, Azure Blob and Amazon S3 connection, so a missing
    /// or malformed setting stops the server at startup and names itself.
    /// </summary>
    public static void Validate(OfficeAgentMcpOptions options)
    {
        Each(options.IManageConnections, nameof(OfficeAgentMcpOptions.IManageConnections), StorageKind.IManage, ValidateIManage);
        Each(options.GoogleDriveConnections, nameof(OfficeAgentMcpOptions.GoogleDriveConnections), StorageKind.GoogleDrive, ValidateGoogleDrive);
        Each(options.BoxConnections, nameof(OfficeAgentMcpOptions.BoxConnections), StorageKind.Box, ValidateBox);
        Each(options.AzureBlobConnections, nameof(OfficeAgentMcpOptions.AzureBlobConnections), StorageKind.AzureBlob, ValidateAzureBlob);
        Each(options.AmazonS3Connections, nameof(OfficeAgentMcpOptions.AmazonS3Connections), StorageKind.AmazonS3, ValidateAmazonS3);
    }

    /// <summary>Names one connection's settings in errors: its kind, id, and environment-variable prefix.</summary>
    private sealed class Setting
    {
        private readonly StorageKind _kind;
        private readonly string _connectionId;
        private readonly string _prefix;

        public Setting(StorageKind kind, string section, int index, string connectionId)
        {
            _kind = kind;
            _connectionId = connectionId;
            _prefix = $"OfficeAgent__{section}__{index}__";
        }

        /// <summary>The environment variable that sets <paramref name="name"/> on this connection.</summary>
        public string Variable(string name) => _prefix + name;

        public InvalidOperationException Error(string problem) =>
            new($"{_kind.Label} connection '{_connectionId}' {problem}");
    }

    private static void Each<T>(
        IList<T> connections, string section, StorageKind kind, Action<T, Setting> validate)
        where T : StorageConnectionOptions
    {
        for (var index = 0; index < connections.Count; index++)
        {
            var connection = connections[index];
            if (string.IsNullOrWhiteSpace(connection.ConnectionId))
                throw new InvalidOperationException(
                    $"{kind.Label} connection {index} has no ConnectionId. Set OfficeAgent__{section}__{index}__ConnectionId.");

            var setting = new Setting(kind, section, index, connection.ConnectionId);
            if (connection.MaximumBytes <= 0)
                throw setting.Error($"has MaximumBytes {connection.MaximumBytes}. It must be positive.");
            if (connection.AllowedExtensions.Count == 0)
                throw setting.Error($"allows no extension. Set {setting.Variable("AllowedExtensions__0")}, such as .docx.");
            OfficeAgentMcpServer.ParseChangeMode(connection.DefaultChangeMode, connection.ConnectionId);
            OfficeAgentMcpServer.ParseChangeGranularity(connection.DefaultChangeGranularity, connection.ConnectionId);
            validate(connection, setting);
        }
    }

    private static void ValidateIManage(IManageConnectionOptions c, Setting setting)
    {
        RequireUrl(c.ServerUrl, httpsOnly: true, "ServerUrl", "an absolute https URL such as https://cloudimanage.com", setting);
        Require(c.CustomerId, "CustomerId", "the customer id from the server's /api response", setting);
        Require(c.ClientId, "ClientId", "the application's client id from iManage Control Center", setting);
        Require(c.ApiBasePath, "ApiBasePath", "the Work REST API base path, /work/api/v2 by default", setting);
        Require(c.TokenPath, "TokenPath", "the OAuth2 token path, /auth/oauth2/token by default", setting);

        switch (ParseIManageGrant(c.GrantType, setting))
        {
            case IManageGrantType.Password:
                Require(c.Username, "Username", "the service account's user id for the Password grant", setting);
                RequireSecret(c.Password, "Password", "the Password grant", setting);
                break;
            case IManageGrantType.RefreshToken:
                RequireSecret(c.RefreshToken, "RefreshToken", "the RefreshToken grant", setting);
                break;
            case IManageGrantType.ClientCredentials:
                RequireSecret(c.ClientSecret, "ClientSecret", "the ClientCredentials grant", setting);
                break;
        }

        if (HasFolder(c.CreationFolderId) && !IManageContainerId.IsMatch(c.CreationFolderId.Trim()))
            throw setting.Error(
                "has a CreationFolderId that is not an iManage folder or workspace id. Use LIBRARY!number, such as ACTIVE!773.");
    }

    private static void ValidateGoogleDrive(GoogleDriveConnectionOptions c, Setting setting)
    {
        var hasFile = !string.IsNullOrWhiteSpace(c.KeyFilePath);
        var hasJson = !string.IsNullOrWhiteSpace(c.KeyJson);
        if (hasFile == hasJson)
            throw setting.Error(
                $"needs exactly one service account key: {setting.Variable("KeyFilePath")} naming the key file, " +
                $"or {setting.Variable("KeyJson")} from a secret store.");
        if (hasFile && !File.Exists(c.KeyFilePath.Trim()))
            throw setting.Error("has a KeyFilePath that names no file. Point it at the service account key file.");
        Require(c.Scope, "Scope", "the OAuth scope, https://www.googleapis.com/auth/drive by default", setting);
        RequireUrl(c.ApiBaseUrl, httpsOnly: true, "ApiBaseUrl", "an absolute https URL", setting);
        RequireUrl(c.UploadBaseUrl, httpsOnly: true, "UploadBaseUrl", "an absolute https URL", setting);
    }

    private static void ValidateBox(BoxConnectionOptions c, Setting setting)
    {
        Require(c.ClientId, "ClientId", "the Box app's client id", setting);
        RequireSecret(c.ClientSecret, "ClientSecret", "the client credentials grant", setting);
        ParseBoxSubject(c.SubjectType, setting);
        Require(c.SubjectId, "SubjectId", "the enterprise id, or the user id when SubjectType is User", setting);
        RequireUrl(c.ApiBaseUrl, httpsOnly: true, "ApiBaseUrl", "an absolute https URL", setting);
        RequireUrl(c.UploadBaseUrl, httpsOnly: true, "UploadBaseUrl", "an absolute https URL", setting);
        if (HasFolder(c.CreationFolderId) && !c.CreationFolderId.Trim().All(ch => ch is >= '0' and <= '9'))
            throw setting.Error("has a CreationFolderId that is not a numeric Box folder id.");
    }

    private static void ValidateAzureBlob(AzureBlobConnectionOptions c, Setting setting)
    {
        Require(c.ContainerName, "ContainerName", "the container this connection is confined to", setting);
        var hasUri = !string.IsNullOrWhiteSpace(c.ServiceUri);
        var hasConnectionString = !string.IsNullOrWhiteSpace(c.ConnectionString);
        if (hasUri == hasConnectionString)
            throw setting.Error(
                $"needs exactly one way to reach the account: {setting.Variable("ServiceUri")} (signs in with " +
                $"DefaultAzureCredential), or {setting.Variable("ConnectionString")} from a secret store.");
        if (hasUri)
            RequireUrl(c.ServiceUri, httpsOnly: false, "ServiceUri", "an absolute http(s) URL such as https://contoso.blob.core.windows.net", setting);
    }

    private static void ValidateAmazonS3(AmazonS3ConnectionOptions c, Setting setting)
    {
        Require(c.BucketName, "BucketName", "the bucket this connection is confined to", setting);
        if (!string.IsNullOrWhiteSpace(c.ServiceUrl))
            RequireUrl(c.ServiceUrl, httpsOnly: false, "ServiceUrl", "an absolute http(s) URL for an S3-compatible store", setting);
    }

    private static void Require(string? value, string name, string meaning, Setting setting)
    {
        if (string.IsNullOrWhiteSpace(value))
            throw setting.Error($"needs {name}, {meaning}. Set {setting.Variable(name)}.");
    }

    /// <summary>A missing secret is named by its setting, and the message says where to put it.</summary>
    private static void RequireSecret(string? value, string name, string use, Setting setting)
    {
        if (string.IsNullOrEmpty(value))
            throw setting.Error(
                $"needs {name} for {use}. Supply it through the {setting.Variable(name)} environment variable or a secret store.");
    }

    /// <summary>The value is never quoted: a URL can carry credentials in its user-info part.</summary>
    private static void RequireUrl(string? value, bool httpsOnly, string name, string meaning, Setting setting)
    {
        if (!Uri.TryCreate(value?.Trim(), UriKind.Absolute, out var uri) ||
            !(uri.Scheme == Uri.UriSchemeHttps || (!httpsOnly && uri.Scheme == Uri.UriSchemeHttp)))
            throw setting.Error($"needs {name}, {meaning}. Set {setting.Variable(name)}.");
    }

    private static IManageGrantType ParseIManageGrant(string value, Setting setting)
    {
        if (string.IsNullOrWhiteSpace(value)) return IManageGrantType.Password;
        var trimmed = value.Trim();
        if (!char.IsDigit(trimmed[0]) && trimmed[0] != '-' &&
            Enum.TryParse<IManageGrantType>(trimmed, ignoreCase: true, out var grant) &&
            Enum.IsDefined(typeof(IManageGrantType), grant))
            return grant;
        throw setting.Error($"has GrantType '{trimmed}'. Expected Password, RefreshToken or ClientCredentials.");
    }

    private static BoxSubjectType ParseBoxSubject(string value, Setting setting)
    {
        if (string.IsNullOrWhiteSpace(value)) return BoxSubjectType.Enterprise;
        var trimmed = value.Trim();
        if (!char.IsDigit(trimmed[0]) && trimmed[0] != '-' &&
            Enum.TryParse<BoxSubjectType>(trimmed, ignoreCase: true, out var subject) &&
            Enum.IsDefined(typeof(BoxSubjectType), subject))
            return subject;
        throw setting.Error($"has SubjectType '{trimmed}'. Expected Enterprise or User.");
    }

    // ── Providers ─────────────────────────────────────────────────────────────

    /// <summary>
    /// Registers a provider for every iManage, Google Drive, Box, Azure Blob and Amazon S3
    /// connection. Each is built when the toolset is composed, which happens at startup, so a key
    /// that cannot be read or a value only the provider checks still stops the server there.
    /// </summary>
    public static void Register(IServiceCollection services, OfficeAgentMcpOptions options)
    {
        foreach (var connection in options.IManageConnections)
            Add(services, StorageKind.IManage, connection, sp => CreateIManage(sp, connection));
        foreach (var connection in options.GoogleDriveConnections)
            Add(services, StorageKind.GoogleDrive, connection, sp => CreateGoogleDrive(sp, connection));
        foreach (var connection in options.BoxConnections)
            Add(services, StorageKind.Box, connection, sp => CreateBox(sp, connection));
        foreach (var connection in options.AzureBlobConnections)
            Add(services, StorageKind.AzureBlob, connection, sp => CreateAzureBlob(sp, connection));
        foreach (var connection in options.AmazonS3Connections)
            Add(services, StorageKind.AmazonS3, connection, sp => CreateAmazonS3(sp, connection));
    }

    private static void Add(
        IServiceCollection services, StorageKind kind, StorageConnectionOptions connection,
        Func<IServiceProvider, IDocumentProvider> create) =>
        services.AddSingleton<IDocumentProvider>(sp =>
        {
            try
            {
                return create(sp);
            }
            catch (ArgumentException error)
            {
                // The providers' own checks name the problem but not the connection. Their
                // messages never quote a key, secret or connection string.
                var message = error.ParamName is null
                    ? error.Message
                    : error.Message.Replace($" (Parameter '{error.ParamName}')", string.Empty);
                throw new InvalidOperationException(
                    $"{kind.Label} connection '{connection.ConnectionId}' is not configured correctly: {message}", error);
            }
        });

    private static HttpClient Http(IServiceProvider services, string provider, Func<HttpClient> fallback) =>
        services.GetService<StorageTransports>()?.Http?.Invoke(provider) ?? fallback();

    private static IDocumentProvider CreateIManage(IServiceProvider services, IManageConnectionOptions c)
    {
        // The provider's own client, which does not follow redirects: iManage's X-Auth-Token
        // header, unlike Authorization, would otherwise travel to wherever a redirect points.
        var http = Http(services, StorageKind.IManage.Provider, IManageDocumentProvider.CreateHttpClient);
        var tokens = new IManageOAuthTokenProvider(new IManageOAuthOptions
        {
            ServerUrl = c.ServerUrl.Trim(),
            TokenPath = c.TokenPath,
            ClientId = c.ClientId,
            ClientSecret = c.ClientSecret,
            GrantType = ParseIManageGrant(c.GrantType, new Setting(StorageKind.IManage, string.Empty, 0, c.ConnectionId)),
            Username = c.Username,
            Password = c.Password,
            RefreshToken = c.RefreshToken
        }, http);
        IIManageRegistrationStore? store = string.IsNullOrWhiteSpace(c.RegistrationIndexPath)
            ? null
            : new JsonFileIManageRegistrationStore(c.RegistrationIndexPath);

        return new IManageDocumentProvider(new IManageDocumentProviderOptions
        {
            ConnectionId = c.ConnectionId,
            ServerUrl = c.ServerUrl.Trim(),
            ApiBasePath = c.ApiBasePath,
            CustomerId = c.CustomerId.Trim(),
            MaximumBytes = c.MaximumBytes,
            AllowedExtensions = c.AllowedExtensions.ToArray(),
            CreationFolderId = c.CreationFolderId,
            CreationDocumentClass = c.CreationDocumentClass,
            SaveComment = c.SaveComment,
            LockDuringSave = c.LockDuringSave,
            DefaultChangeMode = OfficeAgentMcpServer.ParseChangeMode(c.DefaultChangeMode, c.ConnectionId),
            DefaultChangeGranularity = OfficeAgentMcpServer.ParseChangeGranularity(c.DefaultChangeGranularity, c.ConnectionId),
            CalculateFormulas = c.CalculateFormulas,
            RejectFormulaErrors = c.RejectFormulaErrors
        }, http, tokens, store);
    }

    private static IDocumentProvider CreateGoogleDrive(IServiceProvider services, GoogleDriveConnectionOptions c)
    {
        var http = Http(services, StorageKind.GoogleDrive.Provider, () => services.GetService<HttpClient>() ?? new HttpClient());
        var keyJson = string.IsNullOrWhiteSpace(c.KeyFilePath) ? c.KeyJson : File.ReadAllText(c.KeyFilePath.Trim());
        var tokens = new GoogleServiceAccountAccessTokenProvider(new GoogleServiceAccountOptions
        {
            KeyJson = keyJson,
            Subject = c.Subject,
            Scope = c.Scope
        }, http);
        IGoogleDriveRegistrationStore? store = string.IsNullOrWhiteSpace(c.RegistrationIndexPath)
            ? null
            : new JsonFileGoogleDriveRegistrationStore(c.RegistrationIndexPath);

        return new GoogleDriveDocumentProvider(new GoogleDriveDocumentProviderOptions
        {
            ConnectionId = c.ConnectionId,
            ApiBaseUrl = c.ApiBaseUrl.Trim(),
            UploadBaseUrl = c.UploadBaseUrl.Trim(),
            MaximumBytes = c.MaximumBytes,
            AllowedExtensions = c.AllowedExtensions.ToArray(),
            CreationFolderId = c.CreationFolderId,
            DefaultChangeMode = OfficeAgentMcpServer.ParseChangeMode(c.DefaultChangeMode, c.ConnectionId),
            DefaultChangeGranularity = OfficeAgentMcpServer.ParseChangeGranularity(c.DefaultChangeGranularity, c.ConnectionId),
            CalculateFormulas = c.CalculateFormulas,
            RejectFormulaErrors = c.RejectFormulaErrors
        }, http, tokens, store);
    }

    private static IDocumentProvider CreateBox(IServiceProvider services, BoxConnectionOptions c)
    {
        var http = Http(services, StorageKind.Box.Provider, () => services.GetService<HttpClient>() ?? new HttpClient());
        var tokens = new BoxClientCredentialsAccessTokenProvider(new BoxClientCredentialsOptions
        {
            ClientId = c.ClientId,
            ClientSecret = c.ClientSecret,
            SubjectType = ParseBoxSubject(c.SubjectType, new Setting(StorageKind.Box, string.Empty, 0, c.ConnectionId)),
            SubjectId = c.SubjectId.Trim()
        }, http);
        IBoxRegistrationStore? store = string.IsNullOrWhiteSpace(c.RegistrationIndexPath)
            ? null
            : new JsonFileBoxRegistrationStore(c.RegistrationIndexPath);

        return new BoxDocumentProvider(new BoxDocumentProviderOptions
        {
            ConnectionId = c.ConnectionId,
            ApiBaseUrl = c.ApiBaseUrl.Trim(),
            UploadBaseUrl = c.UploadBaseUrl.Trim(),
            MaximumBytes = c.MaximumBytes,
            AllowedExtensions = c.AllowedExtensions.ToArray(),
            CreationFolderId = c.CreationFolderId,
            DefaultChangeMode = OfficeAgentMcpServer.ParseChangeMode(c.DefaultChangeMode, c.ConnectionId),
            DefaultChangeGranularity = OfficeAgentMcpServer.ParseChangeGranularity(c.DefaultChangeGranularity, c.ConnectionId),
            CalculateFormulas = c.CalculateFormulas,
            RejectFormulaErrors = c.RejectFormulaErrors
        }, http, tokens, store);
    }

    private static IDocumentProvider CreateAzureBlob(IServiceProvider services, AzureBlobConnectionOptions c)
    {
        var options = new AzureBlobDocumentProviderOptions
        {
            ConnectionId = c.ConnectionId,
            ServiceUri = c.ServiceUri,
            ConnectionString = c.ConnectionString,
            ContainerName = c.ContainerName.Trim(),
            Prefix = c.Prefix,
            ManagedIdentityClientId = c.ManagedIdentityClientId,
            RegistrationIndexPath = c.RegistrationIndexPath,
            MaximumBytes = c.MaximumBytes,
            AllowedExtensions = c.AllowedExtensions.ToArray(),
            DefaultChangeMode = OfficeAgentMcpServer.ParseChangeMode(c.DefaultChangeMode, c.ConnectionId),
            DefaultChangeGranularity = OfficeAgentMcpServer.ParseChangeGranularity(c.DefaultChangeGranularity, c.ConnectionId),
            CalculateFormulas = c.CalculateFormulas,
            RejectFormulaErrors = c.RejectFormulaErrors
        };
        return services.GetService<StorageTransports>()?.AzureBlob is { } container
            ? new AzureBlobDocumentProvider(options, container(c))
            : new AzureBlobDocumentProvider(options);
    }

    private static IDocumentProvider CreateAmazonS3(IServiceProvider services, AmazonS3ConnectionOptions c)
    {
        var options = new AmazonS3DocumentProviderOptions
        {
            ConnectionId = c.ConnectionId,
            BucketName = c.BucketName.Trim(),
            KeyPrefix = c.KeyPrefix,
            Region = c.Region,
            ServiceUrl = c.ServiceUrl,
            ForcePathStyle = c.ForcePathStyle,
            RegistrationIndexPath = c.RegistrationIndexPath,
            MaximumBytes = c.MaximumBytes,
            AllowedExtensions = c.AllowedExtensions.ToArray(),
            DefaultChangeMode = OfficeAgentMcpServer.ParseChangeMode(c.DefaultChangeMode, c.ConnectionId),
            DefaultChangeGranularity = OfficeAgentMcpServer.ParseChangeGranularity(c.DefaultChangeGranularity, c.ConnectionId),
            CalculateFormulas = c.CalculateFormulas,
            RejectFormulaErrors = c.RejectFormulaErrors
        };
        return services.GetService<StorageTransports>()?.AmazonS3 is { } client
            ? new AmazonS3DocumentProvider(options, client(c))
            : new AmazonS3DocumentProvider(options);
    }
}
