using System.Text.Json;

namespace OfficeAgent.IManage;

/// <summary>
/// Supplies access tokens for iManage Work REST API calls. The provider asks for a token per
/// request; implementations own caching and refresh. Tokens never appear in document
/// references, tool results, or exception messages.
/// </summary>
public interface IIManageAccessTokenProvider
{
    /// <summary>Returns a currently valid iManage Work access token.</summary>
    Task<string> GetAccessTokenAsync(CancellationToken cancellationToken = default);
}

/// <summary>The OAuth2 grant an <see cref="IManageOAuthTokenProvider"/> uses to obtain tokens.</summary>
public enum IManageGrantType
{
    /// <summary>
    /// Resource-owner password grant for a dedicated iManage service account. This is the grant
    /// iManage documents for unattended server applications.
    /// </summary>
    Password,

    /// <summary>
    /// A refresh token obtained once through the authorization-code flow, with
    /// <c>Allow Refresh Token</c> enabled on the application in Control Center. The provider
    /// keeps the rotated refresh token in memory.
    /// </summary>
    RefreshToken,

    /// <summary>
    /// Client-credentials grant. Use it only where your iManage environment enables it for the
    /// registered application; iManage's public documentation does not describe it.
    /// </summary>
    ClientCredentials
}

/// <summary>Configures the OAuth2 token endpoint of an iManage Work server.</summary>
public sealed class IManageOAuthOptions
{
    /// <summary>
    /// Gets or sets the iManage server URL, for example <c>https://cloudimanage.com</c> or your
    /// on-premises Work server.
    /// </summary>
    public string ServerUrl { get; set; } = string.Empty;

    /// <summary>Gets or sets the token endpoint path on <see cref="ServerUrl"/>.</summary>
    public string TokenPath { get; set; } = "/auth/oauth2/token";

    /// <summary>Gets or sets the application's client id from Control Center.</summary>
    public string ClientId { get; set; } = string.Empty;

    /// <summary>Gets or sets the application's client secret. Leave empty for a public client.</summary>
    public string ClientSecret { get; set; } = string.Empty;

    /// <summary>Gets or sets the grant used to obtain tokens.</summary>
    public IManageGrantType GrantType { get; set; } = IManageGrantType.Password;

    /// <summary>Gets or sets the service account's user id, for <see cref="IManageGrantType.Password"/>.</summary>
    public string Username { get; set; } = string.Empty;

    /// <summary>Gets or sets the service account's password, for <see cref="IManageGrantType.Password"/>.</summary>
    public string Password { get; set; } = string.Empty;

    /// <summary>Gets or sets the initial refresh token, for <see cref="IManageGrantType.RefreshToken"/>.</summary>
    public string RefreshToken { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the requested scope. The default, <c>user</c>, acts with the account's own
    /// document permissions; never request <c>admin</c> for document editing.
    /// </summary>
    public string Scope { get; set; } = "user";
}

/// <summary>
/// Acquires iManage Work access tokens from the server's OAuth2 token endpoint and caches them
/// until shortly before expiry. Suitable for daemon-style hosts where no user is present.
/// </summary>
/// <remarks>
/// Give it an <see cref="HttpClient"/> that does not follow redirects, such as
/// <see cref="IManageDocumentProvider.CreateHttpClient"/>, so a redirect cannot carry the form
/// with the secret to another host.
/// </remarks>
public sealed class IManageOAuthTokenProvider : IIManageAccessTokenProvider
{
    private readonly IManageOAuthOptions _options;
    private readonly HttpClient _http;
    private readonly SemaphoreSlim _gate = new(1, 1);
    private string? _token;
    private string _refreshToken;
    private DateTimeOffset _expiresAt = DateTimeOffset.MinValue;

    /// <summary>Initializes the provider over an application registered in iManage Control Center.</summary>
    public IManageOAuthTokenProvider(IManageOAuthOptions options, HttpClient httpClient)
    {
        _options = options ?? throw new ArgumentNullException(nameof(options));
        _http = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
        if (!Uri.TryCreate(options.ServerUrl, UriKind.Absolute, out var server) || server.Scheme != Uri.UriSchemeHttps)
            throw new ArgumentException("An absolute https iManage ServerUrl is required.", nameof(options));
        if (string.IsNullOrWhiteSpace(options.ClientId))
            throw new ArgumentException("A ClientId is required.", nameof(options));
        switch (options.GrantType)
        {
            case IManageGrantType.Password when
                string.IsNullOrWhiteSpace(options.Username) || string.IsNullOrEmpty(options.Password):
                throw new ArgumentException("The password grant requires Username and Password.", nameof(options));
            case IManageGrantType.RefreshToken when string.IsNullOrWhiteSpace(options.RefreshToken):
                throw new ArgumentException("The refresh-token grant requires RefreshToken.", nameof(options));
            case IManageGrantType.ClientCredentials when string.IsNullOrEmpty(options.ClientSecret):
                throw new ArgumentException("The client-credentials grant requires ClientSecret.", nameof(options));
        }
        _refreshToken = options.RefreshToken;
    }

    /// <inheritdoc />
    public async Task<string> GetAccessTokenAsync(CancellationToken cancellationToken = default)
    {
        if (_token is not null && DateTimeOffset.UtcNow < _expiresAt)
            return _token;

        await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            if (_token is not null && DateTimeOffset.UtcNow < _expiresAt)
                return _token;

            var endpoint = _options.ServerUrl.TrimEnd('/') + "/" + _options.TokenPath.TrimStart('/');
            using var request = new HttpRequestMessage(HttpMethod.Post, endpoint)
            {
                Content = new FormUrlEncodedContent(Form())
            };

            using var response = await _http.SendAsync(request, cancellationToken).ConfigureAwait(false);
            var payload = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            if (!response.IsSuccessStatusCode)
            {
                // Only the status and the OAuth error code: the response can echo the request.
                var error = ReadErrorCode(payload);
                var suffix = string.IsNullOrEmpty(error) ? string.Empty : $", error '{error}'";
                throw new InvalidOperationException(
                    $"iManage token acquisition failed (status {(int)response.StatusCode}{suffix}).");
            }

            string token;
            int expiresIn;
            try
            {
                using var json = JsonDocument.Parse(payload);
                var root = json.RootElement;
                token = root.GetProperty("access_token").GetString()
                    ?? throw new InvalidOperationException("The iManage token response did not contain an access_token.");
                expiresIn = root.TryGetProperty("expires_in", out var e) && e.ValueKind == JsonValueKind.Number
                    ? e.GetInt32()
                    : 300;
                if (root.TryGetProperty("refresh_token", out var refresh) &&
                    refresh.ValueKind == JsonValueKind.String &&
                    !string.IsNullOrEmpty(refresh.GetString()))
                    _refreshToken = refresh.GetString()!;
            }
            catch (Exception ex) when (ex is JsonException or KeyNotFoundException)
            {
                throw new InvalidOperationException("The iManage token response was not readable.", ex);
            }

            _token = token;
            _expiresAt = DateTimeOffset.UtcNow.AddSeconds(Math.Max(60, expiresIn - 60));
            return token;
        }
        finally
        {
            _gate.Release();
        }
    }

    private Dictionary<string, string> Form()
    {
        var form = new Dictionary<string, string> { ["client_id"] = _options.ClientId };
        if (!string.IsNullOrEmpty(_options.ClientSecret))
            form["client_secret"] = _options.ClientSecret;
        switch (_options.GrantType)
        {
            case IManageGrantType.Password:
                form["grant_type"] = "password";
                form["username"] = _options.Username;
                form["password"] = _options.Password;
                break;
            case IManageGrantType.RefreshToken:
                form["grant_type"] = "refresh_token";
                form["refresh_token"] = _refreshToken;
                break;
            default:
                form["grant_type"] = "client_credentials";
                break;
        }
        if (!string.IsNullOrWhiteSpace(_options.Scope))
            form["scope"] = _options.Scope;
        return form;
    }

    private static string ReadErrorCode(string payload)
    {
        if (string.IsNullOrWhiteSpace(payload)) return string.Empty;
        try
        {
            using var json = JsonDocument.Parse(payload);
            return json.RootElement.ValueKind == JsonValueKind.Object &&
                   json.RootElement.TryGetProperty("error", out var error) &&
                   error.ValueKind == JsonValueKind.String
                ? error.GetString() ?? string.Empty
                : string.Empty;
        }
        catch (JsonException)
        {
            return string.Empty;
        }
    }
}
