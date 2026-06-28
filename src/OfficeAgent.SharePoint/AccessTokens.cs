using System.Collections.Concurrent;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace OfficeAgent.SharePoint;

/// <summary>
/// Supplies bearer tokens for Microsoft Graph calls. The provider asks for a token
/// per request; implementations own caching and refresh. Tokens never appear in
/// document references or tool results.
/// </summary>
public interface IAccessTokenProvider
{
    /// <summary>Returns a currently valid Graph access token.</summary>
    Task<string> GetAccessTokenAsync(CancellationToken cancellationToken = default);
}

/// <summary>Configures the app-only (OAuth2 client-credentials) flow against Microsoft Entra ID.</summary>
public sealed class AppOnlyOptions
{
    /// <summary>Gets or sets the Entra tenant id (GUID or domain).</summary>
    public string TenantId { get; set; } = string.Empty;

    /// <summary>Gets or sets the app registration's client id.</summary>
    public string ClientId { get; set; } = string.Empty;

    /// <summary>Gets or sets the app registration's client secret.</summary>
    public string ClientSecret { get; set; } = string.Empty;

    /// <summary>Gets or sets the requested scope. The default is the Graph application scope.</summary>
    public string Scope { get; set; } = "https://graph.microsoft.com/.default";

    /// <summary>Gets or sets the token authority base URL.</summary>
    public string Authority { get; set; } = "https://login.microsoftonline.com";
}

/// <summary>
/// Acquires app-only Graph tokens via the OAuth2 client-credentials flow and caches
/// them until shortly before expiry. The service authenticates as itself (one shared
/// app identity, no user). Suitable for daemon-style hosts (the OfficeAgent MCP
/// server, background services) where no user is present.
/// </summary>
public sealed class AppOnlyAccessTokenProvider : IAccessTokenProvider
{
    private readonly AppOnlyOptions _options;
    private readonly HttpClient _http;
    private readonly SemaphoreSlim _gate = new(1, 1);
    private string? _token;
    private DateTimeOffset _expiresAt = DateTimeOffset.MinValue;

    /// <summary>Initializes the provider over an Entra app registration.</summary>
    public AppOnlyAccessTokenProvider(AppOnlyOptions options, HttpClient httpClient)
    {
        _options = options ?? throw new ArgumentNullException(nameof(options));
        _http = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
        if (string.IsNullOrWhiteSpace(options.TenantId) ||
            string.IsNullOrWhiteSpace(options.ClientId) ||
            string.IsNullOrWhiteSpace(options.ClientSecret))
            throw new ArgumentException("TenantId, ClientId, and ClientSecret are required.", nameof(options));
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

            var endpoint = $"{_options.Authority.TrimEnd('/')}/{_options.TenantId}/oauth2/v2.0/token";
            using var request = new HttpRequestMessage(HttpMethod.Post, endpoint)
            {
                Content = new FormUrlEncodedContent(new Dictionary<string, string>
                {
                    ["grant_type"] = "client_credentials",
                    ["client_id"] = _options.ClientId,
                    ["client_secret"] = _options.ClientSecret,
                    ["scope"] = _options.Scope
                })
            };

            using var response = await _http.SendAsync(request, cancellationToken).ConfigureAwait(false);
            var payload = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            if (!response.IsSuccessStatusCode)
                throw new InvalidOperationException(
                    $"Token acquisition failed with status {(int)response.StatusCode} for tenant '{_options.TenantId}'.");

            using var json = JsonDocument.Parse(payload);
            var token = json.RootElement.GetProperty("access_token").GetString()
                ?? throw new InvalidOperationException("The token response did not contain an access_token.");
            var expiresIn = json.RootElement.TryGetProperty("expires_in", out var e) ? e.GetInt32() : 300;

            _token = token;
            _expiresAt = DateTimeOffset.UtcNow.AddSeconds(Math.Max(60, expiresIn - 60));
            return token;
        }
        finally
        {
            _gate.Release();
        }
    }
}

/// <summary>Configures the OAuth2 On-Behalf-Of flow against Microsoft Entra ID.</summary>
public sealed class OnBehalfOfOptions
{
    /// <summary>Gets or sets the Entra tenant id (GUID or domain).</summary>
    public string TenantId { get; set; } = string.Empty;

    /// <summary>Gets or sets the middle-tier API's app registration client id.</summary>
    public string ClientId { get; set; } = string.Empty;

    /// <summary>Gets or sets the middle-tier API's client secret.</summary>
    public string ClientSecret { get; set; } = string.Empty;

    /// <summary>Gets or sets the requested downstream scope. The default is the Graph application scope.</summary>
    public string Scope { get; set; } = "https://graph.microsoft.com/.default";

    /// <summary>Gets or sets the token authority base URL.</summary>
    public string Authority { get; set; } = "https://login.microsoftonline.com";
}

/// <summary>
/// Acquires Graph tokens with the OAuth2 On-Behalf-Of flow: it exchanges the end
/// user's access token (captured per request in <see cref="GraphUserContext"/>) for
/// a Graph token that carries the user's identity and delegated permissions, so the
/// SharePoint provider acts as the signed-in user rather than under a shared app
/// identity. Exchanged tokens are cached per user until shortly before expiry.
/// </summary>
/// <remarks>
/// This requires an inbound user token whose audience is the middle-tier API
/// (<see cref="OnBehalfOfOptions.ClientId"/>), the user's prior consent to that API,
/// and the API having the matching delegated Graph permission. It is therefore a
/// hosted-HTTP concern: a stdio host has no inbound user token, and a call with no
/// captured token fails with a clear error.
/// </remarks>
public sealed class OnBehalfOfAccessTokenProvider : IAccessTokenProvider
{
    private readonly OnBehalfOfOptions _options;
    private readonly HttpClient _http;
    private readonly Func<string?> _userTokenAccessor;
    private readonly ConcurrentDictionary<string, CachedToken> _cache = new(StringComparer.Ordinal);

    /// <summary>Initializes the provider over a middle-tier app registration.</summary>
    public OnBehalfOfAccessTokenProvider(OnBehalfOfOptions options, HttpClient httpClient)
        : this(options, httpClient, () => GraphUserContext.CurrentUserAccessToken)
    {
    }

    /// <summary>Initializes the provider with an explicit accessor for the inbound user token (testing/custom hosts).</summary>
    public OnBehalfOfAccessTokenProvider(OnBehalfOfOptions options, HttpClient httpClient, Func<string?> userTokenAccessor)
    {
        _options = options ?? throw new ArgumentNullException(nameof(options));
        _http = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
        _userTokenAccessor = userTokenAccessor ?? throw new ArgumentNullException(nameof(userTokenAccessor));
        if (string.IsNullOrWhiteSpace(options.TenantId) ||
            string.IsNullOrWhiteSpace(options.ClientId) ||
            string.IsNullOrWhiteSpace(options.ClientSecret))
            throw new ArgumentException("TenantId, ClientId, and ClientSecret are required.", nameof(options));
    }

    /// <inheritdoc />
    public async Task<string> GetAccessTokenAsync(CancellationToken cancellationToken = default)
    {
        var userAssertion = _userTokenAccessor();
        if (string.IsNullOrWhiteSpace(userAssertion))
            throw new InvalidOperationException(
                "The On-Behalf-Of flow requires the caller's access token, but none was captured for this request. " +
                "Host the MCP server over HTTP behind authentication so the inbound bearer token is available, or " +
                "configure client-credentials authentication for unattended scenarios.");

        var key = CacheKey(userAssertion!);
        if (_cache.TryGetValue(key, out var cached) && DateTimeOffset.UtcNow < cached.ExpiresAt)
            return cached.Token;

        var (token, expiresIn) = await ExchangeAsync(userAssertion!, cancellationToken).ConfigureAwait(false);
        var entry = new CachedToken(token, DateTimeOffset.UtcNow.AddSeconds(Math.Max(60, expiresIn - 60)));
        _cache[key] = entry;
        PruneExpired();
        return token;
    }

    private async Task<(string Token, int ExpiresIn)> ExchangeAsync(string userAssertion, CancellationToken cancellationToken)
    {
        var endpoint = $"{_options.Authority.TrimEnd('/')}/{_options.TenantId}/oauth2/v2.0/token";
        using var request = new HttpRequestMessage(HttpMethod.Post, endpoint)
        {
            Content = new FormUrlEncodedContent(new Dictionary<string, string>
            {
                ["grant_type"] = "urn:ietf:params:oauth:grant-type:jwt-bearer",
                ["client_id"] = _options.ClientId,
                ["client_secret"] = _options.ClientSecret,
                ["assertion"] = userAssertion,
                ["scope"] = _options.Scope,
                ["requested_token_use"] = "on_behalf_of"
            })
        };

        using var response = await _http.SendAsync(request, cancellationToken).ConfigureAwait(false);
        var payload = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
        if (!response.IsSuccessStatusCode)
            throw new InvalidOperationException(
                $"On-Behalf-Of token exchange failed with status {(int)response.StatusCode} for tenant '{_options.TenantId}'. " +
                $"{DescribeError(payload)}".TrimEnd());

        using var json = JsonDocument.Parse(payload);
        var token = json.RootElement.GetProperty("access_token").GetString()
            ?? throw new InvalidOperationException("The On-Behalf-Of token response did not contain an access_token.");
        var expiresIn = json.RootElement.TryGetProperty("expires_in", out var e) ? e.GetInt32() : 300;
        return (token, expiresIn);
    }

    private static string DescribeError(string payload)
    {
        if (string.IsNullOrWhiteSpace(payload)) return string.Empty;
        try
        {
            using var json = JsonDocument.Parse(payload);
            return json.RootElement.TryGetProperty("error_description", out var description)
                ? description.GetString()?.Split('\n')[0] ?? string.Empty
                : json.RootElement.TryGetProperty("error", out var error) ? error.GetString() ?? string.Empty : string.Empty;
        }
        catch (JsonException)
        {
            return string.Empty;
        }
    }

    private void PruneExpired()
    {
        if (_cache.Count < 256) return;
        var now = DateTimeOffset.UtcNow;
        foreach (var pair in _cache)
            if (now >= pair.Value.ExpiresAt)
                _cache.TryRemove(pair.Key, out _);
    }

    private static string CacheKey(string userAssertion)
    {
        using var sha = SHA256.Create();
        var hash = sha.ComputeHash(Encoding.UTF8.GetBytes(userAssertion));
        return Convert.ToBase64String(hash);
    }

    private readonly struct CachedToken
    {
        public CachedToken(string token, DateTimeOffset expiresAt)
        {
            Token = token;
            ExpiresAt = expiresAt;
        }

        public string Token { get; }
        public DateTimeOffset ExpiresAt { get; }
    }
}
