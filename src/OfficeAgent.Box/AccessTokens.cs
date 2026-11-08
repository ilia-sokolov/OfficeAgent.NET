using System.Text.Json;

namespace OfficeAgent.Box;

/// <summary>
/// Supplies bearer tokens for Box API calls. The provider asks for a token per request;
/// implementations own caching and refresh. Tokens never appear in document references,
/// tool results, or provider error messages.
/// </summary>
public interface IBoxAccessTokenProvider
{
    /// <summary>Returns a currently valid Box access token.</summary>
    Task<string> GetAccessTokenAsync(CancellationToken cancellationToken = default);
}

/// <summary>Who a Box client-credentials token acts as.</summary>
public enum BoxSubjectType
{
    /// <summary>The application's service account, identified by the enterprise id.</summary>
    Enterprise = 0,

    /// <summary>A managed user or app user, identified by the user id.</summary>
    User = 1
}

/// <summary>Configures Box server authentication with the client credentials grant.</summary>
public sealed class BoxClientCredentialsOptions
{
    /// <summary>Gets or sets the Box app's client id.</summary>
    public string ClientId { get; set; } = string.Empty;

    /// <summary>Gets or sets the Box app's client secret.</summary>
    public string ClientSecret { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets who the token acts as. <see cref="BoxSubjectType.Enterprise"/> (the default)
    /// acts as the app's service account.
    /// </summary>
    public BoxSubjectType SubjectType { get; set; } = BoxSubjectType.Enterprise;

    /// <summary>Gets or sets the enterprise id, or the user id for <see cref="BoxSubjectType.User"/>.</summary>
    public string SubjectId { get; set; } = string.Empty;

    /// <summary>Gets or sets the token endpoint.</summary>
    public string TokenUrl { get; set; } = "https://api.box.com/oauth2/token";
}

/// <summary>
/// Acquires Box access tokens with the client credentials grant (Box "server authentication
/// with client credentials grant") and caches them until shortly before they expire. Box
/// issues no refresh token for this grant; a new token is requested when the old one ends.
/// </summary>
/// <remarks>
/// With an enterprise subject the token acts as the app's service account, which reaches
/// only content it owns or is collaborated on. A user subject needs the app's
/// "Generate user access tokens" setting, and a managed user also needs App + Enterprise Access.
/// </remarks>
public sealed class BoxClientCredentialsAccessTokenProvider : IBoxAccessTokenProvider
{
    private readonly BoxClientCredentialsOptions _options;
    private readonly HttpClient _http;
    private readonly SemaphoreSlim _gate = new(1, 1);
    private string? _token;
    private DateTimeOffset _expiresAt = DateTimeOffset.MinValue;

    /// <summary>Initializes the provider over one Box app.</summary>
    public BoxClientCredentialsAccessTokenProvider(BoxClientCredentialsOptions options, HttpClient httpClient)
    {
        _options = options ?? throw new ArgumentNullException(nameof(options));
        _http = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
        if (string.IsNullOrWhiteSpace(options.ClientId) ||
            string.IsNullOrWhiteSpace(options.ClientSecret) ||
            string.IsNullOrWhiteSpace(options.SubjectId))
            throw new ArgumentException("ClientId, ClientSecret, and SubjectId are required.", nameof(options));
        if (options.SubjectType is not (BoxSubjectType.Enterprise or BoxSubjectType.User))
            throw new ArgumentOutOfRangeException(nameof(options), "SubjectType must be Enterprise or User.");
        if (string.IsNullOrWhiteSpace(options.TokenUrl))
            throw new ArgumentException("A token endpoint is required.", nameof(options));
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

            using var request = new HttpRequestMessage(HttpMethod.Post, _options.TokenUrl)
            {
                Content = new FormUrlEncodedContent(new Dictionary<string, string>
                {
                    ["grant_type"] = "client_credentials",
                    ["client_id"] = _options.ClientId,
                    ["client_secret"] = _options.ClientSecret,
                    ["box_subject_type"] = _options.SubjectType == BoxSubjectType.User ? "user" : "enterprise",
                    ["box_subject_id"] = _options.SubjectId
                })
            };

            using var response = await _http.SendAsync(request, cancellationToken).ConfigureAwait(false);
            var payload = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            if (!response.IsSuccessStatusCode)
            {
                var error = ReadErrorCode(payload);
                var suffix = string.IsNullOrEmpty(error) ? string.Empty : $", error '{error}'";
                throw new InvalidOperationException(
                    $"Box token acquisition failed (status {(int)response.StatusCode}{suffix}).");
            }

            using var json = JsonDocument.Parse(payload);
            var token = json.RootElement.TryGetProperty("access_token", out var accessToken)
                ? accessToken.GetString()
                : null;
            if (string.IsNullOrEmpty(token))
                throw new InvalidOperationException("The Box token response did not contain an access_token.");
            var expiresIn = json.RootElement.TryGetProperty("expires_in", out var e) && e.TryGetInt32(out var seconds)
                ? seconds
                : 300;

            _token = token;
            _expiresAt = DateTimeOffset.UtcNow.AddSeconds(Math.Max(60, expiresIn - 60));
            return token!;
        }
        finally
        {
            _gate.Release();
        }
    }

    private static string ReadErrorCode(string payload)
    {
        if (string.IsNullOrWhiteSpace(payload)) return string.Empty;
        try
        {
            using var json = JsonDocument.Parse(payload);
            return json.RootElement.TryGetProperty("error", out var error) && error.ValueKind == JsonValueKind.String
                ? error.GetString() ?? string.Empty
                : string.Empty;
        }
        catch (JsonException)
        {
            return string.Empty;
        }
    }
}
