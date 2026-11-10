using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace OfficeAgent.GoogleDrive;

/// <summary>
/// Supplies bearer tokens for Google Drive API calls. The provider asks for a token per
/// request; implementations own caching and refresh. Tokens never appear in document
/// references, tool results, or provider error messages.
/// </summary>
public interface IGoogleDriveAccessTokenProvider
{
    /// <summary>Returns a currently valid Google OAuth 2.0 access token for the Drive API.</summary>
    Task<string> GetAccessTokenAsync(CancellationToken cancellationToken = default);
}

/// <summary>The Google OAuth scopes the Drive provider can work with.</summary>
public static class GoogleDriveScopes
{
    /// <summary>
    /// Full Drive access, bounded by what the identity can see. Needed to edit files the
    /// identity did not create. Without domain-wide delegation a service account sees only
    /// the files, folders, and shared drives shared with it.
    /// </summary>
    public const string Drive = "https://www.googleapis.com/auth/drive";

    /// <summary>
    /// Per-file access to files the app created or a user opened with it. Enough for a
    /// connection that only edits documents it created itself; not enough to register an
    /// existing file.
    /// </summary>
    public const string DriveFile = "https://www.googleapis.com/auth/drive.file";
}

/// <summary>
/// Configures the Google service-account flow (OAuth 2.0 JWT bearer grant). Set
/// <see cref="KeyJson"/> to the contents of the key file downloaded from Google Cloud, or set
/// <see cref="ClientEmail"/> and <see cref="PrivateKey"/> from a secret store.
/// </summary>
public sealed class GoogleServiceAccountOptions
{
    /// <summary>
    /// Gets or sets the service account key file's JSON contents. When set, it supplies the
    /// client email, private key, and key id. Its <c>token_uri</c> is ignored:
    /// the assertion is only ever sent to <see cref="TokenUri"/>.
    /// </summary>
    public string KeyJson { get; set; } = string.Empty;

    /// <summary>Gets or sets the service account's email address (the JWT issuer).</summary>
    public string ClientEmail { get; set; } = string.Empty;

    /// <summary>Gets or sets the service account's PEM-encoded RSA private key.</summary>
    public string PrivateKey { get; set; } = string.Empty;

    /// <summary>Gets or sets the private key id, sent as the JWT <c>kid</c>. Optional.</summary>
    public string PrivateKeyId { get; set; } = string.Empty;

    /// <summary>Gets or sets the requested scope. The default is <see cref="GoogleDriveScopes.Drive"/>.</summary>
    public string Scope { get; set; } = GoogleDriveScopes.Drive;

    /// <summary>
    /// Gets or sets the Google Workspace user to act as, through domain-wide delegation. Leave
    /// empty for the service account to act as itself, which is the narrower setup.
    /// </summary>
    public string Subject { get; set; } = string.Empty;

    /// <summary>Gets or sets the token endpoint, which is also the JWT audience.</summary>
    public string TokenUri { get; set; } = "https://oauth2.googleapis.com/token";
}

/// <summary>
/// Acquires Drive access tokens for a Google service account with the OAuth 2.0 JWT bearer
/// grant: it signs a short-lived assertion with the account's RSA key (RS256) and exchanges
/// it at Google's token endpoint. No Google client library is involved. Tokens are cached
/// until shortly before they expire.
/// </summary>
/// <remarks>
/// Without <see cref="GoogleServiceAccountOptions.Subject"/> the service account acts as
/// itself and reaches only what has been shared with it. With a subject, it impersonates that
/// user through domain-wide delegation, which a Workspace super administrator must authorize
/// for the account's client id and the exact scope.
/// </remarks>
public sealed class GoogleServiceAccountAccessTokenProvider : IGoogleDriveAccessTokenProvider
{
    private const int AssertionLifetimeSeconds = 3600;

    private readonly HttpClient _http;
    private readonly string _clientEmail;
    private readonly string _privateKey;
    private readonly string _privateKeyId;
    private readonly string _scope;
    private readonly string _subject;
    private readonly string _tokenUri;
    private readonly SemaphoreSlim _gate = new(1, 1);
    private string? _token;
    private DateTimeOffset _expiresAt = DateTimeOffset.MinValue;

    /// <summary>Initializes the provider over one service account key.</summary>
    public GoogleServiceAccountAccessTokenProvider(GoogleServiceAccountOptions options, HttpClient httpClient)
    {
        if (options is null) throw new ArgumentNullException(nameof(options));
        _http = httpClient ?? throw new ArgumentNullException(nameof(httpClient));

        var clientEmail = options.ClientEmail;
        var privateKey = options.PrivateKey;
        var privateKeyId = options.PrivateKeyId;
        if (!string.IsNullOrWhiteSpace(options.KeyJson))
            ReadKeyJson(options.KeyJson, ref clientEmail, ref privateKey, ref privateKeyId);

        if (string.IsNullOrWhiteSpace(clientEmail) || string.IsNullOrWhiteSpace(privateKey))
            throw new ArgumentException(
                "A service account client email and private key are required: set KeyJson, or ClientEmail and PrivateKey.",
                nameof(options));
        if (string.IsNullOrWhiteSpace(options.Scope))
            throw new ArgumentException("A scope is required.", nameof(options));
        if (string.IsNullOrWhiteSpace(options.TokenUri))
            throw new ArgumentException("A token endpoint is required.", nameof(options));

        // Fail at startup, not on the first request, when the key cannot be read. The message
        // never includes the key.
        try
        {
            using var rsa = RsaPrivateKey.Create(privateKey);
        }
        catch (Exception ex) when (ex is CryptographicException or ArgumentException or FormatException)
        {
            throw new ArgumentException("The service account private key is not a readable PEM RSA key.", nameof(options));
        }

        _clientEmail = clientEmail.Trim();
        _privateKey = privateKey;
        _privateKeyId = privateKeyId?.Trim() ?? string.Empty;
        _scope = options.Scope.Trim();
        _subject = options.Subject?.Trim() ?? string.Empty;
        _tokenUri = options.TokenUri.Trim();
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

            using var request = new HttpRequestMessage(HttpMethod.Post, _tokenUri)
            {
                Content = new FormUrlEncodedContent(new Dictionary<string, string>
                {
                    ["grant_type"] = "urn:ietf:params:oauth:grant-type:jwt-bearer",
                    ["assertion"] = CreateAssertion(DateTimeOffset.UtcNow)
                })
            };

            using var response = await _http.SendAsync(request, cancellationToken).ConfigureAwait(false);
            var payload = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            if (!response.IsSuccessStatusCode)
            {
                var error = ReadErrorCode(payload);
                var suffix = string.IsNullOrEmpty(error) ? string.Empty : $", error '{error}'";
                throw new InvalidOperationException(
                    $"Google token acquisition failed (status {(int)response.StatusCode}{suffix}) for service account '{_clientEmail}'.");
            }

            using var json = JsonDocument.Parse(payload);
            var token = json.RootElement.TryGetProperty("access_token", out var accessToken)
                ? accessToken.GetString()
                : null;
            if (string.IsNullOrEmpty(token))
                throw new InvalidOperationException("The Google token response did not contain an access_token.");
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

    /// <summary>Builds and signs the JWT assertion (RS256) for one token request.</summary>
    internal string CreateAssertion(DateTimeOffset now)
    {
        var header = new Dictionary<string, string> { ["alg"] = "RS256", ["typ"] = "JWT" };
        if (_privateKeyId.Length > 0) header["kid"] = _privateKeyId;

        var issuedAt = now.ToUnixTimeSeconds();
        var claims = new Dictionary<string, object>
        {
            ["iss"] = _clientEmail,
            ["scope"] = _scope,
            ["aud"] = _tokenUri,
            ["iat"] = issuedAt,
            ["exp"] = issuedAt + AssertionLifetimeSeconds
        };
        if (_subject.Length > 0) claims["sub"] = _subject;

        var signingInput =
            Base64Url(JsonSerializer.SerializeToUtf8Bytes(header)) + "." +
            Base64Url(JsonSerializer.SerializeToUtf8Bytes(claims));
        using var rsa = RsaPrivateKey.Create(_privateKey);
        var signature = rsa.SignData(
            Encoding.ASCII.GetBytes(signingInput), HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
        return signingInput + "." + Base64Url(signature);
    }

    private static void ReadKeyJson(string keyJson, ref string clientEmail, ref string privateKey, ref string privateKeyId)
    {
        try
        {
            using var json = JsonDocument.Parse(keyJson);
            var root = json.RootElement;
            if (root.TryGetProperty("type", out var type) &&
                !string.Equals(type.GetString(), "service_account", StringComparison.Ordinal))
                throw new ArgumentException(
                    "The key JSON is not a service account key (its type is not 'service_account').", nameof(keyJson));
            if (root.TryGetProperty("client_email", out var email)) clientEmail = email.GetString() ?? clientEmail;
            if (root.TryGetProperty("private_key", out var key)) privateKey = key.GetString() ?? privateKey;
            if (root.TryGetProperty("private_key_id", out var keyId)) privateKeyId = keyId.GetString() ?? privateKeyId;
        }
        catch (JsonException)
        {
            // JsonException messages can quote the payload, which holds the private key.
            throw new ArgumentException("The service account key JSON could not be parsed.", nameof(keyJson));
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

    internal static string Base64Url(byte[] bytes) =>
        Convert.ToBase64String(bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_');
}

/// <summary>
/// Reads a PEM RSA private key: PKCS#8 (<c>BEGIN PRIVATE KEY</c>, the form Google key files
/// use) or PKCS#1 (<c>BEGIN RSA PRIVATE KEY</c>). On .NET 10 the platform importer does the
/// work; netstandard2.0 has no PEM or PKCS#8 importer, so a minimal DER reader is used there.
/// </summary>
internal static class RsaPrivateKey
{
    private const string RsaEncryptionOid = "1.2.840.113549.1.1.1";

    public static RSA Create(string pem)
    {
        var rsa = RSA.Create();
        try
        {
#if NET
            rsa.ImportFromPem(pem);
#else
            rsa.ImportParameters(ReadParameters(pem));
#endif
            return rsa;
        }
        catch
        {
            rsa.Dispose();
            throw;
        }
    }

    /// <summary>Parses a PEM RSA private key into <see cref="RSAParameters"/>.</summary>
    public static RSAParameters ReadParameters(string pem)
    {
        var der = DecodePem(pem, out var label);
        return label switch
        {
            "PRIVATE KEY" => ReadPkcs8(der),
            "RSA PRIVATE KEY" => ReadPkcs1(der),
            _ => throw new CryptographicException("The PEM block is not an RSA private key.")
        };
    }

    private static byte[] DecodePem(string pem, out string label)
    {
        if (string.IsNullOrWhiteSpace(pem)) throw new CryptographicException("The private key is empty.");
        const string begin = "-----BEGIN ";
        var start = pem.IndexOf(begin, StringComparison.Ordinal);
        if (start < 0) throw new CryptographicException("The private key is not PEM-encoded.");
        var labelStart = start + begin.Length;
        var labelEnd = pem.IndexOf("-----", labelStart, StringComparison.Ordinal);
        if (labelEnd < 0) throw new CryptographicException("The private key is not PEM-encoded.");
        label = pem.Substring(labelStart, labelEnd - labelStart);
        var bodyStart = labelEnd + 5;
        var end = pem.IndexOf("-----END " + label + "-----", bodyStart, StringComparison.Ordinal);
        if (end < 0) throw new CryptographicException("The private key PEM block is not terminated.");
        var body = new StringBuilder(end - bodyStart);
        foreach (var c in pem.Substring(bodyStart, end - bodyStart))
            if (!char.IsWhiteSpace(c)) body.Append(c);
        return Convert.FromBase64String(body.ToString());
    }

    private static RSAParameters ReadPkcs8(byte[] der)
    {
        // PrivateKeyInfo ::= SEQUENCE { version INTEGER, privateKeyAlgorithm AlgorithmIdentifier,
        //                               privateKey OCTET STRING, ... }
        var info = new DerReader(der).ReadSequence();
        info.ReadInteger();
        var algorithm = info.ReadSequence();
        if (algorithm.ReadObjectIdentifier() != RsaEncryptionOid)
            throw new CryptographicException("The private key is not an RSA key.");
        return ReadPkcs1(info.ReadOctetString());
    }

    private static RSAParameters ReadPkcs1(byte[] der)
    {
        // RSAPrivateKey ::= SEQUENCE { version, modulus, publicExponent, privateExponent,
        //                              prime1, prime2, exponent1, exponent2, coefficient }
        var key = new DerReader(der).ReadSequence();
        key.ReadInteger();
        var modulus = key.ReadInteger();
        var exponent = key.ReadInteger();
        var d = key.ReadInteger();
        var p = key.ReadInteger();
        var q = key.ReadInteger();
        var dp = key.ReadInteger();
        var dq = key.ReadInteger();
        var inverseQ = key.ReadInteger();

        // The platform expects D sized like the modulus and the CRT values sized like half of it.
        var half = (modulus.Length + 1) / 2;
        return new RSAParameters
        {
            Modulus = modulus,
            Exponent = exponent,
            D = PadLeft(d, modulus.Length),
            P = PadLeft(p, half),
            Q = PadLeft(q, half),
            DP = PadLeft(dp, half),
            DQ = PadLeft(dq, half),
            InverseQ = PadLeft(inverseQ, half)
        };
    }

    private static byte[] PadLeft(byte[] value, int length)
    {
        if (value.Length == length) return value;
        if (value.Length > length) throw new CryptographicException("An RSA key component is longer than expected.");
        var padded = new byte[length];
        Buffer.BlockCopy(value, 0, padded, length - value.Length, value.Length);
        return padded;
    }

    /// <summary>A forward-only reader for the few DER types a private key uses.</summary>
    private sealed class DerReader
    {
        private readonly byte[] _data;
        private int _position;
        private readonly int _end;

        public DerReader(byte[] data) : this(data, 0, data.Length) { }

        private DerReader(byte[] data, int offset, int length)
        {
            _data = data;
            _position = offset;
            _end = offset + length;
        }

        public DerReader ReadSequence()
        {
            var (offset, length) = ReadElement(0x30);
            return new DerReader(_data, offset, length);
        }

        public byte[] ReadOctetString() => Copy(ReadElement(0x04));

        /// <summary>Reads an unsigned big-endian INTEGER without its sign-padding zeros.</summary>
        public byte[] ReadInteger()
        {
            var (offset, length) = ReadElement(0x02);
            while (length > 1 && _data[offset] == 0)
            {
                offset++;
                length--;
            }
            return Copy((offset, length));
        }

        public string ReadObjectIdentifier()
        {
            var (offset, length) = ReadElement(0x06);
            if (length == 0) throw new CryptographicException("Empty object identifier.");
            var builder = new StringBuilder();
            builder.Append(_data[offset] / 40).Append('.').Append(_data[offset] % 40);
            long value = 0;
            for (var i = 1; i < length; i++)
            {
                value = (value << 7) | (uint)(_data[offset + i] & 0x7F);
                if ((_data[offset + i] & 0x80) == 0)
                {
                    builder.Append('.').Append(value);
                    value = 0;
                }
            }
            return builder.ToString();
        }

        private (int Offset, int Length) ReadElement(byte expectedTag)
        {
            if (_position >= _end || _data[_position] != expectedTag)
                throw new CryptographicException("The private key's DER structure is not as expected.");
            _position++;
            var length = ReadLength();
            if (length < 0 || _position + length > _end)
                throw new CryptographicException("The private key's DER structure is truncated.");
            var offset = _position;
            _position += length;
            return (offset, length);
        }

        private int ReadLength()
        {
            if (_position >= _end) throw new CryptographicException("The private key's DER structure is truncated.");
            int first = _data[_position++];
            if (first < 0x80) return first;
            var count = first & 0x7F;
            if (count is 0 or > 4) throw new CryptographicException("Unsupported DER length.");
            var length = 0;
            for (var i = 0; i < count; i++)
            {
                if (_position >= _end) throw new CryptographicException("The private key's DER structure is truncated.");
                length = (length << 8) | _data[_position++];
            }
            return length;
        }

        private byte[] Copy((int Offset, int Length) element)
        {
            var copy = new byte[element.Length];
            Buffer.BlockCopy(_data, element.Offset, copy, 0, element.Length);
            return copy;
        }
    }
}
