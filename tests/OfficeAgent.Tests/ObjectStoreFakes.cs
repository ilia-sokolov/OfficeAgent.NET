using System.Net;
using System.Net.Http.Headers;
using System.Security;
using Amazon.Runtime;
using Amazon.S3;
using Azure.Core.Pipeline;
using Azure.Storage.Blobs;

namespace OfficeAgent.Tests;

/// <summary>One stored object in a fake blob container or bucket.</summary>
internal sealed class FakeObject
{
    public byte[] Bytes { get; set; } = Array.Empty<byte>();
    public string ETag { get; set; } = string.Empty;
    public string? VersionId { get; set; }
    public string ContentType { get; set; } = "application/octet-stream";
    public Dictionary<string, string> Metadata { get; set; } = new(StringComparer.OrdinalIgnoreCase);
}

/// <summary>A request the fake store answered, with the headers the real SDK sent.</summary>
internal sealed record FakeRequest(string Method, string Name, IReadOnlyDictionary<string, string> Headers, string Query);

/// <summary>
/// An in-memory object store behind an <see cref="HttpMessageHandler"/>, so the real Azure and AWS
/// SDKs run against it: preconditions, metadata and errors take the same wire path as in production.
/// </summary>
internal abstract class FakeObjectStore : HttpMessageHandler
{
    private int _etags;
    private int _versions;

    public Dictionary<string, FakeObject> Objects { get; } = new(StringComparer.Ordinal);
    public List<FakeRequest> Requests { get; } = new();

    /// <summary>Every write answers with this status and stores nothing.</summary>
    public int? WriteStatus { get; set; }
    public string WriteErrorCode { get; set; } = "InvalidRequest";

    /// <summary>The next write is stored, then answered with this status, as a retry after a lost response would be.</summary>
    public int? StoreThenStatus { get; set; }
    public string StoreThenErrorCode { get; set; } = "ServerError";

    /// <summary>The next write is stored, then the response is lost.</summary>
    public bool StoreThenLoseResponse { get; set; }

    /// <summary>Runs before a write's preconditions are evaluated, to simulate a concurrent writer.</summary>
    public Action<string>? BeforeWrite { get; set; }

    /// <summary>When set, every stored object gets a version id, as on a versioning-enabled bucket.</summary>
    public bool Versioning { get; set; }

    public string Seed(string name, byte[] bytes, string? contentType = null, IDictionary<string, string>? metadata = null)
    {
        var item = NewVersion(name, bytes, contentType ?? "application/octet-stream",
            metadata is null ? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) : new Dictionary<string, string>(metadata, StringComparer.OrdinalIgnoreCase));
        return item.ETag;
    }

    /// <summary>Replaces an object's content as another writer would, changing its ETag.</summary>
    public void Touch(string name, byte[]? bytes = null)
    {
        var existing = Objects[name];
        NewVersion(name, bytes ?? existing.Bytes.Concat(new byte[] { 0 }).ToArray(), existing.ContentType, existing.Metadata);
    }

    public bool Has(string name) => Objects.ContainsKey(name);

    public byte[] BytesOf(string name) => Objects[name].Bytes;

    public int Count(string method) => Requests.Count(request => request.Method == method);

    protected FakeObject NewVersion(string name, byte[] bytes, string contentType, Dictionary<string, string> metadata)
    {
        var item = new FakeObject
        {
            Bytes = bytes,
            ETag = NewETag(bytes),
            VersionId = Versioning ? $"v{Interlocked.Increment(ref _versions)}" : null,
            ContentType = contentType,
            Metadata = metadata
        };
        Objects[name] = item;
        return item;
    }

    /// <summary>Azure ETags are opaque and change on every write.</summary>
    protected virtual string NewETag(byte[] bytes) => $"\"0x{Interlocked.Increment(ref _etags):X8}\"";

    protected abstract string? NameOf(Uri uri);

    protected abstract string MetadataPrefix { get; }

    protected abstract HttpResponseMessage Error(HttpStatusCode status, string code, bool head);

    protected abstract void AddObjectHeaders(HttpResponseMessage response, FakeObject item);

    protected abstract string? ContentTypeOf(HttpRequestMessage request);

    protected abstract HttpStatusCode CreatedStatus { get; }

    protected abstract string AlreadyExists(out HttpStatusCode status);

    protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
    {
        var name = NameOf(request.RequestUri!) ?? string.Empty;
        var headers = request.Headers.Concat(request.Content?.Headers ?? Enumerable.Empty<KeyValuePair<string, IEnumerable<string>>>())
            .ToDictionary(pair => pair.Key, pair => string.Join(",", pair.Value), StringComparer.OrdinalIgnoreCase);
        Requests.Add(new FakeRequest(request.Method.Method, name, headers, request.RequestUri!.Query));
        headers.TryGetValue("If-Match", out var ifMatch);
        headers.TryGetValue("If-None-Match", out var ifNoneMatch);
        Objects.TryGetValue(name, out var item);
        var head = request.Method == HttpMethod.Head;

        if (request.Method == HttpMethod.Head || request.Method == HttpMethod.Get)
        {
            if (item is null) return Error(HttpStatusCode.NotFound, "NotFound", head);
            if (ifMatch is not null && ifMatch != item.ETag) return Error(HttpStatusCode.PreconditionFailed, "ConditionNotMet", head);
            var response = new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new ByteArrayContent(head ? Array.Empty<byte>() : item.Bytes)
            };
            response.Content.Headers.ContentLength = item.Bytes.Length;
            AddObjectHeaders(response, item);
            return response;
        }

        if (request.Method != HttpMethod.Put) return Error(HttpStatusCode.MethodNotAllowed, "MethodNotAllowed", head: false);

        var body = request.Content is null ? Array.Empty<byte>() : await request.Content.ReadAsByteArrayAsync(cancellationToken);
        if (WriteStatus is { } refused) return Error((HttpStatusCode)refused, WriteErrorCode, head: false);

        BeforeWrite?.Invoke(name);
        Objects.TryGetValue(name, out item);
        if (ifNoneMatch == "*" && item is not null)
        {
            var code = AlreadyExists(out var status);
            return Error(status, code, head: false);
        }
        if (ifMatch is not null && (item is null || ifMatch != item.ETag))
            return Error(item is null ? HttpStatusCode.NotFound : HttpStatusCode.PreconditionFailed, "ConditionNotMet", head: false);

        var metadata = headers
            .Where(pair => pair.Key.StartsWith(MetadataPrefix, StringComparison.OrdinalIgnoreCase))
            .ToDictionary(pair => pair.Key.Substring(MetadataPrefix.Length), pair => pair.Value, StringComparer.OrdinalIgnoreCase);
        var stored = NewVersion(name, body, ContentTypeOf(request) ?? "application/octet-stream", metadata);

        if (StoreThenLoseResponse)
        {
            StoreThenLoseResponse = false;
            throw new HttpRequestException("The connection was reset after the request was sent.");
        }
        if (StoreThenStatus is { } after)
        {
            StoreThenStatus = null;
            return Error((HttpStatusCode)after, StoreThenErrorCode, head: false);
        }

        var created = new HttpResponseMessage(CreatedStatus) { Content = new ByteArrayContent(Array.Empty<byte>()) };
        created.Headers.TryAddWithoutValidation("ETag", stored.ETag);
        created.Content.Headers.LastModified = DateTimeOffset.UtcNow;
        if (stored.VersionId is not null)
            created.Headers.TryAddWithoutValidation("x-amz-version-id", stored.VersionId);
        return created;
    }

    protected static HttpResponseMessage XmlError(HttpStatusCode status, string code, bool head, string? headerName)
    {
        var response = new HttpResponseMessage(status);
        if (headerName is not null) response.Headers.TryAddWithoutValidation(headerName, code);
        response.Content = head
            ? new ByteArrayContent(Array.Empty<byte>())
            : new StringContent(
                $"<?xml version=\"1.0\" encoding=\"UTF-8\"?><Error><Code>{SecurityElement.Escape(code)}</Code><Message>Fake {code}</Message><RequestId>fake-request</RequestId></Error>",
                System.Text.Encoding.UTF8, "application/xml");
        return response;
    }
}

/// <summary>The Blob service wire protocol over <see cref="FakeObjectStore"/>.</summary>
internal sealed class FakeBlobService : FakeObjectStore
{
    public const string Account = "fakeaccount";
    public const string Container = "docs";
    public static readonly Uri ContainerUri = new($"https://{Account}.blob.core.windows.net/{Container}");

    public BlobContainerClient Client() => new(ContainerUri, new BlobClientOptions
    {
        Transport = new HttpClientTransport(new HttpClient(this, disposeHandler: false)),
        Retry = { MaxRetries = 0 }
    });

    protected override string MetadataPrefix => "x-ms-meta-";

    protected override HttpStatusCode CreatedStatus => HttpStatusCode.Created;

    protected override string? NameOf(Uri uri)
    {
        var path = uri.AbsolutePath;
        var root = "/" + Container + "/";
        return path.StartsWith(root, StringComparison.Ordinal) ? Uri.UnescapeDataString(path.Substring(root.Length)) : null;
    }

    protected override string? ContentTypeOf(HttpRequestMessage request) =>
        request.Headers.TryGetValues("x-ms-blob-content-type", out var values) ? values.First() : null;

    protected override string AlreadyExists(out HttpStatusCode status)
    {
        status = HttpStatusCode.Conflict;
        return "BlobAlreadyExists";
    }

    protected override HttpResponseMessage Error(HttpStatusCode status, string code, bool head) =>
        XmlError(status, code == "NotFound" ? "BlobNotFound" : code, head, "x-ms-error-code");

    protected override void AddObjectHeaders(HttpResponseMessage response, FakeObject item)
    {
        response.Headers.TryAddWithoutValidation("ETag", item.ETag);
        response.Headers.TryAddWithoutValidation("x-ms-blob-type", "BlockBlob");
        response.Content!.Headers.ContentType = MediaTypeHeaderValue.Parse(item.ContentType);
        response.Content.Headers.LastModified = DateTimeOffset.UtcNow;
        foreach (var pair in item.Metadata)
            response.Headers.TryAddWithoutValidation("x-ms-meta-" + pair.Key, pair.Value);
    }
}

/// <summary>The S3 REST protocol (path-style) over <see cref="FakeObjectStore"/>.</summary>
internal sealed class FakeS3Service : FakeObjectStore
{
    public const string Bucket = "fake-bucket";
    public const string Endpoint = "https://s3.fake.test";

    public IAmazonS3 Client() => new AmazonS3Client(new AnonymousAWSCredentials(), new AmazonS3Config
    {
        ServiceURL = Endpoint,
        AuthenticationRegion = "us-east-1",
        ForcePathStyle = true,
        MaxErrorRetry = 0,
        RequestChecksumCalculation = RequestChecksumCalculation.WHEN_REQUIRED,
        ResponseChecksumValidation = ResponseChecksumValidation.WHEN_REQUIRED,
        HttpClientFactory = new Factory(this)
    });

    protected override string MetadataPrefix => "x-amz-meta-";

    protected override HttpStatusCode CreatedStatus => HttpStatusCode.OK;

    /// <summary>A single-part S3 ETag is the content's MD5, and the SDK checks a PUT's ETag against it.</summary>
    protected override string NewETag(byte[] bytes) =>
        "\"" + Convert.ToHexString(System.Security.Cryptography.MD5.HashData(bytes)).ToLowerInvariant() + "\"";

    protected override string? NameOf(Uri uri)
    {
        var path = uri.AbsolutePath;
        var root = "/" + Bucket + "/";
        return path.StartsWith(root, StringComparison.Ordinal) ? Uri.UnescapeDataString(path.Substring(root.Length)) : null;
    }

    protected override string? ContentTypeOf(HttpRequestMessage request) => request.Content?.Headers.ContentType?.ToString();

    protected override string AlreadyExists(out HttpStatusCode status)
    {
        status = HttpStatusCode.PreconditionFailed;
        return "PreconditionFailed";
    }

    protected override HttpResponseMessage Error(HttpStatusCode status, string code, bool head) =>
        XmlError(status, code switch { "NotFound" => "NoSuchKey", "ConditionNotMet" => "PreconditionFailed", _ => code }, head, headerName: null);

    protected override void AddObjectHeaders(HttpResponseMessage response, FakeObject item)
    {
        response.Headers.TryAddWithoutValidation("ETag", item.ETag);
        response.Headers.TryAddWithoutValidation("x-amz-request-id", "fake-request");
        if (item.VersionId is not null)
            response.Headers.TryAddWithoutValidation("x-amz-version-id", item.VersionId);
        response.Content!.Headers.ContentType = MediaTypeHeaderValue.Parse(item.ContentType);
        response.Content.Headers.LastModified = DateTimeOffset.UtcNow;
        foreach (var pair in item.Metadata)
            response.Headers.TryAddWithoutValidation("x-amz-meta-" + pair.Key, pair.Value);
    }

    private sealed class Factory : HttpClientFactory
    {
        private readonly HttpMessageHandler _handler;

        public Factory(HttpMessageHandler handler) => _handler = handler;

        public override HttpClient CreateHttpClient(IClientConfig clientConfig) => new(_handler, disposeHandler: false);

        public override bool UseSDKHttpClientCaching(IClientConfig clientConfig) => false;

        public override bool DisposeHttpClientsAfterUse(IClientConfig clientConfig) => false;

        public override string GetConfigUniqueString(IClientConfig clientConfig) => "fake-s3";
    }
}

/// <summary>
/// A fact that runs only when the named environment variables are set, and is reported as
/// skipped otherwise. Each argument is one requirement; <c>A|B</c> is met by either variable.
/// </summary>
[AttributeUsage(AttributeTargets.Method)]
internal sealed class EnvironmentFactAttribute : FactAttribute
{
    public EnvironmentFactAttribute(params string[] requirements)
    {
        var missing = requirements
            .Where(requirement => requirement.Split('|').All(name => string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable(name))))
            .ToArray();
        if (missing.Length > 0)
            Skip = "Opt-in integration test; set " + string.Join(", ", missing) + " to run it.";
    }

    internal static string? Get(string name)
    {
        var value = Environment.GetEnvironmentVariable(name);
        return string.IsNullOrWhiteSpace(value) ? null : value.Trim();
    }

    internal static bool Flag(string name) =>
        bool.TryParse(Get(name), out var flag) && flag;
}
