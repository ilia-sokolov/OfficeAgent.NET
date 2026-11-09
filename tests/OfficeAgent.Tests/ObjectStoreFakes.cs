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

    /// <summary>Every listing answers with this status.</summary>
    public int? ListStatus { get; set; }
    public string ListErrorCode { get; set; } = "AccessDenied";

    /// <summary>Answers a listing request in the store's wire format, or returns <see langword="null"/> for any other request.</summary>
    protected abstract HttpResponseMessage? List(HttpRequestMessage request, Dictionary<string, string> query);

    /// <summary>
    /// The names under <paramref name="prefix"/>, in ordinal order: an object directly under it, or
    /// with a delimiter the next level's common prefix, once.
    /// </summary>
    protected List<(string Name, FakeObject? Item)> Enumerate(string prefix, string? delimiter)
    {
        var entries = new SortedDictionary<string, FakeObject?>(StringComparer.Ordinal);
        foreach (var (name, item) in Objects)
        {
            if (!name.StartsWith(prefix, StringComparison.Ordinal)) continue;
            var cut = string.IsNullOrEmpty(delimiter) ? -1 : name.IndexOf(delimiter, prefix.Length, StringComparison.Ordinal);
            if (cut < 0) entries[name] = item;
            else entries[name.Substring(0, cut + delimiter!.Length)] = null;
        }
        return entries.Select(pair => (pair.Key, pair.Value)).ToList();
    }

    protected static Dictionary<string, string> QueryOf(Uri uri) => uri.Query.TrimStart('?')
        .Split('&', StringSplitOptions.RemoveEmptyEntries)
        .Select(pair => pair.Split('=', 2))
        .ToDictionary(pair => Uri.UnescapeDataString(pair[0]), pair => pair.Length > 1 ? Uri.UnescapeDataString(pair[1].Replace('+', ' ')) : string.Empty);

    protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
    {
        if (request.Method == HttpMethod.Get)
        {
            var query = QueryOf(request.RequestUri!);
            if (List(request, query) is { } listing)
            {
                Requests.Add(new FakeRequest("LIST", query.TryGetValue("prefix", out var listed) ? listed : string.Empty,
                    new Dictionary<string, string>(), request.RequestUri!.Query));
                if (ListStatus is { } listStatus)
                {
                    listing.Dispose();
                    return Error((HttpStatusCode)listStatus, ListErrorCode, head: false);
                }
                return listing;
            }
        }

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

    /// <summary>List Blobs (<c>comp=list</c>): the marker is the index of the next entry.</summary>
    protected override HttpResponseMessage? List(HttpRequestMessage request, Dictionary<string, string> query)
    {
        if (!query.TryGetValue("comp", out var comp) || comp != "list") return null;
        var prefix = query.TryGetValue("prefix", out var p) ? p : string.Empty;
        var delimiter = query.TryGetValue("delimiter", out var d) ? d : null;
        var max = query.TryGetValue("maxresults", out var m) ? int.Parse(m) : 5000;
        var start = query.TryGetValue("marker", out var marker) && marker.Length > 0 ? int.Parse(marker) : 0;
        var entries = Enumerate(prefix, delimiter);
        var page = entries.Skip(start).Take(max).ToList();
        var next = start + max < entries.Count ? (start + max).ToString() : string.Empty;

        var xml = new System.Text.StringBuilder("<?xml version=\"1.0\" encoding=\"utf-8\"?>")
            .Append($"<EnumerationResults ServiceEndpoint=\"https://{Account}.blob.core.windows.net/\" ContainerName=\"{Container}\">")
            .Append($"<Prefix>{SecurityElement.Escape(prefix)}</Prefix><MaxResults>{max}</MaxResults>")
            .Append(delimiter is null ? string.Empty : $"<Delimiter>{SecurityElement.Escape(delimiter)}</Delimiter>")
            .Append("<Blobs>");
        foreach (var (name, item) in page)
        {
            if (item is null)
            {
                xml.Append($"<BlobPrefix><Name>{SecurityElement.Escape(name)}</Name></BlobPrefix>");
                continue;
            }
            xml.Append($"<Blob><Name>{SecurityElement.Escape(name)}</Name><Properties>")
                .Append($"<Last-Modified>{Modified.ToString("R")}</Last-Modified>")
                .Append($"<Etag>{SecurityElement.Escape(item.ETag.Trim('"'))}</Etag>")
                .Append($"<Content-Length>{item.Bytes.Length}</Content-Length>")
                .Append($"<Content-Type>{SecurityElement.Escape(item.ContentType)}</Content-Type>")
                .Append("<BlobType>BlockBlob</BlobType></Properties></Blob>");
        }
        xml.Append($"</Blobs><NextMarker>{next}</NextMarker></EnumerationResults>");
        return new HttpResponseMessage(HttpStatusCode.OK)
        {
            Content = new StringContent(xml.ToString(), System.Text.Encoding.UTF8, "application/xml")
        };
    }

    /// <summary>The last-modified time every listed blob reports.</summary>
    public static readonly DateTimeOffset Modified = new(2026, 11, 1, 9, 0, 0, TimeSpan.Zero);

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

    /// <summary>ListObjectsV2 (<c>list-type=2</c>): the continuation token is the index of the next entry.</summary>
    protected override HttpResponseMessage? List(HttpRequestMessage request, Dictionary<string, string> query)
    {
        if (!query.TryGetValue("list-type", out var type) || type != "2") return null;
        var prefix = query.TryGetValue("prefix", out var p) ? p : string.Empty;
        var delimiter = query.TryGetValue("delimiter", out var d) ? d : null;
        var max = query.TryGetValue("max-keys", out var m) ? int.Parse(m) : 1000;
        var start = query.TryGetValue("continuation-token", out var token) ? int.Parse(token) : 0;
        var encode = query.TryGetValue("encoding-type", out var encoding) && encoding == "url";
        string Text(string value) => SecurityElement.Escape(encode ? Uri.EscapeDataString(value) : value)!;
        var entries = Enumerate(prefix, delimiter);
        var page = entries.Skip(start).Take(max).ToList();
        var truncated = start + max < entries.Count;

        var xml = new System.Text.StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\"?>")
            .Append("<ListBucketResult xmlns=\"http://s3.amazonaws.com/doc/2006-03-01/\">")
            .Append($"<Name>{Bucket}</Name><Prefix>{Text(prefix)}</Prefix><KeyCount>{page.Count}</KeyCount><MaxKeys>{max}</MaxKeys>")
            .Append(delimiter is null ? string.Empty : $"<Delimiter>{Text(delimiter)}</Delimiter>")
            .Append(encode ? "<EncodingType>url</EncodingType>" : string.Empty)
            .Append($"<IsTruncated>{(truncated ? "true" : "false")}</IsTruncated>")
            .Append(token is null ? string.Empty : $"<ContinuationToken>{token}</ContinuationToken>")
            .Append(truncated ? $"<NextContinuationToken>{start + max}</NextContinuationToken>" : string.Empty);
        foreach (var (name, item) in page.Where(entry => entry.Item is not null))
            xml.Append($"<Contents><Key>{Text(name)}</Key><LastModified>2026-11-01T09:00:00.000Z</LastModified>")
                .Append($"<ETag>{SecurityElement.Escape(item!.ETag)}</ETag><Size>{item.Bytes.Length}</Size><StorageClass>STANDARD</StorageClass></Contents>");
        foreach (var (name, _) in page.Where(entry => entry.Item is null))
            xml.Append($"<CommonPrefixes><Prefix>{Text(name)}</Prefix></CommonPrefixes>");
        xml.Append("</ListBucketResult>");
        return new HttpResponseMessage(HttpStatusCode.OK)
        {
            Content = new StringContent(xml.ToString(), System.Text.Encoding.UTF8, "application/xml")
        };
    }

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
