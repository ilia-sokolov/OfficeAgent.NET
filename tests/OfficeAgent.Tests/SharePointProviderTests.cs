using System.Net;
using System.Text;
using System.Text.Json;
using OfficeAgent.Abstractions;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.SharePoint;

namespace OfficeAgent.Tests;

/// <summary>
/// SharePoint provider tests over a fake Graph drive: drive items simulated by an
/// in-memory <see cref="HttpMessageHandler"/>, so the full register → open → save →
/// remove loop (including registration by URL and by driveId/itemId, ETag
/// concurrency, and the no-credentials-to-storage rule) runs without a tenant.
/// </summary>
public class SharePointProviderTests
{
    [Fact]
    public async Task Register_by_drive_and_item_id_assigns_opaque_id_and_opens_by_it()
    {
        using var drive = new FakeGraphDrive();
        var itemId = drive.Seed("contract.docx", DocxFactory.Contract());
        var provider = drive.Provider();

        var added = await provider.RegisterAsync(drive.SourceById(itemId));

        Assert.Equal("sharepoint", added.Provider);
        Assert.NotEqual(itemId, added.ItemId);
        Assert.DoesNotContain('/', added.ItemId);
        Assert.Equal("contract.docx", added.Name);
        Assert.False(string.IsNullOrEmpty(added.Version));

        using var content = await provider.OpenReadAsync(added);
        Assert.Equal(added.ItemId, content.Reference.ItemId);
        Assert.True(content.Stream.Length > 0);
    }

    [Fact]
    public async Task Register_by_url_resolves_the_drive_item()
    {
        using var drive = new FakeGraphDrive();
        var itemId = drive.Seed("contract.docx", DocxFactory.Contract());
        var provider = drive.Provider();

        var added = await provider.RegisterAsync(drive.UrlFor(itemId));

        Assert.Equal("contract.docx", added.Name);
        using var content = await provider.OpenReadAsync(added);
        Assert.True(content.Stream.Length > 0);
    }

    [Fact]
    public async Task Register_accepts_the_graph_resource_path_form()
    {
        using var drive = new FakeGraphDrive();
        var itemId = drive.Seed("contract.docx", DocxFactory.Contract());
        var provider = drive.Provider();

        var added = await provider.RegisterAsync($"drives/{FakeGraphDrive.DriveId}/items/{itemId}");

        Assert.Equal("contract.docx", added.Name);
    }

    [Fact]
    public async Task Register_rejects_bare_names_and_disallowed_extensions()
    {
        using var drive = new FakeGraphDrive();
        drive.Seed("contract.docx", DocxFactory.Contract());
        var exeId = drive.Seed("contract.exe", DocxFactory.Contract());
        var provider = drive.Provider();

        // A bare name with no driveId/itemId delimiter is not a valid source.
        var bare = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.RegisterAsync("contract.docx"));
        Assert.Equal(ProviderErrorCode.InvalidArgument, bare.Code);

        // An empty item id half is rejected too.
        var empty = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.RegisterAsync($"{FakeGraphDrive.DriveId}/"));
        Assert.Equal(ProviderErrorCode.InvalidArgument, empty.Code);

        // The extension is enforced against the resolved item name.
        var extension = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.RegisterAsync(drive.SourceById(exeId)));
        Assert.Equal(ProviderErrorCode.ExtensionNotAllowed, extension.Code);
    }

    [Fact]
    public async Task Register_surfaces_not_found_for_missing_items()
    {
        using var drive = new FakeGraphDrive();
        var provider = drive.Provider();

        var ex = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.RegisterAsync($"{FakeGraphDrive.DriveId}/item-missing"));
        Assert.Equal(ProviderErrorCode.NotFound, ex.Code);
    }

    [Fact]
    public async Task Save_new_version_mints_sibling_with_versioned_name_and_fresh_id()
    {
        using var drive = new FakeGraphDrive();
        var itemId = drive.Seed("contract.docx", DocxFactory.Contract());
        var provider = drive.Provider();
        var source = await provider.RegisterAsync(drive.SourceById(itemId));

        using var output = new MemoryStream(DocxFactory.Contract());
        var saved = await provider.SaveAsync(source, output, new SaveDocumentOptions());

        Assert.Equal("contract.v2.docx", saved.Name);
        Assert.NotEqual(source.ItemId, saved.ItemId);

        using var roundTrip = await provider.OpenReadAsync(saved);
        Assert.Equal("contract.v2.docx", roundTrip.Reference.Name);
    }

    [Fact]
    public async Task Save_replace_overwrites_in_place_and_rejects_stale_etag()
    {
        using var drive = new FakeGraphDrive();
        var itemId = drive.Seed("contract.docx", DocxFactory.Contract());
        var provider = drive.Provider();
        var source = await provider.RegisterAsync(drive.SourceById(itemId));

        using (var fresh = new MemoryStream(DocxFactory.Contract().Concat(new byte[] { 0 }).ToArray()))
        {
            var replaced = await provider.SaveAsync(source, fresh, new SaveDocumentOptions { Mode = SaveMode.Replace });
            Assert.Equal(source.ItemId, replaced.ItemId);
            Assert.NotEqual(source.Version, replaced.Version);
        }

        // The reference still carries the original (now stale) ETag.
        using var output = new MemoryStream(DocxFactory.Contract());
        await Assert.ThrowsAsync<DocumentVersionConflictException>(() =>
            provider.SaveAsync(source, output, new SaveDocumentOptions { Mode = SaveMode.Replace }));
    }

    [Fact]
    public async Task Save_rejects_caller_chosen_destinations_and_replace_renames()
    {
        using var drive = new FakeGraphDrive();
        var itemId = drive.Seed("contract.docx", DocxFactory.Contract());
        var provider = drive.Provider();
        var source = await provider.RegisterAsync(drive.SourceById(itemId));

        using var output = new MemoryStream(DocxFactory.Contract());
        var destination = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.SaveAsync(
            source, output, new SaveDocumentOptions { DestinationItemId = "chosen-id" }));
        Assert.Equal(ProviderErrorCode.InvalidArgument, destination.Code);

        var rename = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.SaveAsync(
            source, output, new SaveDocumentOptions { Mode = SaveMode.Replace, NewName = "renamed.docx" }));
        Assert.Equal(ProviderErrorCode.InvalidArgument, rename.Code);
    }

    [Fact]
    public async Task Save_new_document_honours_validated_new_name()
    {
        using var drive = new FakeGraphDrive();
        var itemId = drive.Seed("contract.docx", DocxFactory.Contract());
        var provider = drive.Provider();
        var source = await provider.RegisterAsync(drive.SourceById(itemId));

        using var output = new MemoryStream(DocxFactory.Contract());
        var saved = await provider.SaveAsync(source, output, new SaveDocumentOptions
        {
            Mode = SaveMode.NewDocument,
            NewName = "globex-contract.docx"
        });
        Assert.Equal("globex-contract.docx", saved.Name);

        using var bad = new MemoryStream(DocxFactory.Contract());
        var ex = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.SaveAsync(
            source, bad, new SaveDocumentOptions { Mode = SaveMode.NewDocument, NewName = "nested/name.docx" }));
        Assert.Equal(ProviderErrorCode.InvalidArgument, ex.Code);
    }

    [Fact]
    public async Task Remove_drops_registration_but_not_sharepoint_content()
    {
        using var drive = new FakeGraphDrive();
        var itemId = drive.Seed("contract.docx", DocxFactory.Contract());
        var provider = drive.Provider();
        var added = await provider.RegisterAsync(drive.SourceById(itemId));

        await provider.RemoveAsync(added);

        var ex = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.OpenReadAsync(added));
        Assert.Equal(ProviderErrorCode.NotFound, ex.Code);
        // The drive item survives: a new registration of the same item still works.
        var again = await provider.RegisterAsync(drive.SourceById(itemId));
        Assert.NotEqual(added.ItemId, again.ItemId);
    }

    [Fact]
    public async Task Open_rejects_oversized_content()
    {
        using var drive = new FakeGraphDrive();
        var itemId = drive.Seed("contract.docx", DocxFactory.Contract());
        var provider = drive.Provider(maximumBytes: 16);

        var ex = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.RegisterAsync(drive.SourceById(itemId)));
        Assert.Equal(ProviderErrorCode.ContentTooLarge, ex.Code);
    }

    [Fact]
    public async Task Bearer_token_never_travels_to_the_download_host()
    {
        using var drive = new FakeGraphDrive();
        var itemId = drive.Seed("contract.docx", DocxFactory.Contract());
        var provider = drive.Provider();
        var added = await provider.RegisterAsync(drive.SourceById(itemId));

        using var _ = await provider.OpenReadAsync(added);

        Assert.True(drive.SawAuthorizedGraphCall);
        Assert.False(drive.SawAuthorizedDownloadCall);
    }

    [Fact]
    public async Task Json_file_registration_store_survives_reload()
    {
        var path = Path.Combine(Path.GetTempPath(), $"officeagent-spindex-{Guid.NewGuid():N}", "index.json");
        try
        {
            var store = new JsonFileRegistrationStore(path);
            var id = await store.AddAsync(new SharePointItemRef("drive-1", "item-1"));

            var reloaded = new JsonFileRegistrationStore(path);
            Assert.Equal(new SharePointItemRef("drive-1", "item-1"), await reloaded.ResolveAsync(id));

            Assert.True(await reloaded.RemoveAsync(id));
            var reloadedAgain = new JsonFileRegistrationStore(path);
            Assert.Null(await reloadedAgain.ResolveAsync(id));
        }
        finally
        {
            var directory = Path.GetDirectoryName(path)!;
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }

    // ── Fake Graph drive ──────────────────────────────────────────────────────

    /// <summary>
    /// Simulates one Graph drive (`/drives/{id}`) addressed by item id, plus the
    /// `/shares/{id}/driveItem` URL-resolution endpoint, with ETags,
    /// conflictBehavior=fail uploads, and pre-authenticated download URLs on a
    /// separate host. Items report their parent drive id, so registration by URL or
    /// by driveId/itemId both resolve a fully-addressed item.
    /// </summary>
    private sealed class FakeGraphDrive : IDisposable
    {
        private const string GraphBase = "https://graph.fake/v1.0";
        public const string DriveId = "fakedrive";
        private const string DownloadHost = "https://download.fake";
        private const string ShareHost = "https://contoso.fake/doc";
        private const string FolderId = "folder-contracts";

        private readonly Dictionary<string, Entry> _byName = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, Entry> _byId = new(StringComparer.Ordinal);
        private readonly HttpClient _http;
        private int _nextId;
        private int _nextETag;

        public FakeGraphDrive() => _http = new HttpClient(new Handler(this));

        public bool SawAuthorizedGraphCall { get; private set; }
        public bool SawAuthorizedDownloadCall { get; private set; }

        /// <summary>Seeds a drive item and returns its Graph item id.</summary>
        public string Seed(string name, byte[] bytes)
        {
            var entry = new Entry
            {
                Id = $"item-{++_nextId}",
                Name = name,
                Bytes = bytes,
                ETag = NextETag()
            };
            _byName[name] = entry;
            _byId[entry.Id] = entry;
            return entry.Id;
        }

        /// <summary>A driveId/itemId registration source for a seeded item.</summary>
        public string SourceById(string itemId) => $"{DriveId}/{itemId}";

        /// <summary>A sharing-style URL registration source for a seeded item.</summary>
        public string UrlFor(string itemId) => $"{ShareHost}/{itemId}";

        public SharePointDocumentProvider Provider(long maximumBytes = 100 * 1024 * 1024) => new(
            new SharePointDocumentProviderOptions
            {
                ConnectionId = "legal",
                GraphBaseUrl = GraphBase,
                MaximumBytes = maximumBytes
            },
            _http,
            new FakeTokens());

        public void Dispose() => _http.Dispose();

        /// <summary>A fixed Graph token for tests; production has no static-token path.</summary>
        private sealed class FakeTokens : IAccessTokenProvider
        {
            public Task<string> GetAccessTokenAsync(CancellationToken cancellationToken = default) =>
                Task.FromResult("fake-token");
        }

        private string NextETag() => $"\"etag-{++_nextETag}\"";

        private sealed class Entry
        {
            public string Id = "";
            public string Name = "";
            public byte[] Bytes = Array.Empty<byte>();
            public string ETag = "";
        }

        private sealed class Handler : HttpMessageHandler
        {
            private readonly FakeGraphDrive _drive;

            public Handler(FakeGraphDrive drive) => _drive = drive;

            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
            {
                var url = request.RequestUri!.ToString();
                var authorized = request.Headers.Authorization is { Scheme: "Bearer", Parameter: "fake-token" };

                if (url.StartsWith(DownloadHost, StringComparison.Ordinal))
                {
                    if (request.Headers.Authorization is not null) _drive.SawAuthorizedDownloadCall = true;
                    var id = url.Substring(url.LastIndexOf('/') + 1);
                    return Task.FromResult(_drive._byId.TryGetValue(id, out var file)
                        ? Bytes(file.Bytes)
                        : Status(HttpStatusCode.NotFound));
                }

                if (!authorized)
                    return Task.FromResult(Status(HttpStatusCode.Unauthorized));
                _drive.SawAuthorizedGraphCall = true;

                // GET /shares/{shareId}/driveItem - resolve a URL to its drive item.
                var sharesPrefix = $"{GraphBase}/shares/";
                if (url.StartsWith(sharesPrefix, StringComparison.Ordinal))
                {
                    var rest = url.Substring(sharesPrefix.Length);
                    var shareId = rest.Substring(0, rest.IndexOf('/'));
                    var decoded = DecodeShareId(shareId);
                    var id = decoded.Substring(decoded.LastIndexOf('/') + 1);
                    return Task.FromResult(_drive._byId.TryGetValue(id, out var entry)
                        ? Item(entry)
                        : GraphError(HttpStatusCode.NotFound, "itemNotFound"));
                }

                var prefix = $"{GraphBase}/drives/{Uri.EscapeDataString(DriveId)}";
                if (!url.StartsWith(prefix, StringComparison.Ordinal))
                    return Task.FromResult(Status(HttpStatusCode.NotFound));

                return Task.FromResult(Route(request, url.Substring(prefix.Length)));
            }

            private HttpResponseMessage Route(HttpRequestMessage request, string rest)
            {
                if (!rest.StartsWith("/items/", StringComparison.Ordinal))
                    return Status(HttpStatusCode.NotFound);
                rest = rest.Substring("/items/".Length);

                // PUT /items/{folderId}:/{name}:/content?@microsoft.graph.conflictBehavior=fail
                if (request.Method == HttpMethod.Put && rest.Contains(":/"))
                {
                    var name = Uri.UnescapeDataString(rest.Substring(
                        rest.IndexOf(":/", StringComparison.Ordinal) + 2)
                        .Replace(":/content?@microsoft.graph.conflictBehavior=fail", ""));
                    if (_drive._byName.ContainsKey(name))
                        return GraphError(HttpStatusCode.Conflict, "nameAlreadyExists");
                    var bytes = request.Content!.ReadAsByteArrayAsync().GetAwaiter().GetResult();
                    var id = _drive.Seed(name, bytes);
                    return Item(_drive._byId[id]);
                }

                // GET /items/{folderId}:/{name} - existence probe for versioned names
                if (request.Method == HttpMethod.Get && rest.Contains(":/"))
                {
                    var name = Uri.UnescapeDataString(rest.Substring(rest.IndexOf(":/", StringComparison.Ordinal) + 2));
                    return _drive._byName.TryGetValue(name, out var probe)
                        ? Item(probe)
                        : GraphError(HttpStatusCode.NotFound, "itemNotFound");
                }

                // PUT /items/{id}/content - replace
                if (request.Method == HttpMethod.Put && rest.EndsWith("/content", StringComparison.Ordinal))
                {
                    var id = rest.Substring(0, rest.Length - "/content".Length);
                    if (!_drive._byId.TryGetValue(id, out var entry))
                        return GraphError(HttpStatusCode.NotFound, "itemNotFound");
                    if (request.Headers.TryGetValues("If-Match", out var values) &&
                        !values.Contains(entry.ETag, StringComparer.Ordinal))
                        return GraphError(HttpStatusCode.PreconditionFailed, "resourceModified");
                    entry.Bytes = request.Content!.ReadAsByteArrayAsync().GetAwaiter().GetResult();
                    entry.ETag = _drive.NextETag();
                    return Item(entry);
                }

                // GET /items/{id}
                if (request.Method == HttpMethod.Get)
                {
                    return _drive._byId.TryGetValue(rest, out var entry)
                        ? Item(entry)
                        : GraphError(HttpStatusCode.NotFound, "itemNotFound");
                }

                return Status(HttpStatusCode.MethodNotAllowed);
            }

            private static string DecodeShareId(string shareId)
            {
                var b64 = shareId.Substring("u!".Length).Replace('_', '/').Replace('-', '+');
                switch (b64.Length % 4)
                {
                    case 2: b64 += "=="; break;
                    case 3: b64 += "="; break;
                }
                return Encoding.UTF8.GetString(Convert.FromBase64String(b64));
            }

            private static HttpResponseMessage Item(Entry entry) => Json(new Dictionary<string, object?>
            {
                ["id"] = entry.Id,
                ["name"] = entry.Name,
                ["eTag"] = entry.ETag,
                ["size"] = entry.Bytes.LongLength,
                ["parentReference"] = new Dictionary<string, object?>
                {
                    ["id"] = FolderId,
                    ["driveId"] = DriveId
                },
                ["file"] = new Dictionary<string, object?>
                {
                    ["mimeType"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                },
                ["@microsoft.graph.downloadUrl"] = $"{DownloadHost}/{entry.Id}"
            });

            private static HttpResponseMessage Json(object payload) => new(HttpStatusCode.OK)
            {
                Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json")
            };

            private static HttpResponseMessage Bytes(byte[] bytes) => new(HttpStatusCode.OK)
            {
                Content = new ByteArrayContent(bytes)
            };

            private static HttpResponseMessage Status(HttpStatusCode code) => new(code);

            private static HttpResponseMessage GraphError(HttpStatusCode status, string code) => new(status)
            {
                Content = new StringContent(
                    JsonSerializer.Serialize(new { error = new { code, message = code } }),
                    Encoding.UTF8, "application/json")
            };
        }
    }
}
