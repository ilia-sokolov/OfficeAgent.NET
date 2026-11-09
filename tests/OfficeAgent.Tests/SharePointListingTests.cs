using System.Net;
using System.Text;
using System.Text.Json;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.SharePoint;

namespace OfficeAgent.Tests;

/// <summary>
/// SharePoint listing over a fake Graph drive of folders and files: the children walk, Graph's
/// own paging behind the continuation token, recursion, templates folders, folder addressing by
/// URL, pair or path, and the rule that a page link never carries the bearer token elsewhere.
/// </summary>
public class SharePointListingTests
{
    [Fact]
    public async Task Lists_the_creation_folder_with_sources_registration_accepts()
    {
        using var drive = new FakeDrive();
        drive.File("root", "b.docx");
        drive.File("root", "a.docx");
        drive.File("root", "notes.txt");
        drive.Folder("root", "Contracts");
        var provider = drive.Provider();

        var page = await provider.ListAsync(new DocumentListQuery());

        Assert.Equal(new[] { "b.docx", "a.docx" }, page.Items.Select(i => i.Name));
        Assert.All(page.Items, item =>
        {
            Assert.StartsWith($"{FakeDrive.DriveId}/", item.Source);
            Assert.Null(item.Folder);
            Assert.False(item.IsTemplate);
            Assert.Equal(4, item.Length);
            Assert.Equal(new DateTimeOffset(2026, 10, 1, 9, 0, 0, TimeSpan.Zero), item.LastModified);
        });
        Assert.Null(page.ContinuationToken);

        var registered = await provider.RegisterAsync(page.Items[1].Source);
        Assert.Equal("a.docx", registered.Name);
    }

    [Fact]
    public async Task Graph_pages_and_result_pages_resume_without_repeats()
    {
        using var drive = new FakeDrive { GraphPageSize = 2 };
        foreach (var name in new[] { "1", "2", "3", "4", "5" }) drive.File("root", $"{name}.docx");
        var provider = drive.Provider();

        var seen = new List<string>();
        string? token = null;
        var pages = 0;
        do
        {
            var page = await provider.ListAsync(new DocumentListQuery { MaximumResults = 2, ContinuationToken = token });
            seen.AddRange(page.Items.Select(i => i.Name));
            token = page.ContinuationToken;
            pages++;
        } while (token is not null && pages < 10);

        Assert.Equal(new[] { "1.docx", "2.docx", "3.docx", "4.docx", "5.docx" }, seen);
        Assert.Equal(3, pages);
        Assert.False(drive.SawForeignRequest);
    }

    [Fact]
    public async Task Recursion_reports_folders_relative_to_the_listed_folder()
    {
        using var drive = new FakeDrive();
        drive.File("root", "top.docx");
        var contracts = drive.Folder("root", "Contracts");
        drive.File(contracts, "msa.docx");
        var year = drive.Folder(contracts, "2026");
        drive.File(year, "sow.docx");
        var provider = drive.Provider();

        var flat = await provider.ListAsync(new DocumentListQuery());
        var deep = await provider.ListAsync(new DocumentListQuery { Recursive = true });
        var sub = await provider.ListAsync(new DocumentListQuery { Folder = "Contracts", Recursive = true });

        Assert.Equal(new[] { "top.docx" }, flat.Items.Select(i => i.Name));
        Assert.Equal(new[] { "top.docx:", "msa.docx:Contracts", "sow.docx:Contracts/2026" },
            deep.Items.Select(i => $"{i.Name}:{i.Folder}"));
        Assert.Equal(new[] { "msa.docx:Contracts", "sow.docx:Contracts/2026" },
            sub.Items.Select(i => $"{i.Name}:{i.Folder}"));
    }

    [Fact]
    public async Task Name_extension_macro_and_size_filters_apply()
    {
        using var drive = new FakeDrive();
        drive.File("root", "Engagement Letter.docx");
        drive.File("root", "letter.docm");
        drive.File("root", "letter.pptx");
        drive.File("root", "huge letter.docx", size: 1000);
        var provider = drive.Provider(o =>
        {
            o.AllowedExtensions = new[] { ".docx", ".docm" };
            o.MaximumBytes = 100;
        });

        var page = await provider.ListAsync(new DocumentListQuery { NameContains = "LETTER" });
        var decks = await provider.ListAsync(new DocumentListQuery { Extensions = new[] { ".pptx" } });

        Assert.Equal(new[] { "Engagement Letter.docx" }, page.Items.Select(i => i.Name));
        Assert.Empty(decks.Items);
    }

    [Fact]
    public async Task The_templates_folder_is_listed_as_templates()
    {
        using var drive = new FakeDrive();
        drive.File("root", "work.docx");
        var templates = drive.Folder("root", "Templates");
        drive.File(templates, "Letter.dotx");
        var board = drive.Folder(templates, "Board");
        drive.File(board, "Minutes.docx");
        var provider = drive.Provider(o =>
        {
            o.TemplatesFolder = $"{FakeDrive.DriveId}/{templates}";
            o.TemplateLocation = TemplateLocationKind.Organization;
            o.AllowedExtensions = new[] { ".docx", ".dotx" };
        });

        var page = await provider.ListAsync(new DocumentListQuery { Templates = true, Recursive = true });
        var boardPage = await provider.ListAsync(new DocumentListQuery { Templates = true, Folder = "Board" });

        Assert.Equal(TemplateLocationKind.Organization, provider.TemplateLocation);
        Assert.Equal($"{FakeDrive.DriveId}/{templates}", provider.TemplatesFolder);
        Assert.Equal(new[] { "Letter.dotx:", "Minutes.docx:Board" }, page.Items.Select(i => $"{i.Name}:{i.Folder}"));
        Assert.All(page.Items, item => Assert.True(item.IsTemplate));
        Assert.Equal(new[] { "Minutes.docx" }, boardPage.Items.Select(i => i.Name));
    }

    [Fact]
    public async Task Without_a_templates_folder_a_templates_listing_is_empty_and_sends_nothing()
    {
        using var drive = new FakeDrive();
        drive.File("root", "work.docx");
        var provider = drive.Provider();

        var page = await provider.ListAsync(new DocumentListQuery { Templates = true });

        Assert.Empty(page.Items);
        Assert.Equal(0, drive.Requests);
        Assert.Equal(TemplateLocationKind.None, provider.TemplateLocation);
    }

    [Fact]
    public void A_templates_folder_set_without_a_location_is_a_connection_location()
    {
        using var drive = new FakeDrive();
        var provider = drive.Provider(o => o.TemplatesFolder = "https://contoso.fake/sites/brand/Templates");

        Assert.Equal(TemplateLocationKind.Connection, provider.TemplateLocation);
    }

    [Fact]
    public void A_templates_folder_must_be_a_url_or_a_pair()
    {
        using var drive = new FakeDrive();

        Assert.Throws<ArgumentException>(() => drive.Provider(o => o.TemplatesFolder = "Templates"));
    }

    [Fact]
    public async Task A_folder_is_addressed_by_url_or_by_pair()
    {
        using var drive = new FakeDrive();
        var shared = drive.Folder("root", "Shared");
        drive.File(shared, "deck.docx");
        var provider = drive.Provider(configureCreation: false);

        var byUrl = await provider.ListAsync(new DocumentListQuery { Folder = $"https://contoso.fake/sites/x/{shared}" });
        var byPair = await provider.ListAsync(new DocumentListQuery { Folder = $"drives/{FakeDrive.DriveId}/items/{shared}" });

        Assert.Equal(new[] { "deck.docx" }, byUrl.Items.Select(i => i.Name));
        Assert.Equal(new[] { "deck.docx" }, byPair.Items.Select(i => i.Name));
    }

    [Fact]
    public async Task A_connection_without_a_default_folder_needs_one_and_a_path_cannot_climb()
    {
        using var drive = new FakeDrive();
        drive.Folder("root", "Contracts");

        var noDefault = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            drive.Provider(configureCreation: false).ListAsync(new DocumentListQuery()));
        var climb = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            drive.Provider().ListAsync(new DocumentListQuery { Folder = "Contracts/../.." }));
        var file = drive.File("root", "a.docx");
        var notFolder = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            drive.Provider().ListAsync(new DocumentListQuery { Folder = $"drives/{FakeDrive.DriveId}/items/{file}" }));

        Assert.Equal(ProviderErrorCode.InvalidArgument, noDefault.Code);
        Assert.Equal(ProviderErrorCode.AccessDenied, climb.Code);
        Assert.Equal(ProviderErrorCode.InvalidArgument, notFolder.Code);
    }

    [Fact]
    public async Task A_next_page_link_to_another_host_is_never_followed()
    {
        using var drive = new FakeDrive { GraphPageSize = 1, ForeignNextLink = true };
        drive.File("root", "a.docx");
        drive.File("root", "b.docx");
        var provider = drive.Provider();

        var page = await provider.ListAsync(new DocumentListQuery());

        Assert.Equal(new[] { "a.docx" }, page.Items.Select(i => i.Name));
        Assert.Null(page.ContinuationToken);
        Assert.False(drive.SawForeignRequest);
    }

    [Fact]
    public async Task A_token_pointing_at_another_host_or_another_query_is_refused()
    {
        using var drive = new FakeDrive();
        drive.File("root", "a.docx");
        drive.File("root", "b.docx");
        var provider = drive.Provider();
        var query = new DocumentListQuery();
        var fingerprint = DocumentListingRules.Fingerprint(query, SharePointDocumentProvider.ProviderName);
        var forged = DocumentListingRules.EncodeToken(
            "{\"v\":\"sp1\",\"f\":\"" + fingerprint + "\",\"c\":0,\"cur\":{\"d\":\"drive1\",\"i\":\"root\",\"l\":0}," +
            "\"u\":\"https://evil.example/collect\",\"s\":0,\"q\":[]}");
        var first = await provider.ListAsync(new DocumentListQuery { MaximumResults = 1 });

        var foreign = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.ListAsync(new DocumentListQuery { ContinuationToken = forged }));
        var otherQuery = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.ListAsync(new DocumentListQuery { MaximumResults = 1, Recursive = true, ContinuationToken = first.ContinuationToken }));

        Assert.Equal(ProviderErrorCode.InvalidArgument, foreign.Code);
        Assert.Equal(ProviderErrorCode.InvalidArgument, otherQuery.Code);
        Assert.False(drive.SawForeignRequest);
    }

    [Fact]
    public async Task One_call_sends_a_bounded_number_of_requests_and_resumes_after_them()
    {
        using var drive = new FakeDrive { GraphPageSize = 1 };
        for (var index = 0; index < SharePointDocumentProvider.MaximumListingRequests + 3; index++)
            drive.File("root", $"skip{index:00}.txt");
        drive.File("root", "zz.docx");
        var provider = drive.Provider();

        var first = await provider.ListAsync(new DocumentListQuery());
        Assert.Equal(SharePointDocumentProvider.MaximumListingRequests, drive.Requests);
        Assert.Empty(first.Items);
        Assert.NotNull(first.ContinuationToken);

        var second = await provider.ListAsync(new DocumentListQuery { ContinuationToken = first.ContinuationToken });
        Assert.Equal(new[] { "zz.docx" }, second.Items.Select(i => i.Name));
        Assert.Null(second.ContinuationToken);
    }

    /// <summary>A Graph drive of folders and files, answering only what listing and registration send.</summary>
    private sealed class FakeDrive : IDisposable
    {
        public const string DriveId = "drive1";
        private const string GraphBase = "https://graph.fake/v1.0";
        private readonly Dictionary<string, Item> _items = new(StringComparer.Ordinal);
        private readonly HttpClient _http;
        private int _next;

        public FakeDrive()
        {
            _items["root"] = new Item("root", "root", null, IsFolder: true, 0);
            _http = new HttpClient(new Handler(this));
        }

        public int GraphPageSize { get; init; } = 200;
        public bool ForeignNextLink { get; init; }
        public int Requests { get; private set; }
        public bool SawForeignRequest { get; private set; }

        public string File(string parent, string name, long size = 4) => Add(parent, name, folder: false, size);

        public string Folder(string parent, string name) => Add(parent, name, folder: true, 0);

        public SharePointDocumentProvider Provider(
            Action<SharePointDocumentProviderOptions>? configure = null, bool configureCreation = true)
        {
            var options = new SharePointDocumentProviderOptions
            {
                ConnectionId = "legal",
                GraphBaseUrl = GraphBase,
                CreationDriveId = configureCreation ? DriveId : string.Empty,
                CreationFolderItemId = configureCreation ? "root" : string.Empty
            };
            configure?.Invoke(options);
            return new SharePointDocumentProvider(options, _http, new FakeTokens());
        }

        public void Dispose() => _http.Dispose();

        private string Add(string parent, string name, bool folder, long size)
        {
            var id = $"item{++_next:000}";
            _items[id] = new Item(id, name, parent, IsFolder: folder, size);
            return id;
        }

        private HttpResponseMessage Answer(HttpRequestMessage request)
        {
            var uri = request.RequestUri!;
            if (uri.Host != "graph.fake")
            {
                SawForeignRequest = true;
                return new HttpResponseMessage(HttpStatusCode.OK);
            }
            if (request.Headers.Authorization?.Parameter != "fake-token")
                return new HttpResponseMessage(HttpStatusCode.Unauthorized);
            Requests++;

            var path = uri.AbsolutePath.Substring("/v1.0".Length);
            if (path.StartsWith("/shares/", StringComparison.Ordinal))
            {
                var url = DecodeShare(path.Split('/')[2]);
                return ItemResponse(url.Substring(url.LastIndexOf('/') + 1));
            }

            var prefix = $"/drives/{DriveId}/items/";
            if (!Uri.UnescapeDataString(path).StartsWith(prefix, StringComparison.Ordinal))
                return NotFound();
            var rest = Uri.UnescapeDataString(path).Substring(prefix.Length);

            if (rest.EndsWith("/children", StringComparison.Ordinal))
                return Children(rest.Substring(0, rest.Length - "/children".Length), uri.Query);

            var colon = rest.IndexOf(":/", StringComparison.Ordinal);
            if (colon > 0)
            {
                var current = rest.Substring(0, colon);
                foreach (var segment in rest.Substring(colon + 2).TrimEnd(':').Split('/'))
                {
                    var child = _items.Values.FirstOrDefault(i => i.Parent == current && i.Name == segment);
                    if (child is null) return NotFound();
                    current = child.Id;
                }
                return ItemResponse(current);
            }

            return ItemResponse(rest);
        }

        private HttpResponseMessage Children(string folderId, string query)
        {
            if (!_items.TryGetValue(folderId, out var folder) || !folder.IsFolder) return NotFound();
            var skip = 0;
            const string marker = "$skiptoken=";
            var at = query.IndexOf(marker, StringComparison.Ordinal);
            if (at >= 0) skip = int.Parse(query.Substring(at + marker.Length));

            var children = _items.Values.Where(i => i.Parent == folderId).ToList();
            var page = children.Skip(skip).Take(GraphPageSize).ToList();
            var body = new Dictionary<string, object?> { ["value"] = page.Select(Json).ToList() };
            if (skip + page.Count < children.Count)
                body["@odata.nextLink"] = ForeignNextLink
                    ? $"https://elsewhere.fake/children?$skiptoken={skip + page.Count}"
                    : $"{GraphBase}/drives/{DriveId}/items/{folderId}/children?$skiptoken={skip + page.Count}";
            return Respond(body);
        }

        private HttpResponseMessage ItemResponse(string id) =>
            _items.TryGetValue(id, out var item) ? Respond(Json(item)) : NotFound();

        private static Dictionary<string, object?> Json(Item item)
        {
            var json = new Dictionary<string, object?>
            {
                ["id"] = item.Id,
                ["name"] = item.Name,
                ["eTag"] = $"\"{item.Id}\"",
                ["size"] = item.Size,
                ["lastModifiedDateTime"] = "2026-10-01T09:00:00Z",
                ["parentReference"] = new Dictionary<string, object?> { ["id"] = item.Parent, ["driveId"] = DriveId }
            };
            if (item.IsFolder)
                json["folder"] = new Dictionary<string, object?> { ["childCount"] = 0 };
            else
                json["file"] = new Dictionary<string, object?> { ["mimeType"] = "application/octet-stream" };
            return json;
        }

        private static HttpResponseMessage Respond(object payload) => new(HttpStatusCode.OK)
        {
            Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json")
        };

        private static HttpResponseMessage NotFound() => new(HttpStatusCode.NotFound)
        {
            Content = new StringContent("{\"error\":{\"code\":\"itemNotFound\"}}", Encoding.UTF8, "application/json")
        };

        private static string DecodeShare(string shareId)
        {
            var b64 = shareId.Substring("u!".Length).Replace('_', '/').Replace('-', '+');
            b64 += (b64.Length % 4) switch { 2 => "==", 3 => "=", _ => "" };
            return Encoding.UTF8.GetString(Convert.FromBase64String(b64));
        }

        private sealed record Item(string Id, string Name, string? Parent, bool IsFolder, long Size);

        private sealed class Handler : HttpMessageHandler
        {
            private readonly FakeDrive _drive;
            public Handler(FakeDrive drive) => _drive = drive;

            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) =>
                Task.FromResult(_drive.Answer(request));
        }

        private sealed class FakeTokens : IAccessTokenProvider
        {
            public Task<string> GetAccessTokenAsync(CancellationToken cancellationToken = default) => Task.FromResult("fake-token");
        }
    }
}
