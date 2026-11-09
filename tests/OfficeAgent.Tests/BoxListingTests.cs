using System.Net;
using OfficeAgent.Box;
using OfficeAgent.Core.DocumentProviders;
using FakeBox = OfficeAgent.Tests.BoxProviderTests.FakeBox;

namespace OfficeAgent.Tests;

/// <summary>
/// Listing a Box connection over the fake Box API: <c>GET /folders/{id}/items</c> with marker
/// paging, folder paths resolved by name, filters, templates, and the bounds that keep a listing
/// inside the configured folder.
/// </summary>
public class BoxListingTests
{
    private static int ItemRequests(FakeBox box) => box.Requests.Count(r => r.Contains("/items?", StringComparison.Ordinal));

    [Fact]
    public async Task Listing_pages_across_box_pages_and_resumes_inside_one()
    {
        using var box = new FakeBox();
        for (var i = 0; i < 205; i++) box.Seed($"doc{i:000}.docx", new byte[] { 1 });
        box.Seed("elsewhere.docx", new byte[] { 1 }, parent: "9999");
        var provider = box.Provider();

        var first = await provider.ListAsync(new DocumentListQuery { MaximumResults = 150 });
        var second = await provider.ListAsync(new DocumentListQuery { MaximumResults = 150, ContinuationToken = first.ContinuationToken });

        Assert.Equal(150, first.Items.Count);
        Assert.NotNull(first.ContinuationToken);
        Assert.Equal(55, second.Items.Count);
        Assert.Null(second.ContinuationToken);
        Assert.Equal(Enumerable.Range(0, 205).Select(i => $"doc{i:000}.docx"),
            first.Items.Concat(second.Items).Select(item => item.Name));
        Assert.Equal(3, ItemRequests(box));
        Assert.All(box.Requests.Where(r => r.Contains("/items?", StringComparison.Ordinal)),
            r => Assert.Contains("limit=200&usemarker=true", r));
    }

    [Fact]
    public async Task Listing_filters_by_name_and_extension_and_skips_what_cannot_be_registered()
    {
        using var box = new FakeBox();
        box.Seed("Acme NDA.docx", new byte[] { 1 });
        box.Seed("acme deck.pptx", new byte[] { 1 });
        box.Seed("acme budget.xlsx", new byte[] { 1 });
        box.Seed("acme macro.docm", new byte[] { 1 });
        box.Seed("acme link.docx", new byte[] { 1 }, type: "web_link");
        box.Seed("acme big.docx", new byte[4096]);
        box.Trash(box.Seed("acme trashed.docx", new byte[] { 1 }));
        box.Seed("other.docx", new byte[] { 1 });
        var provider = box.Provider(maximumBytes: 1024, allowedExtensions: new[] { ".docx", ".pptx", ".docm" });

        var all = await provider.ListAsync(new DocumentListQuery());
        var named = await provider.ListAsync(new DocumentListQuery { NameContains = "ACME" });
        var narrowed = await provider.ListAsync(new DocumentListQuery { Extensions = new[] { ".pptx", ".xlsx" } });

        Assert.Equal(new[] { "Acme NDA.docx", "acme deck.pptx", "other.docx" }, all.Items.Select(item => item.Name));
        Assert.Equal(new[] { "Acme NDA.docx", "acme deck.pptx" }, named.Items.Select(item => item.Name));
        Assert.Equal("acme deck.pptx", Assert.Single(narrowed.Items).Name);
        var entry = all.Items[0];
        Assert.Equal(1, entry.Length);
        Assert.Equal(new DateTimeOffset(2026, 11, 1, 17, 0, 0, TimeSpan.Zero), entry.LastModified);
    }

    [Fact]
    public async Task Folder_paths_resolve_by_name_and_recursion_names_each_folder()
    {
        using var box = new FakeBox();
        box.Seed("top.docx", new byte[] { 1 });
        var clients = box.SeedFolder("Clients");
        var acme = box.SeedFolder("Acme", clients);
        box.Seed("mid.docx", new byte[] { 1 }, parent: clients);
        box.Seed("deep.docx", new byte[] { 1 }, parent: acme);
        var provider = box.Provider();

        var flat = await provider.ListAsync(new DocumentListQuery());
        var recursive = await provider.ListAsync(new DocumentListQuery { Recursive = true });
        var nested = await provider.ListAsync(new DocumentListQuery { Folder = "clients/acme" });
        var missing = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.ListAsync(new DocumentListQuery { Folder = "Clients/Missing" }));

        Assert.Equal("top.docx", Assert.Single(flat.Items).Name);
        Assert.Equal(
            new[] { ("top.docx", (string?)null), ("mid.docx", "Clients"), ("deep.docx", "Clients/Acme") },
            recursive.Items.Select(item => (item.Name, item.Folder)));
        Assert.Equal(("deep.docx", "clients/acme"), (Assert.Single(nested.Items).Name, nested.Items[0].Folder));
        Assert.Equal(ProviderErrorCode.NotFound, missing.Code);
    }

    [Fact]
    public async Task Every_listed_source_registers_and_opens()
    {
        using var box = new FakeBox();
        box.Seed("contract.docx", DocxFactory.Contract());
        box.Seed("nested.docx", DocxFactory.Contract(), parent: box.SeedFolder("Sub"));
        var provider = box.Provider();

        var listed = await provider.ListAsync(new DocumentListQuery { Recursive = true });

        Assert.Equal(2, listed.Items.Count);
        foreach (var entry in listed.Items)
        {
            var reference = await provider.RegisterAsync(entry.Source);
            Assert.Equal(entry.Name, reference.Name);
            using var opened = await provider.OpenReadAsync(reference);
            Assert.True(opened.Stream.Length > 0);
        }
    }

    [Fact]
    public async Task A_templates_listing_reads_the_templates_folder_and_flags_every_entry()
    {
        using var box = new FakeBox();
        var templates = box.SeedFolder("Firm templates", parent: "0");
        box.Seed("letter.dotx", new byte[] { 1 }, parent: templates);
        box.Seed("memo.docx", new byte[] { 1 }, parent: templates);
        box.Seed("plan.dotx", new byte[] { 1 });
        box.Seed("draft.docx", new byte[] { 1 });
        var provider = new BoxDocumentProvider(new BoxDocumentProviderOptions
        {
            ConnectionId = "box", ApiBaseUrl = FakeBox.Api, UploadBaseUrl = FakeBox.Upload, CreationFolderId = FakeBox.FolderId,
            TemplatesFolder = templates, TemplateLocation = TemplateLocationKind.Organization,
            AllowedExtensions = new[] { ".docx", ".dotx" }
        }, box.Http, new StaticTokens());
        var withoutTemplates = box.Provider();

        var listed = await provider.ListAsync(new DocumentListQuery { Templates = true });
        var documents = await provider.ListAsync(new DocumentListQuery());
        var requests = box.Requests.Count;
        var none = await withoutTemplates.ListAsync(new DocumentListQuery { Templates = true });

        Assert.Equal(templates, provider.TemplatesFolder);
        Assert.Equal(TemplateLocationKind.Organization, provider.TemplateLocation);
        Assert.Equal(new[] { ("letter.dotx", true), ("memo.docx", true) }, listed.Items.Select(item => (item.Name, item.IsTemplate)));
        Assert.Equal(new[] { ("plan.dotx", true), ("draft.docx", false) }, documents.Items.Select(item => (item.Name, item.IsTemplate)));
        Assert.Equal(TemplateLocationKind.None, withoutTemplates.TemplateLocation);
        Assert.Empty(none.Items);
        Assert.Equal(requests, box.Requests.Count);
    }

    [Fact]
    public async Task Listing_stays_inside_the_folder_and_its_tokens_cannot_be_redirected()
    {
        using var box = new FakeBox();
        for (var i = 0; i < 3; i++) box.Seed($"doc{i}.docx", new byte[] { 1 });
        box.SeedFolder("Sub");
        var provider = box.Provider();
        var token = (await provider.ListAsync(new DocumentListQuery { MaximumResults = 1 })).ContinuationToken!;
        var requests = box.Requests.Count;

        var escape = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.ListAsync(new DocumentListQuery { Folder = "Sub/../.." }));
        var reused = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.ListAsync(new DocumentListQuery { Folder = "Sub", MaximumResults = 1, ContinuationToken = token }));
        var otherConnection = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            new BoxDocumentProvider(new BoxDocumentProviderOptions
            {
                ConnectionId = "box-2", ApiBaseUrl = FakeBox.Api, UploadBaseUrl = FakeBox.Upload, CreationFolderId = "4242"
            }, box.Http, new StaticTokens()).ListAsync(new DocumentListQuery { MaximumResults = 1, ContinuationToken = token }));

        Assert.Equal(ProviderErrorCode.AccessDenied, escape.Code);
        Assert.Equal(ProviderErrorCode.InvalidArgument, reused.Code);
        Assert.Equal(ProviderErrorCode.InvalidArgument, otherConnection.Code);
        Assert.Equal(requests, box.Requests.Count);
    }

    [Fact]
    public async Task Listing_failures_map_to_provider_errors()
    {
        using var box = new FakeBox();
        box.Seed("contract.docx", new byte[] { 1 });

        box.ListStatus = (HttpStatusCode.Forbidden, "access_denied_insufficient_permissions");
        var denied = await Assert.ThrowsAsync<DocumentProviderException>(() => box.Provider().ListAsync(new DocumentListQuery()));
        box.ListStatus = null;
        var unconfigured = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            box.Provider(configureCreation: false).ListAsync(new DocumentListQuery()));
        box.RemoveFolder();
        var missing = await Assert.ThrowsAsync<DocumentProviderException>(() => box.Provider().ListAsync(new DocumentListQuery()));

        Assert.Equal(ProviderErrorCode.AccessDenied, denied.Code);
        Assert.Equal(ProviderErrorCode.ConfigurationError, unconfigured.Code);
        Assert.Equal(ProviderErrorCode.NotFound, missing.Code);
        Assert.Throws<ArgumentException>(() => box.Provider(templatesFolder: "Templates"));
    }

    private sealed class StaticTokens : IBoxAccessTokenProvider
    {
        public Task<string> GetAccessTokenAsync(CancellationToken cancellationToken = default) => Task.FromResult(FakeBox.Token);
    }
}
