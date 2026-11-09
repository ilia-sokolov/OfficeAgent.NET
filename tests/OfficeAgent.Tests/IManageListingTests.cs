using System.Net;
using OfficeAgent.Core.DocumentProviders;
using FakeIManage = OfficeAgent.Tests.IManageProviderTests.FakeIManage;

namespace OfficeAgent.Tests;

/// <summary>
/// Listing an iManage connection over the fake Work API: folder children with
/// <c>limit</c>/<c>offset</c> paging, folder paths resolved by name, filters, templates, and the
/// bounds that keep a listing inside the configured folder.
/// </summary>
public class IManageListingTests
{
    private static int ChildRequests(FakeIManage server) => server.Requests.Count(r => r.Contains("/children?", StringComparison.Ordinal));

    [Fact]
    public async Task Listing_pages_across_children_pages_and_resumes_inside_one()
    {
        using var server = new FakeIManage();
        for (var i = 0; i < 205; i++) server.Seed($"doc{i:000}.docx", new byte[] { 1 }, inFolder: true);
        server.Seed("unfiled.docx", new byte[] { 1 });
        var provider = server.Provider();

        var first = await provider.ListAsync(new DocumentListQuery { MaximumResults = 150 });
        var second = await provider.ListAsync(new DocumentListQuery { MaximumResults = 150, ContinuationToken = first.ContinuationToken });

        Assert.Equal(150, first.Items.Count);
        Assert.NotNull(first.ContinuationToken);
        Assert.Equal(55, second.Items.Count);
        Assert.Null(second.ContinuationToken);
        Assert.Equal(Enumerable.Range(0, 205).Select(i => $"doc{i:000}.docx"),
            first.Items.Concat(second.Items).Select(item => item.Name));
        Assert.Equal(3, ChildRequests(server));
        Assert.Contains(server.Requests, r => r.Contains("/children?limit=200&offset=200", StringComparison.Ordinal));
        Assert.DoesNotContain("ACTIVE", first.ContinuationToken);
    }

    [Fact]
    public async Task Listing_filters_by_name_and_extension_and_skips_what_cannot_be_registered()
    {
        using var server = new FakeIManage();
        server.Seed("Acme NDA.docx", new byte[] { 1 }, inFolder: true);
        server.Seed("acme deck.pptx", new byte[] { 1 }, inFolder: true);
        server.Seed("acme budget.xlsx", new byte[] { 1 }, inFolder: true);
        server.Seed("acme macro.docm", new byte[] { 1 }, inFolder: true);
        server.Seed("acme big.docx", new byte[4096], inFolder: true);
        server.Seed("other.docx", new byte[] { 1 }, inFolder: true);
        var provider = server.Provider(maximumBytes: 1024, allowedExtensions: new[] { ".docx", ".pptx", ".docm" });

        var all = await provider.ListAsync(new DocumentListQuery());
        var named = await provider.ListAsync(new DocumentListQuery { NameContains = "ACME" });
        var narrowed = await provider.ListAsync(new DocumentListQuery { Extensions = new[] { ".pptx", ".xlsx" } });

        Assert.Equal(new[] { "Acme NDA.docx", "acme deck.pptx", "other.docx" }, all.Items.Select(item => item.Name));
        Assert.Equal(new[] { "Acme NDA.docx", "acme deck.pptx" }, named.Items.Select(item => item.Name));
        Assert.Equal("acme deck.pptx", Assert.Single(narrowed.Items).Name);
        var entry = all.Items[0];
        Assert.Matches(@"^ACTIVE!\d+\.1$", entry.Source);
        Assert.Equal(1, entry.Length);
        Assert.NotNull(entry.LastModified);
    }

    [Fact]
    public async Task Folder_paths_resolve_by_name_and_recursion_names_each_folder()
    {
        using var server = new FakeIManage();
        server.Seed("top.docx", new byte[] { 1 }, inFolder: true);
        var clients = server.AddFolder("Clients");
        var acme = server.AddFolder("Acme", clients);
        server.SeedIn(clients, "mid.docx", new byte[] { 1 });
        server.SeedIn(acme, "deep.docx", new byte[] { 1 });
        var provider = server.Provider();

        var flat = await provider.ListAsync(new DocumentListQuery());
        var recursive = await provider.ListAsync(new DocumentListQuery { Recursive = true });
        var nested = await provider.ListAsync(new DocumentListQuery { Folder = "clients/ACME" });
        var missing = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.ListAsync(new DocumentListQuery { Folder = "Clients/Missing" }));

        Assert.Equal("top.docx", Assert.Single(flat.Items).Name);
        Assert.Equal(
            new[] { ("top.docx", (string?)null), ("mid.docx", "Clients"), ("deep.docx", "Clients/Acme") },
            recursive.Items.Select(item => (item.Name, item.Folder)));
        Assert.Equal("deep.docx", Assert.Single(nested.Items).Name);
        Assert.Equal(ProviderErrorCode.NotFound, missing.Code);
    }

    [Fact]
    public async Task Every_listed_source_registers_and_opens()
    {
        using var server = new FakeIManage();
        server.Seed("contract.docx", DocxFactory.Contract(), inFolder: true);
        server.SeedIn(server.AddFolder("Sub"), "nested.docx", DocxFactory.Contract());
        var provider = server.Provider();

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
        using var server = new FakeIManage();
        var precedents = server.AddFolder("Precedents", parent: "ACTIVE!1");
        server.SeedIn(precedents, "letter.dotx", new byte[] { 1 });
        server.SeedIn(precedents, "memo.docx", new byte[] { 1 });
        server.Seed("plan.dotx", new byte[] { 1 }, inFolder: true);
        server.Seed("draft.docx", new byte[] { 1 }, inFolder: true);
        var provider = server.Provider(templatesFolder: precedents, allowedExtensions: new[] { ".docx", ".dotx" });
        var withoutTemplates = server.Provider();

        var listed = await provider.ListAsync(new DocumentListQuery { Templates = true });
        var documents = await provider.ListAsync(new DocumentListQuery());
        var requests = server.Requests.Count;
        var none = await withoutTemplates.ListAsync(new DocumentListQuery { Templates = true });

        Assert.Equal(precedents, provider.TemplatesFolder);
        Assert.Equal(TemplateLocationKind.Connection, provider.TemplateLocation);
        Assert.Equal(new[] { ("letter.dotx", true), ("memo.docx", true) }, listed.Items.Select(item => (item.Name, item.IsTemplate)));
        Assert.Equal(new[] { ("plan.dotx", true), ("draft.docx", false) }, documents.Items.Select(item => (item.Name, item.IsTemplate)));
        Assert.Equal(TemplateLocationKind.None, withoutTemplates.TemplateLocation);
        Assert.Empty(none.Items);
        Assert.Equal(requests, server.Requests.Count);
    }

    [Fact]
    public async Task Listing_stays_inside_the_folder_and_its_tokens_cannot_be_redirected()
    {
        using var server = new FakeIManage();
        for (var i = 0; i < 3; i++) server.Seed($"doc{i}.docx", new byte[] { 1 }, inFolder: true);
        var provider = server.Provider();
        var token = (await provider.ListAsync(new DocumentListQuery { MaximumResults = 1 })).ContinuationToken!;
        var requests = server.Requests.Count;

        var escape = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.ListAsync(new DocumentListQuery { Folder = ".." }));
        var reused = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.ListAsync(new DocumentListQuery { MaximumResults = 1, NameContains = "doc", ContinuationToken = token }));
        // A documents token cannot continue a templates listing, even over the same folder.
        var reusedOnTemplates = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            server.Provider(templatesFolder: "ACTIVE!773").ListAsync(
                new DocumentListQuery { MaximumResults = 1, Templates = true, ContinuationToken = token }));

        Assert.Equal(ProviderErrorCode.AccessDenied, escape.Code);
        Assert.Equal(ProviderErrorCode.InvalidArgument, reused.Code);
        Assert.Equal(ProviderErrorCode.InvalidArgument, reusedOnTemplates.Code);
        Assert.Equal(requests, server.Requests.Count);
    }

    [Fact]
    public async Task Listing_failures_map_to_provider_errors()
    {
        using (var server = new FakeIManage())
        {
            server.ListStatus = HttpStatusCode.Forbidden;
            var denied = await Assert.ThrowsAsync<DocumentProviderException>(() => server.Provider().ListAsync(new DocumentListQuery()));
            server.ListStatus = null;
            var unconfigured = await Assert.ThrowsAsync<DocumentProviderException>(() =>
                server.Provider(configureCreation: false).ListAsync(new DocumentListQuery()));

            Assert.Equal(ProviderErrorCode.AccessDenied, denied.Code);
            Assert.Equal(ProviderErrorCode.ConfigurationError, unconfigured.Code);
            Assert.Throws<ArgumentException>(() => server.Provider(templatesFolder: "ACTIVE!773.1"));
        }

        using var missingFolder = new FakeIManage { MissingFolder = true };
        var missing = await Assert.ThrowsAsync<DocumentProviderException>(() => missingFolder.Provider().ListAsync(new DocumentListQuery()));
        Assert.Equal(ProviderErrorCode.NotFound, missing.Code);
        Assert.DoesNotContain("ACTIVE", missing.Message);
    }
}
