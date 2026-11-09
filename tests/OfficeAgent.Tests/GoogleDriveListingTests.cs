using System.Net;
using OfficeAgent.Core.DocumentProviders;
using FakeDrive = OfficeAgent.Tests.GoogleDriveProviderTests.FakeDrive;

namespace OfficeAgent.Tests;

/// <summary>
/// Listing a Google Drive connection over the fake Drive API: <c>files.list</c> paging, folder
/// paths resolved by name, filters, templates, and the bounds that keep a listing inside the
/// configured folder.
/// </summary>
public class GoogleDriveListingTests
{
    private static int ListRequests(FakeDrive drive) =>
        drive.Requests.Count(r => Uri.UnescapeDataString(r).Contains(" in parents and trashed = false&", StringComparison.Ordinal) &&
                                  !Uri.UnescapeDataString(r).Contains("name = ", StringComparison.Ordinal));

    [Fact]
    public async Task Listing_pages_across_drive_pages_and_resumes_inside_one()
    {
        using var drive = new FakeDrive();
        for (var i = 0; i < 205; i++) drive.Seed($"doc{i:000}.docx", new byte[] { 1 });
        drive.Seed("elsewhere.docx", new byte[] { 1 }, parent: "otherFolder01");
        var provider = drive.Provider();

        var first = await provider.ListAsync(new DocumentListQuery { MaximumResults = 150 });
        var second = await provider.ListAsync(new DocumentListQuery { MaximumResults = 150, ContinuationToken = first.ContinuationToken });

        Assert.Equal(150, first.Items.Count);
        Assert.NotNull(first.ContinuationToken);
        Assert.Equal(55, second.Items.Count);
        Assert.Null(second.ContinuationToken);
        Assert.Equal(Enumerable.Range(0, 205).Select(i => $"doc{i:000}.docx"),
            first.Items.Concat(second.Items).Select(item => item.Name));
        Assert.Equal(3, ListRequests(drive));
        Assert.DoesNotContain(FakeDrive.FolderId, first.ContinuationToken);
    }

    [Fact]
    public async Task Listing_filters_by_name_and_extension_and_skips_what_cannot_be_registered()
    {
        using var drive = new FakeDrive();
        drive.Seed("Acme NDA.docx", new byte[] { 1 });
        drive.Seed("acme deck.pptx", new byte[] { 1 }, mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation");
        drive.Seed("acme budget.xlsx", new byte[] { 1 });
        drive.Seed("acme macro.docm", new byte[] { 1 });
        drive.Seed("acme native.docx", Array.Empty<byte>(), mimeType: "application/vnd.google-apps.document", size: null);
        drive.Seed("acme shortcut.docx", Array.Empty<byte>(), mimeType: "application/vnd.google-apps.shortcut", size: null);
        drive.Seed("acme big.docx", new byte[4096]);
        drive.Trash(drive.Seed("acme trashed.docx", new byte[] { 1 }));
        drive.Seed("other.docx", new byte[] { 1 });
        var provider = drive.Provider(maximumBytes: 1024, allowedExtensions: new[] { ".docx", ".pptx", ".docm" });

        var all = await provider.ListAsync(new DocumentListQuery());
        var named = await provider.ListAsync(new DocumentListQuery { NameContains = "ACME" });
        var narrowed = await provider.ListAsync(new DocumentListQuery { Extensions = new[] { ".pptx", "xlsx" } });

        Assert.Equal(new[] { "Acme NDA.docx", "acme deck.pptx", "other.docx" }, all.Items.Select(item => item.Name));
        Assert.Equal(new[] { "Acme NDA.docx", "acme deck.pptx" }, named.Items.Select(item => item.Name));
        Assert.Equal("acme deck.pptx", Assert.Single(narrowed.Items).Name);
        var entry = all.Items[0];
        Assert.Equal(1, entry.Length);
        Assert.Equal(new DateTimeOffset(2026, 11, 1, 9, 0, 0, TimeSpan.Zero), entry.LastModified);
        Assert.Equal("application/vnd.openxmlformats-officedocument.wordprocessingml.document", entry.ContentType);
    }

    [Fact]
    public async Task Folder_paths_resolve_by_name_and_recursion_names_each_folder()
    {
        using var drive = new FakeDrive();
        drive.Seed("top.docx", new byte[] { 1 });
        var clients = drive.SeedFolder("Clients");
        var acme = drive.SeedFolder("O'Brien & Acme", clients);
        drive.Seed("mid.docx", new byte[] { 1 }, parent: clients);
        drive.Seed("deep.docx", new byte[] { 1 }, parent: acme);
        var provider = drive.Provider();

        var flat = await provider.ListAsync(new DocumentListQuery());
        var recursive = await provider.ListAsync(new DocumentListQuery { Recursive = true });
        var nested = await provider.ListAsync(new DocumentListQuery { Folder = "clients/O'Brien & Acme" });
        var missing = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.ListAsync(new DocumentListQuery { Folder = "Clients/Missing" }));

        Assert.Equal("top.docx", Assert.Single(flat.Items).Name);
        Assert.Equal(
            new[] { ("top.docx", (string?)null), ("mid.docx", "Clients"), ("deep.docx", "Clients/O'Brien & Acme") },
            recursive.Items.Select(item => (item.Name, item.Folder)));
        var deep = Assert.Single(nested.Items);
        Assert.Equal(("deep.docx", "clients/O'Brien & Acme"), (deep.Name, deep.Folder));
        Assert.Equal(ProviderErrorCode.NotFound, missing.Code);
    }

    [Fact]
    public async Task Every_listed_source_registers_and_opens()
    {
        using var drive = new FakeDrive();
        drive.Seed("contract.docx", DocxFactory.Contract());
        drive.Seed("nested.docx", DocxFactory.Contract(), parent: drive.SeedFolder("Sub"));
        var provider = drive.Provider();

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
        using var drive = new FakeDrive();
        var templates = drive.SeedFolder("Firm templates", parent: "sharedRoot");
        drive.Seed("letter.dotx", new byte[] { 1 }, parent: templates);
        drive.Seed("memo.docx", new byte[] { 1 }, parent: templates);
        drive.Seed("plan.dotx", new byte[] { 1 });
        drive.Seed("draft.docx", new byte[] { 1 });
        var provider = drive.Provider(templatesFolder: templates, allowedExtensions: new[] { ".docx", ".dotx" });
        var withoutTemplates = drive.Provider();

        var listed = await provider.ListAsync(new DocumentListQuery { Templates = true });
        var documents = await provider.ListAsync(new DocumentListQuery());
        var requests = drive.Requests.Count;
        var none = await withoutTemplates.ListAsync(new DocumentListQuery { Templates = true });

        Assert.Equal(templates, provider.TemplatesFolder);
        Assert.Equal(TemplateLocationKind.Connection, provider.TemplateLocation);
        Assert.Equal(new[] { ("letter.dotx", true), ("memo.docx", true) }, listed.Items.Select(item => (item.Name, item.IsTemplate)));
        Assert.Equal(new[] { ("plan.dotx", true), ("draft.docx", false) }, documents.Items.Select(item => (item.Name, item.IsTemplate)));
        Assert.Null(withoutTemplates.TemplatesFolder);
        Assert.Equal(TemplateLocationKind.None, withoutTemplates.TemplateLocation);
        Assert.Empty(none.Items);
        Assert.Equal(requests, drive.Requests.Count);
    }

    [Fact]
    public async Task A_shared_drive_folder_is_listed_with_its_drive()
    {
        using var drive = new FakeDrive { SharedDriveId = "0AsharedDrive01" };
        drive.Seed("contract.docx", new byte[] { 1 });

        var listed = await drive.Provider().ListAsync(new DocumentListQuery { Folder = "/" });

        Assert.Equal("contract.docx", Assert.Single(listed.Items).Name);
    }

    [Fact]
    public async Task Listing_stays_inside_the_folder_and_its_tokens_cannot_be_redirected()
    {
        using var drive = new FakeDrive();
        for (var i = 0; i < 3; i++) drive.Seed($"doc{i}.docx", new byte[] { 1 });
        drive.SeedFolder("Sub");
        var provider = drive.Provider();
        var token = (await provider.ListAsync(new DocumentListQuery { MaximumResults = 1 })).ContinuationToken!;
        var requests = drive.Requests.Count;

        var escape = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.ListAsync(new DocumentListQuery { Folder = "../elsewhere" }));
        var reused = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.ListAsync(new DocumentListQuery { Folder = "Sub", MaximumResults = 1, ContinuationToken = token }));
        var edited = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.ListAsync(new DocumentListQuery { MaximumResults = 1, ContinuationToken = "x" + token }));

        Assert.Equal(ProviderErrorCode.AccessDenied, escape.Code);
        Assert.Equal(ProviderErrorCode.InvalidArgument, reused.Code);
        Assert.Equal(ProviderErrorCode.InvalidArgument, edited.Code);
        Assert.Equal(requests, drive.Requests.Count);
    }

    [Fact]
    public async Task Listing_failures_map_to_provider_errors()
    {
        using var drive = new FakeDrive();
        var file = drive.Seed("contract.docx", new byte[] { 1 });

        drive.ListStatus = HttpStatusCode.Forbidden;
        var denied = await Assert.ThrowsAsync<DocumentProviderException>(() => drive.Provider().ListAsync(new DocumentListQuery()));
        drive.ListStatus = null;
        var notFolder = await Assert.ThrowsAsync<DocumentProviderException>(() => drive.Provider(templatesFolder: file)
            .ListAsync(new DocumentListQuery { Templates = true }));
        var unconfigured = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            drive.Provider(configureCreation: false).ListAsync(new DocumentListQuery()));
        drive.RemoveFolder();
        var missing = await Assert.ThrowsAsync<DocumentProviderException>(() => drive.Provider().ListAsync(new DocumentListQuery()));

        Assert.Equal(ProviderErrorCode.AccessDenied, denied.Code);
        Assert.Equal(ProviderErrorCode.ConfigurationError, notFolder.Code);
        Assert.Equal(ProviderErrorCode.ConfigurationError, unconfigured.Code);
        Assert.Equal(ProviderErrorCode.NotFound, missing.Code);
        Assert.All(new[] { denied, notFolder, unconfigured, missing }, error => Assert.DoesNotContain(FakeDrive.FolderId, error.Message));
    }

    [Fact]
    public void A_templates_folder_must_be_a_drive_id_and_a_location_needs_a_folder()
    {
        using var drive = new FakeDrive();

        Assert.Throws<ArgumentException>(() => drive.Provider(templatesFolder: "Templates/Firm"));
        var organization = new OfficeAgent.GoogleDrive.GoogleDriveDocumentProvider(
            new OfficeAgent.GoogleDrive.GoogleDriveDocumentProviderOptions
            {
                ConnectionId = "drive", TemplatesFolder = "templatesFolder01", TemplateLocation = TemplateLocationKind.Organization
            }, drive.Http, new NoTokens());
        Assert.Equal(TemplateLocationKind.Organization, organization.TemplateLocation);
        Assert.Throws<ArgumentException>(() => new OfficeAgent.GoogleDrive.GoogleDriveDocumentProvider(
            new OfficeAgent.GoogleDrive.GoogleDriveDocumentProviderOptions
            {
                ConnectionId = "drive", TemplateLocation = TemplateLocationKind.Workgroup
            }, drive.Http, new NoTokens()));
    }

    private sealed class NoTokens : OfficeAgent.GoogleDrive.IGoogleDriveAccessTokenProvider
    {
        public Task<string> GetAccessTokenAsync(CancellationToken cancellationToken = default) => Task.FromResult("unused");
    }
}
