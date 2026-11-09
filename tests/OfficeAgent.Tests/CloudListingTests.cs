using OfficeAgent.Core.DocumentProviders;

namespace OfficeAgent.Tests;

/// <summary>
/// The bounds every cloud listing shares, run over an in-memory folder tree: the page size clamp,
/// the request budget, the recursion depth, and the continuation token's binding.
/// </summary>
public class CloudListingTests
{
    /// <summary>A folder tree: each folder id maps to its children, in service order.</summary>
    private sealed class Tree
    {
        public Dictionary<string, List<CloudListingChild>> Folders { get; } = new();
        public int Calls { get; private set; }
        public int PageSize { get; init; } = 2;

        public Task<CloudListingPage> Fetch(string folderId, string? token, CancellationToken cancellationToken)
        {
            Calls++;
            var all = Folders.TryGetValue(folderId, out var children) ? children : new List<CloudListingChild>();
            var start = token is null ? 0 : int.Parse(token);
            return Task.FromResult(new CloudListingPage
            {
                Items = all.Skip(start).Take(PageSize).ToList(),
                NextToken = start + PageSize < all.Count ? (start + PageSize).ToString() : null
            });
        }
    }

    private static CloudListingChild Doc(string name) => new() { Id = name, Name = name };

    private static CloudListingChild Folder(string id) => new() { Id = id, Name = id, IsFolder = true };

    private static CloudListingRequest Request(DocumentListQuery query, params string[] allowed) =>
        CloudListingRequest.Create(query, allowed.Length == 0 ? new[] { ".docx" } : allowed, out _)!;

    [Theory]
    [InlineData(0, 1)]
    [InlineData(-5, 1)]
    [InlineData(100, 100)]
    [InlineData(10_000, 500)]
    public void The_page_size_is_clamped(int requested, int expected) =>
        Assert.Equal(expected, Request(new DocumentListQuery { MaximumResults = requested }).MaximumResults);

    [Fact]
    public void Macro_enabled_extensions_are_never_listed_and_a_query_only_narrows()
    {
        var all = Request(new DocumentListQuery(), ".docx", ".DOCM", "xlsm", ".dotx");
        var narrowed = Request(new DocumentListQuery { Extensions = new[] { "dotx", ".pptx" } }, ".docx", ".dotx");

        Assert.Equal(new[] { ".docx", ".dotx" }, all.Extensions.OrderBy(e => e));
        Assert.Equal(new[] { ".dotx" }, narrowed.Extensions);
    }

    [Fact]
    public async Task A_page_spends_at_most_its_request_budget_and_carries_on_from_there()
    {
        // Thirty pages of files no query wants, then one it does.
        var tree = new Tree { PageSize = 1 };
        tree.Folders["root"] = Enumerable.Range(0, 30).Select(i => Doc($"skip{i}.xlsx")).Append(Doc("wanted.docx")).ToList();
        var request = Request(new DocumentListQuery());

        var first = await CloudListing.WalkAsync(request, CloudListing.Start("root", request), tree.Fetch,
            CloudListing.MaximumServiceCalls, "c", CancellationToken.None);
        Assert.Empty(first.Items);
        Assert.NotNull(first.ContinuationToken);
        Assert.Equal(CloudListing.MaximumServiceCalls, tree.Calls);

        var found = new List<DocumentListEntry>();
        var token = first.ContinuationToken;
        for (var page = 0; token is not null && page < 10; page++)
        {
            Assert.True(CloudListingCursor.TryUnprotect(token, "c", request.Fingerprint, out var cursor));
            var next = await CloudListing.WalkAsync(request, cursor, tree.Fetch, CloudListing.MaximumServiceCalls, "c", CancellationToken.None);
            found.AddRange(next.Items);
            token = next.ContinuationToken;
        }

        Assert.Null(token);
        Assert.Equal("wanted.docx", Assert.Single(found).Name);
        Assert.Equal(31, tree.Calls);
    }

    [Fact]
    public async Task Recursion_stops_at_the_documented_depth()
    {
        var tree = new Tree { PageSize = 10 };
        for (var depth = 0; depth <= CloudListing.MaximumDepth + 2; depth++)
            tree.Folders[$"f{depth}"] = new List<CloudListingChild> { Doc($"at{depth}.docx"), Folder($"f{depth + 1}") };
        var request = Request(new DocumentListQuery { Recursive = true });

        var page = await CloudListing.WalkAsync(request, CloudListing.Start("f0", request), tree.Fetch, 100, "c", CancellationToken.None);

        Assert.Equal(Enumerable.Range(0, CloudListing.MaximumDepth + 1).Select(depth => $"at{depth}.docx"), page.Items.Select(item => item.Name));
        Assert.Equal("f1/f2", page.Items[2].Folder);
        Assert.Null(page.ContinuationToken);
    }

    [Fact]
    public void A_token_is_bound_to_its_connection_and_query()
    {
        var request = Request(new DocumentListQuery { Folder = "a/b" });
        var token = CloudListing.Start("secret-folder-id", request).Protect("legal", request.Fingerprint);
        var other = Request(new DocumentListQuery { Folder = "a" });

        Assert.DoesNotContain("secret-folder-id", token);
        Assert.True(CloudListingCursor.TryUnprotect(token, "legal", request.Fingerprint, out var cursor));
        Assert.Equal("secret-folder-id", cursor.Current!.Id);
        Assert.Equal("a/b", cursor.Current.Path);
        Assert.False(CloudListingCursor.TryUnprotect(token, "other", request.Fingerprint, out _));
        Assert.False(CloudListingCursor.TryUnprotect(token, "legal", other.Fingerprint, out _));
        Assert.False(CloudListingCursor.TryUnprotect(token.Replace('.', '!'), "legal", request.Fingerprint, out _));
    }

    [Fact]
    public void A_template_location_needs_a_folder_and_defaults_to_the_connection()
    {
        Assert.Equal(TemplateLocationKind.None, CloudListing.TemplateLocationFor(false, null, "o"));
        Assert.Equal(TemplateLocationKind.Connection, CloudListing.TemplateLocationFor(true, null, "o"));
        Assert.Equal(TemplateLocationKind.Organization, CloudListing.TemplateLocationFor(true, TemplateLocationKind.Organization, "o"));
        Assert.Throws<ArgumentException>(() => CloudListing.TemplateLocationFor(false, TemplateLocationKind.Personal, "o"));
        Assert.Throws<ArgumentException>(() => CloudListing.TemplateLocationFor(true, TemplateLocationKind.None, "o"));
        Assert.Throws<ArgumentException>(() => CloudListing.TemplateLocationFor(true, (TemplateLocationKind)42, "o"));
    }
}
