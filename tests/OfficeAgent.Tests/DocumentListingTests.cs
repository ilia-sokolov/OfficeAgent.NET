using OfficeAgent.Abstractions;
using OfficeAgent.Core.DocumentProviders;

namespace OfficeAgent.Tests;

/// <summary>
/// Listing on the filesystem and session-memory providers: what is listed and what is not, the
/// page and work bounds, continuation tokens, the extension allow-list, templates folders, the
/// root boundary (links and junctions included) and read-only connections.
/// </summary>
public class DocumentListingTests
{
    private static readonly byte[] Bytes = { 1, 2, 3, 4 };

    [Fact]
    public async Task Lists_documents_in_name_order_with_sources_registration_accepts()
    {
        using var workspace = new Workspace();
        workspace.Write("b.docx");
        workspace.Write("A.docx");
        workspace.Write("Contracts/c.docx");
        var provider = workspace.Provider();

        var page = await provider.ListAsync(new DocumentListQuery { Recursive = true });

        Assert.Equal(new[] { "A.docx", "b.docx", "Contracts/c.docx" }, page.Items.Select(i => i.Source));
        Assert.Equal(new string?[] { null, null, "Contracts" }, page.Items.Select(i => i.Folder));
        Assert.Null(page.ContinuationToken);
        Assert.All(page.Items, item =>
        {
            Assert.False(Path.IsPathRooted(item.Source));
            Assert.DoesNotContain(workspace.Root, item.Source);
            Assert.Equal("application/vnd.openxmlformats-officedocument.wordprocessingml.document", item.ContentType);
            Assert.Equal(Bytes.Length, item.Length);
            Assert.NotNull(item.LastModified);
            Assert.False(item.IsTemplate);
        });

        foreach (var item in page.Items)
        {
            var reference = await provider.RegisterAsync(item.Source);
            Assert.Equal(item.Name, reference.Name);
        }
    }

    [Fact]
    public async Task Without_recursion_only_the_folder_itself_is_listed()
    {
        using var workspace = new Workspace();
        workspace.Write("top.docx");
        workspace.Write("Contracts/inner.docx");
        workspace.Write("Contracts/2026/deep.docx");
        var provider = workspace.Provider();

        Assert.Equal(new[] { "top.docx" },
            (await provider.ListAsync(new DocumentListQuery())).Items.Select(i => i.Source));
        Assert.Equal(new[] { "Contracts/inner.docx" },
            (await provider.ListAsync(new DocumentListQuery { Folder = "Contracts" })).Items.Select(i => i.Source));
        Assert.Equal(new[] { "Contracts/2026/deep.docx", "Contracts/inner.docx" },
            (await provider.ListAsync(new DocumentListQuery { Folder = "Contracts", Recursive = true })).Items.Select(i => i.Source));
    }

    [Fact]
    public async Task Lock_files_hidden_entries_and_the_index_are_never_listed()
    {
        using var workspace = new Workspace();
        workspace.Write("kept.docx");
        workspace.Write("~$kept.docx");
        workspace.Write(".staging.docx");
        workspace.Write("notes.txt");
        var hidden = workspace.Write("hidden.docx");
        File.SetAttributes(hidden, File.GetAttributes(hidden) | FileAttributes.Hidden);
        // Linux has no hidden attribute, only dot-prefixed names, so there the file stays visible.
        var attributeKept = (File.GetAttributes(hidden) & FileAttributes.Hidden) != 0;
        var provider = workspace.Provider();
        await provider.RegisterAsync("kept.docx");

        var page = await provider.ListAsync(new DocumentListQuery { Recursive = true });

        Assert.Equal(attributeKept ? new[] { "kept.docx" } : new[] { "hidden.docx", "kept.docx" },
            page.Items.Select(i => i.Source));
    }

    [Fact]
    public async Task Extensions_narrow_the_allow_list_and_never_widen_it()
    {
        using var workspace = new Workspace();
        workspace.Write("a.docx");
        workspace.Write("b.pptx");
        workspace.Write("c.xlsx");
        var provider = workspace.Provider(o => o.AllowedExtensions = new[] { ".docx", ".pptx" });

        Assert.Equal(new[] { "a.docx", "b.pptx" },
            (await provider.ListAsync(new DocumentListQuery())).Items.Select(i => i.Name));
        Assert.Equal(new[] { "b.pptx" },
            (await provider.ListAsync(new DocumentListQuery { Extensions = new[] { "pptx" } })).Items.Select(i => i.Name));
        Assert.Empty((await provider.ListAsync(new DocumentListQuery { Extensions = new[] { ".xlsx" } })).Items);
    }

    [Fact]
    public async Task Macro_enabled_files_are_never_listed_even_when_allowed()
    {
        using var workspace = new Workspace();
        workspace.Write("plain.docx");
        workspace.Write("macro.docm");
        workspace.Write("macro.dotm");
        var provider = workspace.Provider(o => o.AllowedExtensions = new[] { ".docx", ".docm", ".dotm" });

        var page = await provider.ListAsync(new DocumentListQuery { Extensions = new[] { ".docm", ".docx", ".dotm" } });

        Assert.Equal(new[] { "plain.docx" }, page.Items.Select(i => i.Name));
    }

    [Fact]
    public async Task Name_filter_ignores_case_and_files_over_the_size_limit_are_skipped()
    {
        using var workspace = new Workspace();
        workspace.Write("Engagement Letter.docx");
        workspace.Write("invoice.docx");
        workspace.Write("letter-large.docx", new byte[64]);
        var provider = workspace.Provider(o => o.MaximumBytes = 16);

        var page = await provider.ListAsync(new DocumentListQuery { NameContains = "LETTER" });

        Assert.Equal(new[] { "Engagement Letter.docx" }, page.Items.Select(i => i.Name));
    }

    [Fact]
    public async Task Pages_follow_continuation_tokens_to_the_end_without_repeats()
    {
        using var workspace = new Workspace();
        foreach (var name in new[] { "a", "b", "c" }) workspace.Write($"{name}.docx");
        foreach (var name in new[] { "d", "e" }) workspace.Write($"Sub/{name}.docx");
        workspace.Write("Sub/Deeper/f.docx");
        workspace.Write("z.docx");
        var provider = workspace.Provider();

        var seen = new List<string>();
        string? token = null;
        var pages = 0;
        do
        {
            var page = await provider.ListAsync(new DocumentListQuery { Recursive = true, MaximumResults = 2, ContinuationToken = token });
            Assert.True(page.Items.Count <= 2);
            seen.AddRange(page.Items.Select(i => i.Source));
            token = page.ContinuationToken;
            pages++;
        } while (token is not null && pages < 10);

        Assert.Equal(new[] { "a.docx", "b.docx", "c.docx", "Sub/d.docx", "Sub/Deeper/f.docx", "Sub/e.docx", "z.docx" }, seen);
        Assert.Equal(4, pages);
    }

    [Fact]
    public async Task Page_size_is_clamped()
    {
        using var workspace = new Workspace();
        workspace.Write("a.docx");
        workspace.Write("b.docx");
        var provider = workspace.Provider();

        var page = await provider.ListAsync(new DocumentListQuery { MaximumResults = 0 });

        Assert.Single(page.Items);
        Assert.NotNull(page.ContinuationToken);
    }

    [Fact]
    public async Task A_token_is_refused_for_another_query_or_when_forged()
    {
        using var workspace = new Workspace();
        workspace.Write("a.docx");
        workspace.Write("b.docx");
        var provider = workspace.Provider();
        var first = await provider.ListAsync(new DocumentListQuery { MaximumResults = 1 });

        var otherQuery = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.ListAsync(new DocumentListQuery { MaximumResults = 1, NameContains = "a", ContinuationToken = first.ContinuationToken }));
        Assert.Equal(ProviderErrorCode.InvalidArgument, otherQuery.Code);

        var forged = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.ListAsync(new DocumentListQuery { ContinuationToken = "not-a-token" }));
        Assert.Equal(ProviderErrorCode.InvalidArgument, forged.Code);
    }

    [Fact]
    public async Task A_token_naming_a_path_outside_the_folder_reaches_nothing_outside_it()
    {
        using var workspace = new Workspace();
        using var outsider = new Workspace();
        outsider.Write("secret.docx");
        workspace.Write("a.docx");
        var provider = workspace.Provider();
        var query = new DocumentListQuery { Recursive = true };
        var fingerprint = DocumentListingRules.Fingerprint(query, FileSystemDocumentProvider.ProviderName);
        var forged = DocumentListingRules.EncodeToken($"fs1\n{fingerprint}\n../{Path.GetFileName(outsider.Root)}/secret.docx");

        var page = await provider.ListAsync(new DocumentListQuery { Recursive = true, ContinuationToken = forged });

        Assert.DoesNotContain(page.Items, item => item.Name == "secret.docx");
    }

    [Fact]
    public async Task Recursion_stops_at_the_documented_depth()
    {
        using var workspace = new Workspace();
        var path = "";
        for (var depth = 0; depth <= FileSystemDocumentProvider.MaximumListingDepth + 1; depth++)
        {
            workspace.Write($"{path}level{depth}.docx");
            path += $"d{depth}/";
        }
        var provider = workspace.Provider();

        var page = await provider.ListAsync(new DocumentListQuery { Recursive = true });

        Assert.Equal(FileSystemDocumentProvider.MaximumListingDepth + 1, page.Items.Count);
        Assert.DoesNotContain(page.Items, i => i.Name == $"level{FileSystemDocumentProvider.MaximumListingDepth + 1}.docx");
    }

    [Theory]
    [InlineData("..")]
    [InlineData("../elsewhere")]
    [InlineData("Contracts/../../elsewhere")]
    public async Task A_folder_outside_the_root_is_refused(string folder)
    {
        using var workspace = new Workspace();
        workspace.Write("Contracts/a.docx");
        var provider = workspace.Provider();

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.ListAsync(new DocumentListQuery { Folder = folder }));

        Assert.Equal(ProviderErrorCode.AccessDenied, error.Code);
    }

    [Fact]
    public async Task An_absolute_folder_outside_the_root_and_the_index_directory_are_refused()
    {
        using var workspace = new Workspace();
        using var outsider = new Workspace();
        outsider.Write("secret.docx");
        var provider = workspace.Provider();

        Assert.Equal(ProviderErrorCode.AccessDenied, (await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.ListAsync(new DocumentListQuery { Folder = outsider.Root }))).Code);
        Assert.Equal(ProviderErrorCode.AccessDenied, (await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.ListAsync(new DocumentListQuery { Folder = ".officeagent" }))).Code);
    }

    [Fact]
    public async Task A_linked_folder_is_neither_followed_nor_listable()
    {
        using var workspace = new Workspace();
        using var outsider = new Workspace();
        outsider.Write("secret.docx");
        workspace.Write("inside.docx");
        var link = Path.Combine(workspace.Root, "linked");
        if (!TryCreateDirectoryLink(link, outsider.Root)) return;

        try
        {
            var provider = workspace.Provider();
            var page = await provider.ListAsync(new DocumentListQuery { Recursive = true });
            Assert.Equal(new[] { "inside.docx" }, page.Items.Select(i => i.Source));

            var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
                provider.ListAsync(new DocumentListQuery { Folder = "linked" }));
            Assert.Equal(ProviderErrorCode.AccessDenied, error.Code);
        }
        finally
        {
            if (Directory.Exists(link)) Directory.Delete(link);
        }
    }

    [Fact]
    public async Task A_templates_folder_that_is_a_link_is_refused()
    {
        using var workspace = new Workspace();
        using var outsider = new Workspace();
        outsider.Write("secret.dotx");
        var link = Path.Combine(workspace.Root, "Templates");
        if (!TryCreateDirectoryLink(link, outsider.Root)) return;

        try
        {
            var provider = workspace.Provider(o =>
            {
                o.TemplatesFolder = "Templates";
                o.AllowedExtensions = new[] { ".docx", ".dotx" };
            });
            var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
                provider.ListAsync(new DocumentListQuery { Templates = true }));
            Assert.Equal(ProviderErrorCode.AccessDenied, error.Code);
        }
        finally
        {
            if (Directory.Exists(link)) Directory.Delete(link);
        }
    }

    [Fact]
    public async Task The_templates_folder_lists_templates_and_reports_them_as_such()
    {
        using var workspace = new Workspace();
        workspace.Write("work.docx");
        workspace.Write("loose.dotx");
        workspace.Write("Templates/Letter.dotx");
        workspace.Write("Templates/Board/Minutes.docx");
        var provider = workspace.Provider(o =>
        {
            o.TemplatesFolder = "Templates/";
            o.AllowedExtensions = new[] { ".docx", ".dotx" };
        });

        Assert.Equal("Templates", provider.TemplatesFolder);
        Assert.Equal(TemplateLocationKind.Connection, provider.TemplateLocation);

        var templates = await provider.ListAsync(new DocumentListQuery { Templates = true, Recursive = true });
        Assert.Equal(new[] { "Templates/Board/Minutes.docx", "Templates/Letter.dotx" }, templates.Items.Select(i => i.Source));
        Assert.Equal(new string?[] { "Board", null }, templates.Items.Select(i => i.Folder));
        Assert.All(templates.Items, item => Assert.True(item.IsTemplate));

        var everything = await provider.ListAsync(new DocumentListQuery { Recursive = true });
        Assert.True(everything.Items.Single(i => i.Name == "loose.dotx").IsTemplate);
        Assert.False(everything.Items.Single(i => i.Name == "work.docx").IsTemplate);
        Assert.True(everything.Items.Single(i => i.Name == "Minutes.docx").IsTemplate);

        var registered = await provider.RegisterAsync(templates.Items[1].Source);
        Assert.Equal("Letter.dotx", registered.Name);
    }

    [Fact]
    public async Task Template_extensions_are_listed_only_when_the_connection_allows_them()
    {
        using var workspace = new Workspace();
        workspace.Write("Templates/Letter.dotx");
        workspace.Write("Templates/Memo.docx");
        var provider = workspace.Provider(o => o.TemplatesFolder = "Templates");

        var page = await provider.ListAsync(new DocumentListQuery { Templates = true });

        Assert.Equal(new[] { "Memo.docx" }, page.Items.Select(i => i.Name));
    }

    [Fact]
    public async Task Without_a_templates_folder_a_templates_listing_is_empty_and_a_missing_one_lists_nothing()
    {
        using var workspace = new Workspace();
        workspace.Write("a.docx");

        Assert.Empty((await workspace.Provider().ListAsync(new DocumentListQuery { Templates = true })).Items);
        Assert.Equal(TemplateLocationKind.None, workspace.Provider().TemplateLocation);
        Assert.Empty((await workspace.Provider(o => o.TemplatesFolder = "NotYet")
            .ListAsync(new DocumentListQuery { Templates = true })).Items);
    }

    [Fact]
    public async Task A_template_location_without_a_folder_is_templates_as_a_whole()
    {
        using var workspace = new Workspace();
        workspace.Write("Letter.dotx");
        workspace.Write("Report.docx");
        var provider = workspace.Provider(o =>
        {
            o.TemplateLocation = TemplateLocationKind.Personal;
            o.AllowedExtensions = new[] { ".docx", ".dotx" };
        });

        var page = await provider.ListAsync(new DocumentListQuery { Templates = true });

        Assert.Null(provider.TemplatesFolder);
        Assert.Equal(new[] { "Letter.dotx", "Report.docx" }, page.Items.Select(i => i.Source));
        Assert.All(page.Items, item => Assert.True(item.IsTemplate));
    }

    [Theory]
    [InlineData("../Templates")]
    [InlineData(".")]
    [InlineData(".officeagent")]
    public void A_templates_folder_must_lie_below_the_root(string folder)
    {
        using var workspace = new Workspace();

        Assert.Throws<ArgumentException>(() => workspace.Provider(o => o.TemplatesFolder = folder));
    }

    [Fact]
    public void An_absolute_templates_folder_is_refused()
    {
        using var workspace = new Workspace();

        Assert.Throws<ArgumentException>(() =>
            workspace.Provider(o => o.TemplatesFolder = Path.Combine(workspace.Root, "Templates")));
    }

    [Fact]
    public async Task A_read_only_connection_lists_and_registers_but_refuses_every_write()
    {
        using var workspace = new Workspace();
        workspace.Write("Letter.dotx", DocxFactory.Contract());
        var provider = workspace.Provider(o =>
        {
            o.ReadOnly = true;
            o.AllowedExtensions = new[] { ".docx", ".dotx" };
        });

        Assert.True(provider.ReadOnly);
        var listed = Assert.Single((await provider.ListAsync(new DocumentListQuery())).Items);
        var reference = await provider.RegisterAsync(listed.Source);
        using (var content = await provider.OpenReadAsync(reference))
            Assert.True(content.Stream.Length > 0);

        // Nothing is written under a read-only root, not even the registration index.
        Assert.False(Directory.Exists(Path.Combine(workspace.Root, ".officeagent")));

        using var bytes = new MemoryStream(DocxFactory.Contract());
        var save = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.SaveAsync(reference, bytes, new SaveDocumentOptions { Mode = SaveMode.Replace }));
        var newVersion = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.SaveAsync(reference, bytes, new SaveDocumentOptions { Mode = SaveMode.NewVersion }));
        var create = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.CreateAsync("new.docx", bytes));
        var remove = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.RemoveAsync(reference));

        Assert.All(new[] { save, newVersion, create, remove }, error => Assert.Equal(ProviderErrorCode.AccessDenied, error.Code));
        Assert.Equal(new[] { "Letter.dotx" }, Directory.GetFiles(workspace.Root).Select(Path.GetFileName));
    }

    [Fact]
    public async Task A_read_only_connection_over_a_missing_folder_creates_nothing()
    {
        var root = Path.Combine(Path.GetTempPath(), $"officeagent-missing-{Guid.NewGuid():N}");
        var provider = new FileSystemDocumentProvider(new FileSystemDocumentProviderOptions
        {
            ConnectionId = "templates",
            RootPath = root,
            ReadOnly = true,
            TemplateLocation = TemplateLocationKind.Workgroup
        });

        Assert.Empty((await provider.ListAsync(new DocumentListQuery { Templates = true })).Items);
        Assert.False(Directory.Exists(root));
    }

    [Fact]
    public async Task A_read_only_save_is_reported_as_nothing_written_through_the_client()
    {
        using var workspace = new Workspace();
        workspace.Write("contract.docx", DocxFactory.Contract());
        var provider = workspace.Provider(o => o.ReadOnly = true);
        var client = new OfficeAgent.Core.OfficeAgentClient(
            new DocumentProviderRegistry(new IDocumentProvider[] { provider }), new OfficeAgent.Word.WordModule());
        var tools = new OfficeAgent.AgentFramework.OfficeAgentTools(client);

        using var json = System.Text.Json.JsonDocument.Parse(await tools.EditDocument("workspace", "contract.docx",
            "[{\"op\":\"changeText\",\"target\":{\"find\":\"Client:\"},\"with\":\"Fabrikam\",\"mode\":\"Direct\"}]"));

        Assert.False(json.RootElement.GetProperty("committed").GetBoolean());
        Assert.Equal("notWritten", json.RootElement.GetProperty("writeOutcome").GetString());
        Assert.Equal("access-denied", json.RootElement.GetProperty("errors")[0].GetProperty("code").GetString());
    }

    // ── Session memory ────────────────────────────────────────────────────────

    [Fact]
    public async Task A_session_connection_lists_its_documents_by_id()
    {
        var provider = new MemoryDocumentProvider("session");
        var b = provider.Add("b.docx", DocxFactory.Contract());
        var a = provider.Add("A.pptx", PptxFactory.Deck());
        provider.Add("c.xlsx", XlsxFactory.WorkbookWithTable());

        var first = await provider.ListAsync(new DocumentListQuery { MaximumResults = 2 });
        var second = await provider.ListAsync(new DocumentListQuery { MaximumResults = 2, ContinuationToken = first.ContinuationToken });

        Assert.Equal(new[] { "A.pptx", "b.docx" }, first.Items.Select(i => i.Name));
        Assert.Equal(new[] { a.ItemId, b.ItemId }, first.Items.Select(i => i.Source));
        Assert.Equal(new[] { "c.xlsx" }, second.Items.Select(i => i.Name));
        Assert.Null(second.ContinuationToken);
        Assert.All(first.Items, item => Assert.Null(item.Folder));

        // The listed source registers back to the same document; nothing new is minted.
        var registered = await provider.RegisterAsync(first.Items[1].Source);
        Assert.Equal(b.ItemId, registered.ItemId);
        Assert.Equal(3, provider.Count);
    }

    [Fact]
    public async Task A_session_connection_filters_and_has_no_templates_or_folders()
    {
        var provider = new MemoryDocumentProvider("session");
        provider.Add("Quarterly.docx", DocxFactory.Contract());
        provider.Add("deck.pptx", PptxFactory.Deck());

        Assert.Equal(new[] { "Quarterly.docx" },
            (await provider.ListAsync(new DocumentListQuery { NameContains = "quart" })).Items.Select(i => i.Name));
        Assert.Equal(new[] { "deck.pptx" },
            (await provider.ListAsync(new DocumentListQuery { Extensions = new[] { ".pptx" } })).Items.Select(i => i.Name));
        Assert.Empty((await provider.ListAsync(new DocumentListQuery { Templates = true })).Items);
        Assert.Equal(ProviderErrorCode.NotFound, (await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.ListAsync(new DocumentListQuery { Folder = "Contracts" }))).Code);
        Assert.Equal(ProviderErrorCode.InvalidArgument, (await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.RegisterAsync("Quarterly.docx"))).Code);
    }

    // ── Registry ──────────────────────────────────────────────────────────────

    [Fact]
    public void The_registry_reports_which_connections_list_and_hold_templates()
    {
        using var workspace = new Workspace();
        var registry = new DocumentProviderRegistry(new IDocumentProvider[]
        {
            workspace.Provider(o => o.TemplatesFolder = "Templates"),
            new MemoryDocumentProvider("session"),
            new NotListing("plain")
        });

        var listings = registry.DescribeListings();

        Assert.Equal(new[] { "workspace", "session", "plain" }, listings.Select(l => l.ConnectionId));
        Assert.True(listings[0].HasTemplates);
        Assert.Equal("Templates", listings[0].TemplatesFolder);
        Assert.Equal(TemplateLocationKind.Connection, listings[0].TemplateLocation);
        Assert.True(listings[1].Listable);
        Assert.False(listings[1].HasTemplates);
        Assert.False(listings[2].Listable);
        Assert.Equal(ProviderErrorCode.ConfigurationError,
            Assert.Throws<DocumentProviderException>(() => registry.ResolveListing("plain")).Code);
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    /// <summary>A provider that serves documents but cannot list them.</summary>
    internal sealed class NotListing : IDocumentProvider
    {
        public NotListing(string connectionId) => ConnectionId = connectionId;
        public string Provider => "custom";
        public string ConnectionId { get; }
        public Task<DocumentReference> RegisterAsync(string source, CancellationToken cancellationToken = default) => throw new NotSupportedException();
        public Task<DocumentContent> OpenReadAsync(DocumentReference reference, CancellationToken cancellationToken = default) => throw new NotSupportedException();
        public Task<DocumentReference> SaveAsync(DocumentReference source, Stream content, SaveDocumentOptions options, CancellationToken cancellationToken = default) => throw new NotSupportedException();
        public Task RemoveAsync(DocumentReference reference, CancellationToken cancellationToken = default) => throw new NotSupportedException();
    }

    /// <summary>
    /// A directory link by whatever means the machine allows: a junction first on Windows,
    /// which needs no privilege, then a symbolic link.
    /// </summary>
    internal static bool TryCreateDirectoryLink(string path, string target)
    {
        if (OperatingSystem.IsWindows())
        {
            try
            {
                var info = new System.Diagnostics.ProcessStartInfo("cmd.exe")
                {
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                info.ArgumentList.Add("/c");
                info.ArgumentList.Add("mklink");
                info.ArgumentList.Add("/J");
                info.ArgumentList.Add(path);
                info.ArgumentList.Add(target);
                using var process = System.Diagnostics.Process.Start(info);
                if (process is not null && process.WaitForExit(15_000) && process.ExitCode == 0 && Directory.Exists(path))
                    return true;
            }
            catch (Exception)
            {
            }
        }

        try
        {
            Directory.CreateSymbolicLink(path, target);
            return true;
        }
        catch (Exception ex) when (ex is UnauthorizedAccessException or PlatformNotSupportedException or IOException)
        {
            return false;
        }
    }

    internal sealed class Workspace : IDisposable
    {
        public Workspace()
        {
            Root = Path.Combine(Path.GetTempPath(), $"officeagent-listing-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Root);
        }

        public string Root { get; }

        public FileSystemDocumentProvider Provider(Action<FileSystemDocumentProviderOptions>? configure = null)
        {
            var options = new FileSystemDocumentProviderOptions { ConnectionId = "workspace", RootPath = Root };
            configure?.Invoke(options);
            return new FileSystemDocumentProvider(options);
        }

        public string Write(string relativePath, byte[]? bytes = null)
        {
            var path = Path.Combine(Root, relativePath.Replace('/', Path.DirectorySeparatorChar));
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.WriteAllBytes(path, bytes ?? Bytes);
            return path;
        }

        public void Dispose()
        {
            if (!Directory.Exists(Root)) return;
            foreach (var file in Directory.GetFiles(Root, "*", SearchOption.AllDirectories))
                File.SetAttributes(file, FileAttributes.Normal);
            Directory.Delete(Root, recursive: true);
        }
    }
}
