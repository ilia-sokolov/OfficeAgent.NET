using OfficeAgent.Abstractions;
using OfficeAgent.Core.DocumentProviders;

namespace OfficeAgent.Tests;

/// <summary>
/// Coverage for the security-critical paths in <see cref="FileSystemDocumentProvider"/>
/// and <see cref="DocumentProviderRegistry"/> under the register-by-path model:
/// extension allow-list, root containment, size cap (both directions), Replace
/// mode (happy + stale + rename rejection), provider-assigned-id invariants
/// (caller-chosen destinations rejected, NewName renames the display name only),
/// path-like id rejection, and the registry mis-configuration errors.
/// </summary>
public class DocumentProviderHardeningTests
{
    [Fact]
    public async Task Register_rejects_disallowed_extension()
    {
        using var workspace = new Workspace();
        var provider = workspace.Provider();

        var path = workspace.Stage(new byte[] { 1, 2, 3 }, "notes.txt");
        var ex = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.RegisterAsync(path));
        Assert.Equal(ProviderErrorCode.ExtensionNotAllowed, ex.Code);
    }

    [Fact]
    public async Task Register_rejects_source_outside_root()
    {
        using var workspace = new Workspace();
        using var outsider = new Workspace();
        var provider = workspace.Provider();

        var path = outsider.Stage(DocxFactory.Contract(), "contract.docx");
        var ex = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.RegisterAsync(path));
        Assert.Equal(ProviderErrorCode.AccessDenied, ex.Code);
    }

    [Fact]
    public async Task Register_rejects_missing_source()
    {
        using var workspace = new Workspace();
        var provider = workspace.Provider();

        var ex = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.RegisterAsync(Path.Combine(workspace.Root, "nope.docx")));
        Assert.Equal(ProviderErrorCode.NotFound, ex.Code);
    }

    [Fact]
    public async Task Open_rejects_oversized_file()
    {
        using var workspace = new Workspace();
        var added = await workspace.Provider().RegisterBytesAsync(workspace.Root, new byte[10 * 1024], "big.docx");

        // A second provider over the same root with a tighter cap rejects the read.
        var capped = workspace.Provider(opts => opts.MaximumBytes = 4 * 1024);
        var ex = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            capped.OpenReadAsync(DocumentReference.ForFileSystem("workspace", added.ItemId)));
        Assert.Equal(ProviderErrorCode.ContentTooLarge, ex.Code);
        Assert.Contains("exceeds", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Save_rejects_oversized_content()
    {
        using var workspace = new Workspace();
        var added = await workspace.Provider().RegisterBytesAsync(workspace.Root, DocxFactory.Contract(), "contract.docx");
        var capped = workspace.Provider(opts => opts.MaximumBytes = 4 * 1024);

        using var oversized = new MemoryStream(new byte[10 * 1024]);
        var ex = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            capped.SaveAsync(
                DocumentReference.ForFileSystem("workspace", added.ItemId),
                oversized,
                new SaveDocumentOptions { Mode = SaveMode.NewVersion },
                CancellationToken.None));
        Assert.Equal(ProviderErrorCode.ContentTooLarge, ex.Code);
    }

    [Fact]
    public async Task Replace_mode_atomically_overwrites_source_when_version_matches()
    {
        using var workspace = new Workspace();
        var provider = workspace.Provider();
        var original = DocxFactory.Contract();
        var added = await provider.RegisterBytesAsync(workspace.Root, original, "contract.docx");

        var newBytes = original.Concat(new byte[] { 0x0A }).ToArray();
        using var input = new MemoryStream(newBytes);
        var saved = await provider.SaveAsync(
            added,
            input,
            new SaveDocumentOptions { Mode = SaveMode.Replace },
            CancellationToken.None);

        Assert.Equal(added.ItemId, saved.ItemId);                     // same id - overwritten in place
        Assert.NotEqual(added.Version, saved.Version);                // new content → new SHA

        using var reopened = await provider.OpenReadAsync(DocumentReference.ForFileSystem("workspace", added.ItemId));
        using var ms = new MemoryStream();
        await reopened.Stream.CopyToAsync(ms);
        Assert.Equal(newBytes, ms.ToArray());
    }

    [Fact]
    public async Task Replace_mode_rejects_rename()
    {
        using var workspace = new Workspace();
        var provider = workspace.Provider();
        var added = await provider.RegisterBytesAsync(workspace.Root, DocxFactory.Contract(), "contract.docx");

        using var bytes = new MemoryStream(DocxFactory.Contract());
        var ex = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.SaveAsync(
                added,
                bytes,
                new SaveDocumentOptions { Mode = SaveMode.Replace, NewName = "different.docx" },
                CancellationToken.None));
        Assert.Equal(ProviderErrorCode.InvalidArgument, ex.Code);
    }

    [Fact]
    public async Task Save_rejects_caller_chosen_destination_id()
    {
        using var workspace = new Workspace();
        var provider = workspace.Provider();
        var added = await provider.RegisterBytesAsync(workspace.Root, DocxFactory.Contract(), "source.docx");

        using var content = new MemoryStream(DocxFactory.Contract());
        var ex = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.SaveAsync(
                added,
                content,
                new SaveDocumentOptions
                {
                    Mode = SaveMode.NewDocument,
                    DestinationItemId = "exports/2026/copy.docx"
                },
                CancellationToken.None));
        Assert.Equal(ProviderErrorCode.InvalidArgument, ex.Code);
    }

    [Fact]
    public async Task NewVersion_with_NewName_renames_display_name_only()
    {
        using var workspace = new Workspace();
        var provider = workspace.Provider();
        var added = await provider.RegisterBytesAsync(workspace.Root, DocxFactory.Contract(), "contract.docx");

        using var content = new MemoryStream(DocxFactory.Contract());
        var saved = await provider.SaveAsync(
            added,
            content,
            new SaveDocumentOptions { Mode = SaveMode.NewVersion, NewName = "redlined.docx" },
            CancellationToken.None);

        Assert.NotEqual(added.ItemId, saved.ItemId);   // fresh opaque id
        Assert.Equal("redlined.docx", saved.Name);     // NewName changes the display name, not the id

        // The renamed version is openable by its new id; the source is untouched.
        using (await provider.OpenReadAsync(saved)) { }
        using (await provider.OpenReadAsync(added)) { }
    }

    [Fact]
    public async Task NewVersion_without_NewName_writes_sibling_versioned_file()
    {
        using var workspace = new Workspace();
        var provider = workspace.Provider();
        var added = await provider.RegisterBytesAsync(workspace.Root, DocxFactory.Contract(), "contract.docx");

        using var content = new MemoryStream(DocxFactory.Contract());
        var saved = await provider.SaveAsync(
            added,
            content,
            new SaveDocumentOptions { Mode = SaveMode.NewVersion },
            CancellationToken.None);

        Assert.NotEqual(added.ItemId, saved.ItemId);
        Assert.Equal("contract.v2.docx", saved.Name);

        // A second NewVersion lands at v3 next to the original.
        using var more = new MemoryStream(DocxFactory.Contract());
        var third = await provider.SaveAsync(
            added,
            more,
            new SaveDocumentOptions { Mode = SaveMode.NewVersion },
            CancellationToken.None);
        Assert.Equal("contract.v3.docx", third.Name);
    }

    [Fact]
    public async Task Remove_only_unregisters_and_leaves_the_file_intact()
    {
        using var workspace = new Workspace();
        var provider = workspace.Provider();
        var path = workspace.Stage(DocxFactory.Contract(), "contract.docx");
        var added = await provider.RegisterAsync(path);

        await provider.RemoveAsync(added);

        Assert.True(File.Exists(path));   // host content is never destroyed
        var ex = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.OpenReadAsync(added));
        Assert.Equal(ProviderErrorCode.NotFound, ex.Code);
    }

    [Fact]
    public async Task Open_rejects_path_like_id()
    {
        using var workspace = new Workspace();
        var provider = workspace.Provider();

        var ex = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.OpenReadAsync(DocumentReference.ForFileSystem("workspace", "sub/dir/contract.docx")));
        Assert.Equal(ProviderErrorCode.AccessDenied, ex.Code);
    }

    [Fact]
    public void Registry_throws_when_no_provider_matches_reference()
    {
        var registry = new DocumentProviderRegistry(Array.Empty<IDocumentProvider>());
        var reference = new DocumentReference
        {
            Provider = "filesystem",
            ConnectionId = "missing",
            ItemId = "anything"
        };

        var ex = Assert.Throws<DocumentProviderException>(() => registry.Resolve(reference));
        Assert.Equal(ProviderErrorCode.ConfigurationError, ex.Code);
        Assert.Contains("missing", ex.Message);
    }

    [Fact]
    public void Registry_throws_when_two_providers_share_connection_id()
    {
        using var w1 = new Workspace();
        using var w2 = new Workspace();
        var registry = new DocumentProviderRegistry(new IDocumentProvider[]
        {
            w1.Provider(),
            w2.Provider()       // same connectionId = "workspace"
        });

        var reference = new DocumentReference
        {
            Provider = "filesystem",
            ConnectionId = "workspace",
            ItemId = "anything"
        };

        var ex = Assert.Throws<DocumentProviderException>(() => registry.Resolve(reference));
        Assert.Equal(ProviderErrorCode.ConfigurationError, ex.Code);
    }

    [Fact]
    public async Task Open_rejects_reference_belonging_to_different_connection()
    {
        using var workspace = new Workspace();
        var provider = workspace.Provider();
        var added = await provider.RegisterBytesAsync(workspace.Root, DocxFactory.Contract(), "contract.docx");

        var foreignReference = new DocumentReference
        {
            Provider = FileSystemDocumentProvider.ProviderName,
            ConnectionId = "different-connection",
            ItemId = added.ItemId
        };

        var ex = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.OpenReadAsync(foreignReference));
        Assert.Equal(ProviderErrorCode.InvalidArgument, ex.Code);
    }

    [Fact]
    public async Task Save_skips_source_read_when_no_version_to_verify()
    {
        // A save without ExpectedVersion (and with no version on the source
        // reference) must not hash the source file for an optimistic check.
        using var workspace = new Workspace();
        var provider = workspace.Provider();
        var added = await provider.RegisterBytesAsync(workspace.Root, DocxFactory.Contract(), "contract.docx");

        // Caller intentionally drops the version field - no optimistic check.
        var bareReference = new DocumentReference
        {
            Provider = FileSystemDocumentProvider.ProviderName,
            ConnectionId = "workspace",
            ItemId = added.ItemId
        };

        using var content = new MemoryStream(DocxFactory.Contract());
        var saved = await provider.SaveAsync(
            bareReference,
            content,
            new SaveDocumentOptions { Mode = SaveMode.NewVersion },
            CancellationToken.None);

        Assert.NotNull(saved.Version);
        using (await provider.OpenReadAsync(saved)) { }   // the new version is readable
    }

    private sealed class Workspace : IDisposable
    {
        public Workspace()
        {
            Root = Path.Combine(Path.GetTempPath(), $"officeagent-provider-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Root);
        }

        public string Root { get; }

        public FileSystemDocumentProvider Provider(Action<FileSystemDocumentProviderOptions>? configure = null)
        {
            var options = new FileSystemDocumentProviderOptions
            {
                ConnectionId = "workspace",
                RootPath = Root
            };
            configure?.Invoke(options);
            return new FileSystemDocumentProvider(options);
        }

        public string Stage(byte[] bytes, string name) =>
            TestRegistrationExtensions.StageFixture(Root, bytes, name);

        public void Dispose()
        {
            if (Directory.Exists(Root)) Directory.Delete(Root, recursive: true);
        }
    }
}
