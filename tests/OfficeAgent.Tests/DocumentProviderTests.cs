using Microsoft.Extensions.DependencyInjection;
using DocumentFormat.OpenXml.Packaging;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

public class DocumentProviderTests
{
    [Fact]
    public async Task Filesystem_provider_assigns_opaque_id_and_opens_by_it()
    {
        using var workspace = new TemporaryWorkspace();
        var provider = workspace.Provider();

        var added = await provider.RegisterBytesAsync(workspace.Root, DocxFactory.Contract(), "contract.docx");

        // The id is opaque: not the supplied name, not a path segment.
        Assert.NotEqual("contract.docx", added.ItemId);
        Assert.DoesNotContain('/', added.ItemId);
        Assert.DoesNotContain('\\', added.ItemId);
        Assert.Equal("contract.docx", added.Name);

        using var content = await provider.OpenReadAsync(DocumentReference.ForFileSystem("workspace", added.ItemId));

        Assert.StartsWith("sha256:", content.Reference.Version);
        Assert.Equal(added.ItemId, content.Reference.ItemId);
        Assert.Equal("contract.docx", content.Reference.Name);
        Assert.True(content.Stream.Length > 0);
    }

    [Fact]
    public async Task Filesystem_provider_rejects_path_like_ids()
    {
        using var workspace = new TemporaryWorkspace();
        var provider = workspace.Provider();

        var ex = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.OpenReadAsync(DocumentReference.ForFileSystem("workspace", "../outside.docx")));
        Assert.Equal(ProviderErrorCode.AccessDenied, ex.Code);
    }

    [Fact]
    public async Task Filesystem_provider_rejects_unknown_id()
    {
        using var workspace = new TemporaryWorkspace();
        var provider = workspace.Provider();

        var ex = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.OpenReadAsync(DocumentReference.ForFileSystem("workspace", "deadbeef")));
        Assert.Equal(ProviderErrorCode.NotFound, ex.Code);
    }

    [Fact]
    public async Task Filesystem_provider_removes_document_by_opaque_id()
    {
        using var workspace = new TemporaryWorkspace();
        var provider = workspace.Provider();
        var added = await provider.RegisterBytesAsync(workspace.Root, DocxFactory.Contract(), "contract.docx");

        await provider.RemoveAsync(added);

        var ex = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.OpenReadAsync(added));
        Assert.Equal(ProviderErrorCode.NotFound, ex.Code);
    }

    [Fact]
    public async Task Filesystem_provider_rejects_removing_unknown_id()
    {
        using var workspace = new TemporaryWorkspace();
        var provider = workspace.Provider();

        var ex = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.RemoveAsync(DocumentReference.ForFileSystem("workspace", "deadbeef")));
        Assert.Equal(ProviderErrorCode.NotFound, ex.Code);
    }

    [Fact]
    public async Task Provider_client_commits_to_non_destructive_new_version()
    {
        using var workspace = new TemporaryWorkspace();

        var services = new ServiceCollection();
        services.AddWordFormat();
        services.AddFileSystemDocumentProvider("workspace", workspace.Root);
        services.AddOfficeAgent();

        using var serviceProvider = services.BuildServiceProvider();
        var client = serviceProvider.GetRequiredService<OfficeAgentClient>();
        var source = await client.RegisterBytesAsync("workspace", workspace.Root, DocxFactory.Contract(), "contract.docx");
        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new SetPropertyOp
                {
                    Target = new NodeAnchor { Kind = "docProperty", Path = "core/title" },
                    Value = "Updated title"
                }
            }
        };

        var result = await client.CommitAsync(source, plan);

        Assert.True(result.Committed);
        Assert.NotNull(result.Document);
        // The new version receives a fresh opaque id and a sibling versioned file
        // beside the source; the source id keeps its original content.
        Assert.NotEqual(source.ItemId, result.Document!.ItemId);
        Assert.Equal("contract.v2.docx", result.Document.Name);

        using var original = await client.OpenReadAsync(source);
        using var originalDoc = WordprocessingDocument.Open(original.Stream, false);
        Assert.NotEqual("Updated title", originalDoc.PackageProperties.Title);

        using var updated = await client.OpenReadAsync(result.Document);
        using var updatedDoc = WordprocessingDocument.Open(updated.Stream, false);
        Assert.Equal("Updated title", updatedDoc.PackageProperties.Title);
    }

    [Fact]
    public async Task Filesystem_provider_rejects_stale_save()
    {
        using var workspace = new TemporaryWorkspace();
        var provider = workspace.Provider();
        var added = await provider.RegisterBytesAsync(workspace.Root, DocxFactory.Contract(), "contract.docx");

        // Overwrite the stored content so the recorded version goes stale.
        using (var fresh = new MemoryStream(DocxFactory.Contract().Concat(new byte[] { 0 }).ToArray()))
            await provider.SaveAsync(added, fresh, new SaveDocumentOptions { Mode = SaveMode.Replace }, CancellationToken.None);

        using var output = new MemoryStream(DocxFactory.Contract());
        await Assert.ThrowsAsync<DocumentVersionConflictException>(() => provider.SaveAsync(
            added,          // still carries the original (now stale) version
            output,
            new SaveDocumentOptions { Mode = SaveMode.Replace },
            CancellationToken.None));
    }

    private sealed class TemporaryWorkspace : IDisposable
    {
        public TemporaryWorkspace()
        {
            Root = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"officeagent-provider-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Root);
        }

        public string Root { get; }

        public FileSystemDocumentProvider Provider() => new(new FileSystemDocumentProviderOptions
        {
            ConnectionId = "workspace",
            RootPath = Root
        });

        public void Dispose()
        {
            if (Directory.Exists(Root)) Directory.Delete(Root, recursive: true);
        }
    }
}
