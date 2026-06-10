using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// Integration coverage for the provider boundary: typed exception hierarchy,
/// opaque-id storage round-tripping on <see cref="OfficeAgentClient"/>, registry
/// enumeration, and <see cref="DocumentReference"/> factories.
/// </summary>
public class ProviderIntegrationTests
{
    [Fact]
    public void DocumentVersionConflictException_is_a_DocumentProviderException()
    {
        var ex = new DocumentVersionConflictException("v1", "v2", "filesystem", "ws", "x.docx");
        DocumentProviderException baseEx = ex;        // assignment compiles
        Assert.Equal(ProviderErrorCode.VersionConflict, baseEx.Code);
        Assert.Equal("filesystem", baseEx.Provider);
        Assert.Equal("ws", baseEx.ConnectionId);
        Assert.Equal("x.docx", baseEx.ItemId);
    }

    [Fact]
    public void Single_catch_handles_every_provider_failure()
    {
        // Demonstrates the DX win: developer writes one catch block.
        var thrown = new List<ProviderErrorCode>();
        foreach (var ex in new DocumentProviderException[]
        {
            new(ProviderErrorCode.NotFound, "x"),
            new(ProviderErrorCode.AccessDenied, "x"),
            new(ProviderErrorCode.ContentTooLarge, "x"),
            new(ProviderErrorCode.ExtensionNotAllowed, "x"),
            new(ProviderErrorCode.VersionConflict, "x"),
            new(ProviderErrorCode.InvalidArgument, "x"),
            new(ProviderErrorCode.ConfigurationError, "x"),
            new(ProviderErrorCode.IO, "x")
        })
        {
            try { throw ex; }
            catch (DocumentProviderException caught) { thrown.Add(caught.Code); }
        }
        Assert.Equal(8, thrown.Count);
    }

    [Fact]
    public void DocumentReference_factories_produce_canonical_references()
    {
        var fs = DocumentReference.ForFileSystem("contracts", "1a2b3c");
        Assert.Equal("filesystem", fs.Provider);
        Assert.Equal("contracts", fs.ConnectionId);
        Assert.Equal("1a2b3c", fs.ItemId);

        var generic = DocumentReference.For("sharepoint", "tenant1", "items/42", version: "etag-x");
        Assert.Equal("sharepoint", generic.Provider);
        Assert.Equal("etag-x", generic.Version);
    }

    [Fact]
    public void Registry_Connections_lists_registered_providers_and_Contains_works()
    {
        using var w1 = new Workspace();
        using var w2 = new Workspace(connectionId: "templates");
        var registry = new DocumentProviderRegistry(new IDocumentProvider[] { w1.Provider(), w2.Provider() });

        Assert.Contains(("filesystem", "workspace"), registry.Connections);
        Assert.Contains(("filesystem", "templates"), registry.Connections);
        Assert.True(registry.Contains("filesystem", "workspace"));
        Assert.False(registry.Contains("filesystem", "unknown"));
    }

    [Fact]
    public async Task OfficeAgentClient_addresses_documents_by_provider_assigned_id()
    {
        using var workspace = new Workspace();
        var client = workspace.Client();

        var reference = await client.RegisterBytesAsync("workspace", workspace.Root, DocxFactory.Contract(), "contract.docx");

        // The id is opaque - not the supplied name, not a path.
        Assert.NotEqual("contract.docx", reference.ItemId);
        Assert.DoesNotContain('/', reference.ItemId);
        Assert.DoesNotContain('\\', reference.ItemId);

        var inspect = await client.InspectAsync("workspace", reference.ItemId);
        Assert.NotEmpty(inspect.Paragraphs);
    }

    [Fact]
    public async Task OfficeAgentClient_OpenReadAsync_returns_bytes_and_canonical_reference()
    {
        using var workspace = new Workspace();
        var client = workspace.Client();
        var added = await client.RegisterBytesAsync("workspace", workspace.Root, DocxFactory.Contract(), "contract.docx");

        using var content = await client.OpenReadAsync(added);

        Assert.StartsWith("sha256:", content.Reference.Version);
        Assert.Equal(added.ItemId, content.Reference.ItemId);
        Assert.Equal("contract.docx", content.Reference.Name);
        Assert.True(content.Stream.Length > 0);
    }

    [Fact]
    public void ProviderApplyResult_Committed_unlocks_Document_via_MemberNotNullWhen()
    {
        var failed = new ProviderApplyResult { Committed = false };
        Assert.Null(failed.Document);

        var ok = new ProviderApplyResult
        {
            Committed = true,
            Document = DocumentReference.ForFileSystem("ws", "x.docx", version: "v1")
        };
        // After the Committed check the compiler treats Document as non-null
        // (no `!` needed). Validated indirectly by reading without a null check.
        if (ok.Committed)
            Assert.Equal("x.docx", ok.Document.ItemId);
    }

    private sealed class Workspace : IDisposable
    {
        public Workspace(string connectionId = "workspace")
        {
            ConnectionId = connectionId;
            Root = Path.Combine(Path.GetTempPath(), $"officeagent-provider-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Root);
        }

        public string Root { get; }
        public string ConnectionId { get; }

        public FileSystemDocumentProvider Provider() => new(new FileSystemDocumentProviderOptions
        {
            ConnectionId = ConnectionId,
            RootPath = Root
        });

        public OfficeAgentClient Client()
        {
            var services = new ServiceCollection();
            services.AddWordFormat();
            services.AddFileSystemDocumentProvider(ConnectionId, Root);
            services.AddOfficeAgent();
            _serviceProvider = services.BuildServiceProvider();
            return _serviceProvider.GetRequiredService<OfficeAgentClient>();
        }

        private ServiceProvider? _serviceProvider;

        public void Dispose()
        {
            _serviceProvider?.Dispose();
            if (Directory.Exists(Root)) Directory.Delete(Root, recursive: true);
        }
    }
}
