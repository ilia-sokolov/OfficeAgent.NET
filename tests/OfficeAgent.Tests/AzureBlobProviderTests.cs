using Azure.Storage.Blobs;
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.Abstractions;
using OfficeAgent.AzureBlob;
using OfficeAgent.Core.DocumentProviders;

namespace OfficeAgent.Tests;

/// <summary>
/// The Azure Blob provider over <see cref="FakeBlobService"/>: the real Azure.Storage.Blobs client
/// sends its requests to an in-memory container, so preconditions travel as real headers.
/// </summary>
public sealed class AzureBlobProviderTests : ObjectStoreProviderTests
{
    private protected override FakeObjectStore NewStore() => new FakeBlobService();

    private protected override IDocumentCreatingProvider Provider(
        FakeObjectStore store, string prefix = Prefix, long maximumBytes = 100 * 1024 * 1024,
        string? registrationIndexPath = null, bool registerOnce = false, string? templatesFolder = null,
        IReadOnlyCollection<string>? allowedExtensions = null, bool readOnly = false) =>
        new AzureBlobDocumentProvider(
            new AzureBlobDocumentProviderOptions
            {
                ConnectionId = Connection,
                Prefix = prefix,
                MaximumBytes = maximumBytes,
                RegistrationIndexPath = registrationIndexPath ?? string.Empty,
                AllowedExtensions = allowedExtensions ?? new[] { ".docx" },
                TemplatesFolder = templatesFolder ?? string.Empty,
                ReadOnly = readOnly
            },
            ((FakeBlobService)store).Client(),
            registerOnce ? new RegisterOnceStore() : null);

    protected override string ProviderName => AzureBlobDocumentProvider.ProviderName;

    protected override string LocatorFor(string fullName) => $"{FakeBlobService.ContainerUri}/{fullName}";

    protected override string ForeignLocatorFor(string fullName) =>
        $"https://{FakeBlobService.Account}.blob.core.windows.net/other/{fullName}";

    protected override string StorageIdentity => FakeBlobService.Account;

    protected override int PreconditionStatusForCreate => 409;
    protected override string PreconditionCodeForCreate => "BlobAlreadyExists";
    protected override string PreconditionCodeForReplace => "ConditionNotMet";

    [Fact]
    public async Task A_leased_blob_says_why_the_save_was_refused()
    {
        var store = new FakeBlobService();
        store.Seed("contracts/acme.docx", Contract);
        var provider = Provider(store);
        var reference = await provider.RegisterAsync("acme.docx");
        store.WriteStatus = 409;
        store.WriteErrorCode = "LeaseIdMissing";

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.SaveAsync(
            reference, new MemoryStream(Contract), new SaveDocumentOptions { Mode = SaveMode.Replace }));

        Assert.Equal(ProviderErrorCode.WriteRejected, error.Code);
        Assert.Contains("active lease", error.Message);
        Assert.Contains("Storage code: LeaseIdMissing", error.Message);
    }

    [Fact]
    public async Task Uploads_carry_the_content_type_and_a_write_marker()
    {
        var store = new FakeBlobService();
        var provider = Provider(store);

        await provider.CreateAsync("brief.docx", new MemoryStream(Contract));

        var put = store.Requests.Single(request => request.Method == "PUT");
        Assert.Equal("*", put.Headers["If-None-Match"]);
        Assert.Equal("BlockBlob", put.Headers["x-ms-blob-type"]);
        Assert.Matches("^[0-9a-f]{32}$", store.Objects["contracts/brief.docx"].Metadata["officeagent_write_id"]);
    }

    [Fact]
    public void Configuration_needs_exactly_one_way_to_reach_the_account()
    {
        Assert.Throws<ArgumentException>(() => new AzureBlobDocumentProvider(new AzureBlobDocumentProviderOptions
        {
            ConnectionId = "c", ContainerName = "docs"
        }));
        Assert.Throws<ArgumentException>(() => new AzureBlobDocumentProvider(new AzureBlobDocumentProviderOptions
        {
            ConnectionId = "c", ContainerName = "docs",
            ServiceUri = "https://contoso.blob.core.windows.net", ConnectionString = "UseDevelopmentStorage=true"
        }));
        Assert.Throws<ArgumentException>(() => new AzureBlobDocumentProvider(new AzureBlobDocumentProviderOptions
        {
            ConnectionId = "c", ServiceUri = "https://contoso.blob.core.windows.net"
        }));
        Assert.Throws<ArgumentException>(() => new AzureBlobDocumentProvider(new AzureBlobDocumentProviderOptions
        {
            ConnectionId = "c", ContainerName = "docs", ServiceUri = "https://contoso.blob.core.windows.net", Prefix = "a/../b"
        }));
    }

    [Fact]
    public void A_malformed_connection_string_is_not_echoed()
    {
        const string marker = "not-a-real-secret-value";
        var error = Assert.Throws<ArgumentException>(() => new AzureBlobDocumentProvider(new AzureBlobDocumentProviderOptions
        {
            ConnectionId = "c", ContainerName = "docs", ConnectionString = "garbage " + marker
        }));

        Assert.DoesNotContain(marker, error.ToString());
    }

    [Fact]
    public void Managed_identity_configuration_builds_without_contacting_azure()
    {
        var provider = new AzureBlobDocumentProvider(new AzureBlobDocumentProviderOptions
        {
            ConnectionId = "legal",
            ServiceUri = "https://contoso.blob.core.windows.net",
            ContainerName = "contracts",
            ManagedIdentityClientId = "00000000-0000-0000-0000-000000000000"
        });

        Assert.Equal("legal", provider.ConnectionId);
        Assert.Equal(AzureBlobDocumentProvider.ProviderName, provider.Provider);
    }

    [Fact]
    public async Task The_container_factory_registration_resolves_a_working_provider()
    {
        var store = new FakeBlobService();
        store.Seed("legal/acme.docx", Contract);
        var services = new ServiceCollection();
        services.AddAzureBlobDocumentProvider("legal", _ => store.Client(), o =>
        {
            o.Prefix = "legal";
            o.DefaultChangeMode = ChangeMode.Direct;
        });
        using var built = services.BuildServiceProvider();

        var provider = Assert.IsType<AzureBlobDocumentProvider>(built.GetRequiredService<IDocumentProvider>());
        var reference = await provider.RegisterAsync("acme.docx");

        Assert.Equal(ChangeMode.Direct, provider.DefaultChangeMode);
        Assert.Equal("acme.docx", reference.Name);
    }

    /// <summary>Accepts the first registration, then fails every later one.</summary>
    private sealed class RegisterOnceStore : IAzureBlobRegistrationStore
    {
        private readonly Dictionary<string, string> _items = new();

        public Task<string> AddAsync(string blobName, CancellationToken cancellationToken = default)
        {
            if (_items.Count > 0) throw new IOException("Registration store unavailable.");
            _items["reg0"] = blobName;
            return Task.FromResult("reg0");
        }

        public Task<string?> ResolveAsync(string registrationId, CancellationToken cancellationToken = default) =>
            Task.FromResult(_items.TryGetValue(registrationId, out var name) ? name : null);

        public Task<bool> RemoveAsync(string registrationId, CancellationToken cancellationToken = default) =>
            Task.FromResult(_items.Remove(registrationId));
    }
}
