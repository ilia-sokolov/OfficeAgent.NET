using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.Abstractions;
using OfficeAgent.AmazonS3;
using OfficeAgent.Core.DocumentProviders;

namespace OfficeAgent.Tests;

/// <summary>
/// The Amazon S3 provider over <see cref="FakeS3Service"/>: the real AWS SDK client sends its
/// requests to an in-memory bucket, so <c>If-Match</c> and <c>If-None-Match</c> travel as real headers.
/// </summary>
public sealed class AmazonS3ProviderTests : ObjectStoreProviderTests
{
    private protected override FakeObjectStore NewStore() => new FakeS3Service();

    private protected override IDocumentCreatingProvider Provider(
        FakeObjectStore store, string prefix = Prefix, long maximumBytes = 100 * 1024 * 1024,
        string? registrationIndexPath = null, bool registerOnce = false) =>
        new AmazonS3DocumentProvider(
            new AmazonS3DocumentProviderOptions
            {
                ConnectionId = Connection,
                BucketName = FakeS3Service.Bucket,
                KeyPrefix = prefix,
                MaximumBytes = maximumBytes,
                RegistrationIndexPath = registrationIndexPath ?? string.Empty,
                AllowedExtensions = new[] { ".docx" }
            },
            ((FakeS3Service)store).Client(),
            registerOnce ? new RegisterOnceStore() : null);

    protected override string ProviderName => AmazonS3DocumentProvider.ProviderName;

    protected override string LocatorFor(string fullName) => $"s3://{FakeS3Service.Bucket}/{fullName}";

    protected override string ForeignLocatorFor(string fullName) => $"s3://another-bucket/{fullName}";

    protected override string StorageIdentity => FakeS3Service.Bucket;

    protected override int PreconditionStatusForCreate => 412;
    protected override string PreconditionCodeForCreate => "PreconditionFailed";
    protected override string PreconditionCodeForReplace => "PreconditionFailed";

    [Fact]
    public async Task With_versioning_the_reference_carries_the_s3_version_id()
    {
        var store = new FakeS3Service { Versioning = true };
        store.Seed("contracts/acme.docx", Contract);
        var provider = Provider(store);

        var reference = await provider.RegisterAsync("acme.docx");
        var etag = store.Objects["contracts/acme.docx"].ETag;
        Assert.Equal(store.Objects["contracts/acme.docx"].VersionId, reference.Version);

        using (var opened = await provider.OpenReadAsync(reference))
            Assert.Equal(reference.Version, opened.Reference.Version);
        // The download names the exact version that was checked.
        Assert.Contains("versionId=" + reference.Version, store.Requests.Single(request => request.Method == "GET").Query);

        var saved = await provider.SaveAsync(reference, new MemoryStream(Contract),
            new SaveDocumentOptions { Mode = SaveMode.Replace });
        Assert.Equal(store.Objects["contracts/acme.docx"].VersionId, saved.Version);
        Assert.NotEqual(reference.Version, saved.Version);
        // The write is still conditioned on the ETag: S3 compares If-Match with the current ETag.
        Assert.Equal(etag, store.Requests.Single(request => request.Method == "PUT").Headers["If-Match"]);
    }

    [Fact]
    public async Task A_conditional_request_conflict_is_a_conflict_on_replace_and_a_refusal_on_create()
    {
        var store = new FakeS3Service();
        store.Seed("contracts/acme.docx", Contract);
        var provider = Provider(store);
        var reference = await provider.RegisterAsync("acme.docx");
        store.WriteStatus = 409;
        store.WriteErrorCode = "ConditionalRequestConflict";

        var onReplace = await Assert.ThrowsAsync<DocumentVersionConflictException>(() => provider.SaveAsync(
            reference, new MemoryStream(Contract), new SaveDocumentOptions { Mode = SaveMode.Replace }));
        var onCreate = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.CreateAsync("brief.docx", new MemoryStream(Contract)));

        Assert.Equal(ProviderErrorCode.VersionConflict, onReplace.Code);
        Assert.Equal(ProviderErrorCode.WriteRejected, onCreate.Code);
        Assert.False(store.Has("contracts/brief.docx"));
    }

    [Fact]
    public async Task Writes_carry_the_content_type_and_a_write_marker()
    {
        var store = new FakeS3Service();
        var provider = Provider(store);

        await provider.CreateAsync("brief.docx", new MemoryStream(Contract));

        var put = store.Requests.Single(request => request.Method == "PUT");
        Assert.Equal("*", put.Headers["If-None-Match"]);
        Assert.StartsWith("application/vnd.openxmlformats-officedocument.wordprocessingml.document", store.Objects["contracts/brief.docx"].ContentType);
        Assert.Matches("^[0-9a-f]{32}$", store.Objects["contracts/brief.docx"].Metadata["officeagent-write-id"]);
    }

    [Fact]
    public void Configuration_needs_a_bucket_and_a_plain_prefix()
    {
        Assert.Throws<ArgumentException>(() => new AmazonS3DocumentProvider(new AmazonS3DocumentProviderOptions
        {
            ConnectionId = "c", Region = "eu-west-1"
        }));
        Assert.Throws<ArgumentException>(() => new AmazonS3DocumentProvider(new AmazonS3DocumentProviderOptions
        {
            ConnectionId = "c", BucketName = "b", Region = "eu-west-1", KeyPrefix = "../b"
        }));
    }

    [Fact]
    public void Default_chain_configuration_builds_without_contacting_aws()
    {
        var provider = new AmazonS3DocumentProvider(new AmazonS3DocumentProviderOptions
        {
            ConnectionId = "legal",
            BucketName = "contracts",
            Region = "eu-west-1"
        });
        var compatible = new AmazonS3DocumentProvider(new AmazonS3DocumentProviderOptions
        {
            ConnectionId = "minio",
            BucketName = "contracts",
            ServiceUrl = "http://localhost:9000",
            ForcePathStyle = true
        });

        Assert.Equal(AmazonS3DocumentProvider.ProviderName, provider.Provider);
        Assert.Equal("minio", compatible.ConnectionId);
    }

    [Fact]
    public async Task The_client_factory_registration_resolves_a_working_provider()
    {
        var store = new FakeS3Service();
        store.Seed("acme.docx", Contract);
        var services = new ServiceCollection();
        services.AddAmazonS3DocumentProvider("legal", _ => store.Client(), o => o.BucketName = FakeS3Service.Bucket);
        using var built = services.BuildServiceProvider();

        var provider = Assert.IsType<AmazonS3DocumentProvider>(built.GetRequiredService<IDocumentProvider>());
        var reference = await provider.RegisterAsync("acme.docx");

        Assert.Equal("acme.docx", reference.Name);
    }

    /// <summary>Accepts the first registration, then fails every later one.</summary>
    private sealed class RegisterOnceStore : IAmazonS3RegistrationStore
    {
        private readonly Dictionary<string, string> _items = new();

        public Task<string> AddAsync(string key, CancellationToken cancellationToken = default)
        {
            if (_items.Count > 0) throw new IOException("Registration store unavailable.");
            _items["reg0"] = key;
            return Task.FromResult("reg0");
        }

        public Task<string?> ResolveAsync(string registrationId, CancellationToken cancellationToken = default) =>
            Task.FromResult(_items.TryGetValue(registrationId, out var key) ? key : null);

        public Task<bool> RemoveAsync(string registrationId, CancellationToken cancellationToken = default) =>
            Task.FromResult(_items.Remove(registrationId));
    }
}
