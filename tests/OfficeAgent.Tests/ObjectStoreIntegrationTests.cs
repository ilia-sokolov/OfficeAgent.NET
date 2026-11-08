using Amazon.S3;
using Amazon.S3.Model;
using Azure.Identity;
using Azure.Storage.Blobs;
using OfficeAgent.Abstractions;
using OfficeAgent.AmazonS3;
using OfficeAgent.AzureBlob;
using OfficeAgent.Core.DocumentProviders;

namespace OfficeAgent.Tests;

/// <summary>
/// Opt-in tests against a real Azure Storage account (or Azurite) and a real S3 bucket (or an
/// S3-compatible store). Each is skipped unless its environment variables are set. Every run works
/// under a fresh prefix and deletes what it wrote there; nothing outside that prefix is touched.
/// </summary>
/// <remarks>
/// Azure Blob: <c>OFFICEAGENT_TEST_AZUREBLOB_CONTAINER</c>, plus either
/// <c>OFFICEAGENT_TEST_AZUREBLOB_SERVICE_URI</c> (signed in through <c>DefaultAzureCredential</c>)
/// or <c>OFFICEAGENT_TEST_AZUREBLOB_CONNECTION_STRING</c> (for Azurite,
/// <c>UseDevelopmentStorage=true</c>). The container must exist.
/// Amazon S3: <c>OFFICEAGENT_TEST_S3_BUCKET</c>, with optional <c>OFFICEAGENT_TEST_S3_REGION</c>,
/// <c>OFFICEAGENT_TEST_S3_SERVICE_URL</c> and <c>OFFICEAGENT_TEST_S3_FORCE_PATH_STYLE</c>.
/// Credentials come from the AWS SDK's default chain.
/// </remarks>
public sealed class ObjectStoreIntegrationTests
{
    [EnvironmentFact("OFFICEAGENT_TEST_AZUREBLOB_CONTAINER",
        "OFFICEAGENT_TEST_AZUREBLOB_SERVICE_URI|OFFICEAGENT_TEST_AZUREBLOB_CONNECTION_STRING")]
    public async Task Azure_blob_round_trip_against_the_real_service()
    {
        var prefix = $"officeagent-it/{Guid.NewGuid():N}/";
        var options = new AzureBlobDocumentProviderOptions
        {
            ConnectionId = "it",
            ContainerName = EnvironmentFactAttribute.Get("OFFICEAGENT_TEST_AZUREBLOB_CONTAINER")!,
            ServiceUri = EnvironmentFactAttribute.Get("OFFICEAGENT_TEST_AZUREBLOB_SERVICE_URI") ?? string.Empty,
            ConnectionString = EnvironmentFactAttribute.Get("OFFICEAGENT_TEST_AZUREBLOB_CONNECTION_STRING") ?? string.Empty,
            Prefix = prefix
        };
        var provider = new AzureBlobDocumentProvider(options);
        try
        {
            await RoundTripAsync(provider);
        }
        finally
        {
            var container = string.IsNullOrEmpty(options.ConnectionString)
                ? new BlobContainerClient(new Uri(options.ServiceUri.TrimEnd('/') + "/" + options.ContainerName), new DefaultAzureCredential())
                : new BlobContainerClient(options.ConnectionString, options.ContainerName);
            await foreach (var blob in container.GetBlobsAsync(
                               Azure.Storage.Blobs.Models.BlobTraits.None, Azure.Storage.Blobs.Models.BlobStates.None, prefix, CancellationToken.None))
                await container.DeleteBlobIfExistsAsync(blob.Name);
        }
    }

    [EnvironmentFact("OFFICEAGENT_TEST_S3_BUCKET")]
    public async Task Amazon_s3_round_trip_against_the_real_service()
    {
        var prefix = $"officeagent-it/{Guid.NewGuid():N}/";
        var options = new AmazonS3DocumentProviderOptions
        {
            ConnectionId = "it",
            BucketName = EnvironmentFactAttribute.Get("OFFICEAGENT_TEST_S3_BUCKET")!,
            Region = EnvironmentFactAttribute.Get("OFFICEAGENT_TEST_S3_REGION") ?? string.Empty,
            ServiceUrl = EnvironmentFactAttribute.Get("OFFICEAGENT_TEST_S3_SERVICE_URL") ?? string.Empty,
            ForcePathStyle = EnvironmentFactAttribute.Flag("OFFICEAGENT_TEST_S3_FORCE_PATH_STYLE"),
            KeyPrefix = prefix
        };
        var config = new AmazonS3Config { ForcePathStyle = options.ForcePathStyle };
        if (options.ServiceUrl.Length > 0) config.ServiceURL = options.ServiceUrl;
        else if (options.Region.Length > 0) config.RegionEndpoint = Amazon.RegionEndpoint.GetBySystemName(options.Region);
        using var client = new AmazonS3Client(config);
        try
        {
            await RoundTripAsync(new AmazonS3DocumentProvider(options, client));
        }
        finally
        {
            var listed = await client.ListObjectsV2Async(new ListObjectsV2Request { BucketName = options.BucketName, Prefix = prefix });
            foreach (var item in listed.S3Objects ?? new List<S3Object>())
                await client.DeleteObjectAsync(options.BucketName, item.Key);
        }
    }

    /// <summary>Create, register, open, each save mode, a conflict and a refused overwrite.</summary>
    private static async Task RoundTripAsync(IDocumentCreatingProvider provider)
    {
        var created = await provider.CreateAsync("contract.docx", new MemoryStream(DocxFactory.Contract()));
        var duplicate = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.CreateAsync("contract.docx", new MemoryStream(DocxFactory.Contract())));
        Assert.Equal(ProviderErrorCode.AlreadyExists, duplicate.Code);

        var registered = await provider.RegisterAsync("contract.docx");
        using (var opened = await provider.OpenReadAsync(registered))
            Assert.Equal(created.Version, opened.Reference.Version);

        var v2 = await provider.SaveAsync(registered, new MemoryStream(DocxFactory.Contract()),
            new SaveDocumentOptions { Mode = SaveMode.NewVersion });
        Assert.Equal("contract.v2.docx", v2.Name);

        var replaced = await provider.SaveAsync(registered, new MemoryStream(DocxFactory.Contract()),
            new SaveDocumentOptions { Mode = SaveMode.Replace });
        Assert.NotEqual(registered.Version, replaced.Version);

        var stale = await Assert.ThrowsAsync<DocumentVersionConflictException>(() => provider.SaveAsync(
            registered, new MemoryStream(DocxFactory.Contract()), new SaveDocumentOptions { Mode = SaveMode.Replace }));
        Assert.Equal(ProviderErrorCode.VersionConflict, stale.Code);

        var escaped = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.RegisterAsync("../contract.docx"));
        Assert.Equal(ProviderErrorCode.AccessDenied, escaped.Code);

        await provider.RemoveAsync(replaced);
    }
}
