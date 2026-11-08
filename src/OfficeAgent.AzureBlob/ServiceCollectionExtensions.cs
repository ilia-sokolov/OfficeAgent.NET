using Azure.Storage.Blobs;
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.Core.DocumentProviders;

namespace OfficeAgent.AzureBlob;

/// <summary>Dependency-injection registrations for the Azure Blob Storage document provider.</summary>
public static class ServiceCollectionExtensions
{
    /// <summary>
    /// Registers an Azure Blob provider connection over one container and an optional prefix.
    /// Set <see cref="AzureBlobDocumentProviderOptions.ServiceUri"/> to authenticate with
    /// <c>DefaultAzureCredential</c> (managed identity in Azure), or
    /// <see cref="AzureBlobDocumentProviderOptions.ConnectionString"/> from a secret store.
    /// </summary>
    public static IServiceCollection AddAzureBlobDocumentProvider(
        this IServiceCollection services,
        string connectionId,
        Action<AzureBlobDocumentProviderOptions>? configure = null)
    {
        if (services is null) throw new ArgumentNullException(nameof(services));
        var options = Configure(connectionId, configure);
        services.AddSingleton<IDocumentProvider>(_ => new AzureBlobDocumentProvider(options));
        return services;
    }

    /// <summary>
    /// Registers an Azure Blob provider connection over a container client the host builds - for
    /// a custom <c>TokenCredential</c>, a SAS, or the Azurite emulator.
    /// </summary>
    public static IServiceCollection AddAzureBlobDocumentProvider(
        this IServiceCollection services,
        string connectionId,
        Func<IServiceProvider, BlobContainerClient> containerFactory,
        Action<AzureBlobDocumentProviderOptions>? configure = null)
    {
        if (services is null) throw new ArgumentNullException(nameof(services));
        if (containerFactory is null) throw new ArgumentNullException(nameof(containerFactory));
        var options = Configure(connectionId, configure);
        services.AddSingleton<IDocumentProvider>(sp => new AzureBlobDocumentProvider(options, containerFactory(sp)));
        return services;
    }

    private static AzureBlobDocumentProviderOptions Configure(
        string connectionId, Action<AzureBlobDocumentProviderOptions>? configure)
    {
        var options = new AzureBlobDocumentProviderOptions { ConnectionId = connectionId };
        configure?.Invoke(options);
        return options;
    }
}
