using Amazon.S3;
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.Core.DocumentProviders;

namespace OfficeAgent.AmazonS3;

/// <summary>Dependency-injection registrations for the Amazon S3 document provider.</summary>
public static class ServiceCollectionExtensions
{
    /// <summary>
    /// Registers an Amazon S3 provider connection over one bucket and an optional key prefix.
    /// Credentials come from the AWS SDK's default chain, such as an instance or task role.
    /// </summary>
    public static IServiceCollection AddAmazonS3DocumentProvider(
        this IServiceCollection services,
        string connectionId,
        Action<AmazonS3DocumentProviderOptions>? configure = null)
    {
        if (services is null) throw new ArgumentNullException(nameof(services));
        var options = Configure(connectionId, configure);
        services.AddSingleton<IDocumentProvider>(_ => new AmazonS3DocumentProvider(options));
        return services;
    }

    /// <summary>
    /// Registers an Amazon S3 provider connection over an S3 client the host builds, for custom
    /// credentials, endpoint, checksum or retry settings.
    /// </summary>
    public static IServiceCollection AddAmazonS3DocumentProvider(
        this IServiceCollection services,
        string connectionId,
        Func<IServiceProvider, IAmazonS3> clientFactory,
        Action<AmazonS3DocumentProviderOptions>? configure = null)
    {
        if (services is null) throw new ArgumentNullException(nameof(services));
        if (clientFactory is null) throw new ArgumentNullException(nameof(clientFactory));
        var options = Configure(connectionId, configure);
        services.AddSingleton<IDocumentProvider>(sp => new AmazonS3DocumentProvider(options, clientFactory(sp)));
        return services;
    }

    private static AmazonS3DocumentProviderOptions Configure(
        string connectionId, Action<AmazonS3DocumentProviderOptions>? configure)
    {
        var options = new AmazonS3DocumentProviderOptions { ConnectionId = connectionId };
        configure?.Invoke(options);
        return options;
    }
}
