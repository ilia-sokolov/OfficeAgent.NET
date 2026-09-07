using Microsoft.Extensions.DependencyInjection;

namespace OfficeAgent.Core.DocumentProviders;

/// <summary>Dependency-injection registrations for document providers.</summary>
public static class ServiceCollectionExtensions
{
    /// <summary>Registers a rooted local-filesystem document provider.</summary>
    public static IServiceCollection AddFileSystemDocumentProvider(
        this IServiceCollection services,
        string connectionId,
        string rootPath,
        Action<FileSystemDocumentProviderOptions>? configure = null)
    {
        var options = new FileSystemDocumentProviderOptions
        {
            ConnectionId = connectionId,
            RootPath = rootPath
        };
        configure?.Invoke(options);
        services.AddSingleton<IDocumentProvider>(new FileSystemDocumentProvider(options));
        return services;
    }

    /// <summary>
    /// Adds a connection whose documents live in this process's memory for as long as it
    /// runs, so an agent can create and edit documents with no storage configured.
    /// </summary>
    /// <remarks>
    /// The instance is registered as itself as well as as an <see cref="IDocumentProvider"/>,
    /// because a host needs the concrete type to put content in and take it back out -
    /// there is no path or URL to name a document by.
    /// </remarks>
    public static IServiceCollection AddMemoryDocumentProvider(
        this IServiceCollection services,
        string connectionId,
        Action<MemoryDocumentProviderOptions>? configure = null)
    {
        var options = new MemoryDocumentProviderOptions();
        configure?.Invoke(options);

        var provider = new MemoryDocumentProvider(connectionId, options);
        services.AddSingleton(provider);
        services.AddSingleton<IDocumentProvider>(provider);
        return services;
    }
}
