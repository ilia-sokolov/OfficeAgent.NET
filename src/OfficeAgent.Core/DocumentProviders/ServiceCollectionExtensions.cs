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
}
