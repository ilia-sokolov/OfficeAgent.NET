using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using OfficeAgent.Core.DocumentProviders;

namespace OfficeAgent.Core;

/// <summary>
/// DI entry point. Hosts register <c>AddOfficeAgent</c> and contribute format
/// modules (and optionally handle resolvers) via <see cref="IServiceCollection"/>.
/// An <see cref="ILoggerFactory"/> registered in the container is picked up
/// automatically; otherwise logging is disabled.
/// </summary>
public static class ServiceCollectionExtensions
{
    /// <summary>
    /// Registers <see cref="IDocumentService"/> and <see cref="OfficeAgentClient"/>.
    /// Callers contribute <see cref="IFormatModule"/> (and optionally <see cref="IHandleResolver"/>)
    /// implementations via <see cref="ServiceLifetime.Singleton"/> registrations.
    /// </summary>
    public static IServiceCollection AddOfficeAgent(this IServiceCollection services)
    {
        services.AddSingleton<IDocumentService>(sp =>
        {
            var modules = sp.GetServices<IFormatModule>().ToArray();
            if (modules.Length == 0)
                throw new InvalidOperationException(
                    "OfficeAgent requires at least one IFormatModule. Call AddWordFormat() " +
                    "(or another format module's registration extension) before AddOfficeAgent().");
            var resolvers = sp.GetServices<IHandleResolver>();
            var resolverList = resolvers.Any() ? resolvers : DefaultHandleResolver.All;
            var loggerFactory = sp.GetService<ILoggerFactory>() ?? NullLoggerFactory.Instance;
            return new OfficeAgentEngine(modules, resolverList, loggerFactory);
        });
        services.AddSingleton<DocumentProviderRegistry>();
        services.AddSingleton(sp => new OfficeAgentClient(
            sp.GetRequiredService<IDocumentService>(),
            sp.GetRequiredService<DocumentProviderRegistry>(),
            sp.GetService<ILoggerFactory>() ?? NullLoggerFactory.Instance));
        return services;
    }
}
