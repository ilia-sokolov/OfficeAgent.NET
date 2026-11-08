using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.Core.DocumentProviders;

namespace OfficeAgent.Box;

/// <summary>Dependency-injection registrations for the Box document provider.</summary>
public static class BoxServiceCollectionExtensions
{
    /// <summary>
    /// Registers a Box document provider. Documents are registered by a Box file id or URL.
    /// The <see cref="HttpClient"/>, the <see cref="IBoxAccessTokenProvider"/> and
    /// (optionally) a durable <see cref="IBoxRegistrationStore"/> are resolved from the
    /// container; register them before calling this method.
    /// </summary>
    public static IServiceCollection AddBoxDocumentProvider(
        this IServiceCollection services,
        string connectionId,
        Action<BoxDocumentProviderOptions>? configure = null)
    {
        if (services is null) throw new ArgumentNullException(nameof(services));
        var options = new BoxDocumentProviderOptions
        {
            ConnectionId = connectionId
        };
        configure?.Invoke(options);
        services.AddSingleton<IDocumentProvider>(sp => new BoxDocumentProvider(
            options,
            sp.GetRequiredService<HttpClient>(),
            sp.GetRequiredService<IBoxAccessTokenProvider>(),
            sp.GetService<IBoxRegistrationStore>()));
        return services;
    }

    /// <summary>
    /// Registers a <see cref="BoxClientCredentialsAccessTokenProvider"/> as the Box token
    /// provider. Supply the client secret from a secret store or environment, never from source.
    /// </summary>
    public static IServiceCollection AddBoxClientCredentialsAuthentication(
        this IServiceCollection services,
        Action<BoxClientCredentialsOptions> configure)
    {
        if (services is null) throw new ArgumentNullException(nameof(services));
        if (configure is null) throw new ArgumentNullException(nameof(configure));
        var options = new BoxClientCredentialsOptions();
        configure(options);
        services.AddSingleton<IBoxAccessTokenProvider>(sp =>
            new BoxClientCredentialsAccessTokenProvider(options, sp.GetRequiredService<HttpClient>()));
        return services;
    }
}
