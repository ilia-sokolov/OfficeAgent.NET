using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.Core.DocumentProviders;

namespace OfficeAgent.IManage;

/// <summary>Dependency-injection registrations for the iManage Work document provider.</summary>
public static class ServiceCollectionExtensions
{
    /// <summary>
    /// Registers an iManage Work document provider. Documents are registered by their iManage
    /// id (<c>LIBRARY!number.version</c>) or an iManage Work link. The token provider and
    /// (optionally) a durable <see cref="IIManageRegistrationStore"/> are resolved from the
    /// container; register them before calling this method. The provider makes its calls through
    /// its own <see cref="HttpClient"/>, which does not follow redirects.
    /// </summary>
    public static IServiceCollection AddIManageDocumentProvider(
        this IServiceCollection services,
        string connectionId,
        Action<IManageDocumentProviderOptions> configure)
    {
        if (configure is null) throw new ArgumentNullException(nameof(configure));
        var options = new IManageDocumentProviderOptions { ConnectionId = connectionId };
        configure(options);
        services.AddSingleton<IDocumentProvider>(sp => new IManageDocumentProvider(
            options,
            IManageDocumentProvider.CreateHttpClient(),
            sp.GetRequiredService<IIManageAccessTokenProvider>(),
            sp.GetService<IIManageRegistrationStore>()));
        return services;
    }

    /// <summary>
    /// Registers an <see cref="IManageOAuthTokenProvider"/> as the iManage token provider. Keep
    /// the secret, password, or refresh token in a secret store or the environment, not source.
    /// </summary>
    public static IServiceCollection AddIManageOAuthAuthentication(
        this IServiceCollection services,
        Action<IManageOAuthOptions> configure)
    {
        if (configure is null) throw new ArgumentNullException(nameof(configure));
        var options = new IManageOAuthOptions();
        configure(options);
        services.AddSingleton<IIManageAccessTokenProvider>(_ =>
            new IManageOAuthTokenProvider(options, IManageDocumentProvider.CreateHttpClient()));
        return services;
    }
}
