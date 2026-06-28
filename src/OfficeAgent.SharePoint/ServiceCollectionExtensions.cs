using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.Core.DocumentProviders;

namespace OfficeAgent.SharePoint;

/// <summary>Dependency-injection registrations for the SharePoint document provider.</summary>
public static class ServiceCollectionExtensions
{
    /// <summary>
    /// Registers a SharePoint document provider. Documents are registered by a
    /// SharePoint/OneDrive URL or a <c>driveId/itemId</c> pair, so the connection is
    /// not pinned to a single drive. The token provider and (optionally) a durable
    /// <see cref="ISharePointRegistrationStore"/> are resolved from the container;
    /// register them before calling this method.
    /// </summary>
    public static IServiceCollection AddSharePointDocumentProvider(
        this IServiceCollection services,
        string connectionId,
        Action<SharePointDocumentProviderOptions>? configure = null)
    {
        var options = new SharePointDocumentProviderOptions
        {
            ConnectionId = connectionId
        };
        configure?.Invoke(options);
        services.AddSingleton<IDocumentProvider>(sp => new SharePointDocumentProvider(
            options,
            sp.GetRequiredService<HttpClient>(),
            sp.GetRequiredService<IAccessTokenProvider>(),
            sp.GetService<ISharePointRegistrationStore>()));
        return services;
    }

    /// <summary>
    /// Registers an <see cref="OnBehalfOfAccessTokenProvider"/> as the connection's
    /// token provider, so the SharePoint provider acts as the signed-in user. The host
    /// must capture each caller's access token into <see cref="GraphUserContext"/> for
    /// the duration of the request (the OfficeAgent MCP HTTP server does this from the
    /// inbound <c>Authorization</c> header).
    /// </summary>
    public static IServiceCollection AddSharePointOnBehalfOfAuthentication(
        this IServiceCollection services,
        Action<OnBehalfOfOptions> configure)
    {
        if (configure is null) throw new ArgumentNullException(nameof(configure));
        var options = new OnBehalfOfOptions();
        configure(options);
        services.AddSingleton<IAccessTokenProvider>(sp =>
            new OnBehalfOfAccessTokenProvider(options, sp.GetRequiredService<HttpClient>()));
        return services;
    }
}
