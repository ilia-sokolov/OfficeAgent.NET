using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.Core.DocumentProviders;

namespace OfficeAgent.GoogleDrive;

/// <summary>Dependency-injection registrations for the Google Drive document provider.</summary>
public static class GoogleDriveServiceCollectionExtensions
{
    /// <summary>
    /// Registers a Google Drive document provider. Documents are registered by a Drive file
    /// id or URL. The <see cref="HttpClient"/>, the <see cref="IGoogleDriveAccessTokenProvider"/>
    /// and (optionally) a durable <see cref="IGoogleDriveRegistrationStore"/> are resolved from
    /// the container; register them before calling this method.
    /// </summary>
    public static IServiceCollection AddGoogleDriveDocumentProvider(
        this IServiceCollection services,
        string connectionId,
        Action<GoogleDriveDocumentProviderOptions>? configure = null)
    {
        if (services is null) throw new ArgumentNullException(nameof(services));
        var options = new GoogleDriveDocumentProviderOptions
        {
            ConnectionId = connectionId
        };
        configure?.Invoke(options);
        services.AddSingleton<IDocumentProvider>(sp => new GoogleDriveDocumentProvider(
            options,
            sp.GetRequiredService<HttpClient>(),
            sp.GetRequiredService<IGoogleDriveAccessTokenProvider>(),
            sp.GetService<IGoogleDriveRegistrationStore>()));
        return services;
    }

    /// <summary>
    /// Registers a <see cref="GoogleServiceAccountAccessTokenProvider"/> as the Drive token
    /// provider. Supply the key from a secret store or environment, never from source.
    /// </summary>
    public static IServiceCollection AddGoogleDriveServiceAccountAuthentication(
        this IServiceCollection services,
        Action<GoogleServiceAccountOptions> configure)
    {
        if (services is null) throw new ArgumentNullException(nameof(services));
        if (configure is null) throw new ArgumentNullException(nameof(configure));
        var options = new GoogleServiceAccountOptions();
        configure(options);
        services.AddSingleton<IGoogleDriveAccessTokenProvider>(sp =>
            new GoogleServiceAccountAccessTokenProvider(options, sp.GetRequiredService<HttpClient>()));
        return services;
    }
}
