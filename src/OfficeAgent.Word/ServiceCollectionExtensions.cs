using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;
using OfficeAgent.Core;

namespace OfficeAgent.Word;

/// <summary>
/// DI registration for the Word format module. Hosts call <c>AddWordFormat</c> and may
/// additionally register <see cref="IOperationHandler"/> and <see cref="IWordNodeProvider"/>
/// implementations to extend the module to new verbs and node kinds, plus a
/// <see cref="TimeProvider"/> for deterministic timestamps.
/// </summary>
public static class ServiceCollectionExtensions
{
    /// <summary>
    /// Registers a <see cref="WordModule"/> as an <see cref="IFormatModule"/>, composing it
    /// from the container's <see cref="TimeProvider"/> (defaulting to <see cref="TimeProvider.System"/>)
    /// and any contributed handlers and node providers.
    /// </summary>
    public static IServiceCollection AddWordFormat(this IServiceCollection services)
    {
        services.TryAddSingleton(TimeProvider.System);
        services.AddSingleton<IFormatModule>(sp => new WordModule(
            sp.GetRequiredService<TimeProvider>(),
            sp.GetServices<IOperationHandler>(),
            sp.GetServices<IWordNodeProvider>()));
        return services;
    }
}
