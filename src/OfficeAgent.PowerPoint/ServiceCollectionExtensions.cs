using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;
using OfficeAgent.Core;

namespace OfficeAgent.PowerPoint;

/// <summary>
/// DI registration for the PowerPoint format module. Hosts call <c>AddPowerPointFormat</c>
/// and may additionally register <see cref="IOperationHandler"/> and
/// <see cref="IPowerPointNodeProvider"/> implementations to extend the module to new verbs
/// and node kinds, plus a <see cref="TimeProvider"/> for deterministic timestamps.
/// </summary>
public static class ServiceCollectionExtensions
{
    /// <summary>
    /// Registers a <see cref="PowerPointModule"/> as an <see cref="IFormatModule"/>,
    /// composing it from the container's <see cref="TimeProvider"/> (defaulting to
    /// <see cref="TimeProvider.System"/>) and any contributed handlers and node providers.
    /// Sits alongside <c>AddWordFormat</c>: a host that registers both serves either
    /// format, and the engine routes each document to the module that can handle it.
    /// </summary>
    /// <remarks>
    /// Contributed handlers are resolved as <see cref="IPowerPointOperationHandler"/>
    /// rather than as bare <see cref="IOperationHandler"/>. A bare registration says
    /// nothing about which format it understands, so every module in the container would
    /// pick it up - a handler written for Word would then be offered a deck.
    /// </remarks>
    public static IServiceCollection AddPowerPointFormat(this IServiceCollection services)
    {
        services.TryAddSingleton(TimeProvider.System);
        services.AddSingleton<IFormatModule>(sp => new PowerPointModule(
            sp.GetRequiredService<TimeProvider>(),
            sp.GetServices<IPowerPointOperationHandler>(),
            sp.GetServices<IPowerPointNodeProvider>()));
        return services;
    }
}
