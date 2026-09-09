using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;
using OfficeAgent.Core;

namespace OfficeAgent.Excel;

/// <summary>Registers Excel support with OfficeAgent.</summary>
public static class ServiceCollectionExtensions
{
    /// <summary>Adds the Excel format module.</summary>
    public static IServiceCollection AddExcelFormat(this IServiceCollection services)
    {
        services.TryAddSingleton(TimeProvider.System);
        services.AddSingleton<IFormatModule>(sp =>
            new ExcelModule(sp.GetRequiredService<TimeProvider>()));
        return services;
    }
}
