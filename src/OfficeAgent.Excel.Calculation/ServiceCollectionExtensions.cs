using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;
using OfficeAgent.Abstractions;

namespace OfficeAgent.Excel.Calculation;

/// <summary>Registers optional formula calculation with OfficeAgent.</summary>
public static class ServiceCollectionExtensions
{
    /// <summary>
    /// Adds the ClosedXML formula evaluator. Calculation then runs on any preview or commit of a
    /// workbook that sets <see cref="ApplyOptions.CalculateFormulas"/>, or on a connection that
    /// turns it on; it stays off everywhere else.
    /// </summary>
    public static IServiceCollection AddExcelFormulaCalculation(this IServiceCollection services)
    {
        services.TryAddSingleton<IFormulaEvaluator, ClosedXmlFormulaEvaluator>();
        return services;
    }
}
