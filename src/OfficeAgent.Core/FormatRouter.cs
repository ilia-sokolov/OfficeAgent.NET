using OfficeAgent.Core;

namespace OfficeAgent.Core;

/// <summary>Dispatches a package to the format module that can handle it.</summary>
internal sealed class FormatRouter
{
    private readonly IReadOnlyList<IFormatModule> _modules;

    public FormatRouter(IEnumerable<IFormatModule> modules) => _modules = modules.ToList();

    public IFormatModule Route(IOpenXmlPackage package) =>
        _modules.FirstOrDefault(m => m.CanHandle(package))
        ?? throw new NotSupportedException($"No format module registered for {package.Format}.");
}
