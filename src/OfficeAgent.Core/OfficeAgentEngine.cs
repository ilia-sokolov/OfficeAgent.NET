using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using OfficeAgent.Abstractions;

namespace OfficeAgent.Core;

/// <summary>
/// The reference implementation of <see cref="IDocumentService"/>. Composed with
/// the format modules it should serve. Optionally accepts an <see cref="ILoggerFactory"/>
/// for structured logging; uses <see cref="NullLoggerFactory.Instance"/> by default.
/// </summary>
internal sealed class OfficeAgentEngine : IDocumentService
{
    private readonly FlowOrchestrator _flow;

    public OfficeAgentEngine(
        IEnumerable<IFormatModule> modules,
        IEnumerable<IHandleResolver>? resolvers = null,
        ILoggerFactory? loggerFactory = null)
        => _flow = new FlowOrchestrator(
            modules,
            resolvers ?? DefaultHandleResolver.All,
            (loggerFactory ?? NullLoggerFactory.Instance).CreateLogger(OfficeAgentTelemetry.LogCategory));

    public InspectResult Inspect(DocumentHandle handle, InspectOptions options) =>
        _flow.Inspect(handle, options);

    public IReadOnlyList<FindHit> Find(DocumentHandle handle, FindQuery query) =>
        _flow.Find(handle, query);

    public ChangeReport Validate(DocumentHandle handle, DocumentPlan plan) =>
        _flow.Validate(handle, plan);

    public ApplyResult Apply(DocumentHandle handle, DocumentPlan plan, ApplyOptions options) =>
        _flow.Apply(handle, plan, options);

    public Task<InspectResult> InspectAsync(DocumentHandle handle, InspectOptions options, CancellationToken cancellationToken = default) =>
        _flow.InspectAsync(handle, options, cancellationToken);

    public Task<IReadOnlyList<FindHit>> FindAsync(DocumentHandle handle, FindQuery query, CancellationToken cancellationToken = default) =>
        _flow.FindAsync(handle, query, cancellationToken);

    public Task<ChangeReport> ValidateAsync(DocumentHandle handle, DocumentPlan plan, CancellationToken cancellationToken = default) =>
        _flow.ValidateAsync(handle, plan, cancellationToken);

    public Task<ApplyResult> ApplyAsync(DocumentHandle handle, DocumentPlan plan, ApplyOptions options, CancellationToken cancellationToken = default) =>
        _flow.ApplyAsync(handle, plan, options, cancellationToken);
}
