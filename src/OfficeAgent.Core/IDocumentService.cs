using OfficeAgent.Abstractions;

namespace OfficeAgent.Core;

/// <summary>
/// Engine surface for the inspect → find → validate → apply flow. Composed by
/// <see cref="OfficeAgentClient"/>; consumers typically use the client rather
/// than implementing this interface directly. Async members observe the
/// supplied <see cref="System.Threading.CancellationToken"/> at operation
/// boundaries.
/// </summary>
public interface IDocumentService
{
    InspectResult Inspect(DocumentHandle handle, InspectOptions options);

    IReadOnlyList<FindHit> Find(DocumentHandle handle, FindQuery query);

    ChangeReport Validate(DocumentHandle handle, DocumentPlan plan);

    ApplyResult Apply(DocumentHandle handle, DocumentPlan plan, ApplyOptions options);

    Task<InspectResult> InspectAsync(DocumentHandle handle, InspectOptions options, CancellationToken cancellationToken = default);

    Task<IReadOnlyList<FindHit>> FindAsync(DocumentHandle handle, FindQuery query, CancellationToken cancellationToken = default);

    Task<ChangeReport> ValidateAsync(DocumentHandle handle, DocumentPlan plan, CancellationToken cancellationToken = default);

    Task<ApplyResult> ApplyAsync(DocumentHandle handle, DocumentPlan plan, ApplyOptions options, CancellationToken cancellationToken = default);
}
