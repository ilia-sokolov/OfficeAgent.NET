using System.Diagnostics.CodeAnalysis;
using OfficeAgent.Abstractions;

namespace OfficeAgent.Core;

/// <summary>
/// Engine surface for the inspect → find → validate → apply flow. Composed by
/// <see cref="OfficeAgentClient"/>; consumers typically use the client rather
/// than implementing this interface directly. Async members observe the
/// supplied <see cref="System.Threading.CancellationToken"/> at operation
/// boundaries.
/// </summary>
[Experimental(EngineExtensibility.DiagnosticId, UrlFormat = EngineExtensibility.UrlFormat)]
public interface IDocumentService
{
    /// <summary>Reads a document's structure, text and anchors without changing it.</summary>
    InspectResult Inspect(DocumentHandle handle, InspectOptions options);

    /// <summary>Finds text in a document and returns a content-verified anchor for each match.</summary>
    IReadOnlyList<FindHit> Find(DocumentHandle handle, FindQuery query);

    /// <summary>Validates a plan against a document and reports the changes it would make, writing nothing.</summary>
    ChangeReport Validate(DocumentHandle handle, DocumentPlan plan);

    /// <summary>Applies a plan as the options direct, all operations or none.</summary>
    ApplyResult Apply(DocumentHandle handle, DocumentPlan plan, ApplyOptions options);

    /// <summary>Reads a document's structure, text and anchors without changing it.</summary>
    Task<InspectResult> InspectAsync(DocumentHandle handle, InspectOptions options, CancellationToken cancellationToken = default);

    /// <summary>Finds text in a document and returns a content-verified anchor for each match.</summary>
    Task<IReadOnlyList<FindHit>> FindAsync(DocumentHandle handle, FindQuery query, CancellationToken cancellationToken = default);

    /// <summary>Validates a plan against a document and reports the changes it would make, writing nothing.</summary>
    Task<ChangeReport> ValidateAsync(DocumentHandle handle, DocumentPlan plan, CancellationToken cancellationToken = default);

    /// <summary>Applies a plan as the options direct, all operations or none.</summary>
    Task<ApplyResult> ApplyAsync(DocumentHandle handle, DocumentPlan plan, ApplyOptions options, CancellationToken cancellationToken = default);
}
