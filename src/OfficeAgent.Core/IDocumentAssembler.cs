using OfficeAgent.Abstractions;

namespace OfficeAgent.Core;

/// <summary>Optional format-module capability for full document assembly.</summary>
public interface IDocumentAssembler
{
    /// <summary>Builds and validates one assembled candidate from ordered exact source bytes.</summary>
    DocumentAssemblyCandidate Assemble(IReadOnlyList<byte[]> sources, DocumentMergeOptions options,
        DocumentMergeLimits limits, CancellationToken cancellationToken);
}

/// <summary>Validated candidate bytes, source coverage, or blocking diagnostics.</summary>
public sealed class DocumentAssemblyCandidate
{
    /// <summary>Gets valid assembled bytes, or null when diagnostics block assembly.</summary>
    public byte[]? Content { get; init; }
    /// <summary>Gets source coverage and remapping reports.</summary>
    public IReadOnlyList<DocumentMergeSourceReport> Sources { get; init; } = Array.Empty<DocumentMergeSourceReport>();
    /// <summary>Gets blocking diagnostics.</summary>
    public IReadOnlyList<WorkflowDiagnostic> Diagnostics { get; init; } = Array.Empty<WorkflowDiagnostic>();
}
