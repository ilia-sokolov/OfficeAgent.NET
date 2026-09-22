using System.Text;
using System.Text.Json;
using OfficeAgent.Abstractions;
using OfficeAgent.Core.DocumentProviders;

namespace OfficeAgent.Core;

public sealed partial class OfficeAgentClient
{
    /// <summary>Host-controlled assembly limits. Tool arguments cannot override these ceilings.</summary>
    public DocumentMergeLimits MergeLimits { get; init; } = new();

    /// <summary>Reads provider sources and previews an ordered Word assembly without saving.</summary>
    public async Task<DocumentMergePreview> PreviewMergeAsync(DocumentMergeRequest request,
        CancellationToken cancellationToken = default)
    {
        if (request is null) throw new ArgumentNullException(nameof(request));
        if (request.Sources is null || request.Options is null)
            throw new ArgumentException("Merge sources and options are required.", nameof(request));
        CheckMergeCount(request.Sources.Count);
        var (bytes, references) = await ReadMergeSources(request.Sources, cancellationToken).ConfigureAwait(false);
        return PreviewMerge(bytes, references, request.Options, cancellationToken);
    }

    /// <summary>Previews an ordered in-memory assembly without saving output.</summary>
    public DocumentMergePreview PreviewMerge(IReadOnlyList<byte[]> sources,
        DocumentMergeOptions? options = null, CancellationToken cancellationToken = default) =>
        PreviewMerge(sources, null, options ?? new(), cancellationToken);

    private DocumentMergePreview PreviewMerge(IReadOnlyList<byte[]> sources,
        IReadOnlyList<DocumentReference>? references, DocumentMergeOptions options, CancellationToken cancellationToken)
    {
        CheckMergeBytes(sources);
        var candidate = AssembleMerge(sources, options, cancellationToken);
        if (candidate.Content is null || candidate.Diagnostics.Count != 0)
            return new() { Sources = candidate.Sources, Diagnostics = candidate.Diagnostics };
        var inputs = sources.Select((bytes, index) => new DocumentMergeInput
        {
            Index = index,
            Sha256 = Sha256(bytes),
            Document = references?[index]
        }).ToArray();
        return new()
        {
            Sources = candidate.Sources,
            Plan = new DocumentMergePlan
            {
                Inputs = inputs,
                Options = options,
                PlanSha256 = MergePlanHash(inputs, options)
            }
        };
    }

    /// <summary>Revalidates all exact inputs and creates a new .docx through the destination provider.</summary>
    public async Task<DocumentMergeResult> CommitMergeAsync(DocumentMergePlan plan,
        string destinationConnectionId, string outputName, CancellationToken cancellationToken = default)
    {
        ValidateMergePlan(plan);
        if (!string.Equals(Path.GetExtension(outputName), ".docx", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("Assembly output must have a .docx extension.", nameof(outputName));
        var provider = _providers.ResolveConnection(destinationConnectionId);
        if (provider is not IDocumentCreatingProvider creator)
            throw new DocumentProviderException(ProviderErrorCode.ConfigurationError,
                "The destination does not support creating documents.", provider.Provider, provider.ConnectionId, null);
        var references = plan.Inputs.Select(input => input.Document
            ?? throw new ArgumentException("Provider merge requires provider input references.", nameof(plan))).ToArray();
        var (bytes, _) = await ReadMergeSources(references, cancellationToken).ConfigureAwait(false);
        var result = CommitMerge(plan, bytes, cancellationToken);
        if (!result.Committed) return result;
        cancellationToken.ThrowIfCancellationRequested();
        using var output = new MemoryStream(result.Content!, writable: false);
        // A failure after the create began is reported as an uncertain write, never as no write.
        var document = await StorageWrite.RunAsync(provider, itemId: null, outputName, result.Content!,
            token => creator.CreateAsync(outputName, output, token), cancellationToken).ConfigureAwait(false);
        return new()
        {
            Committed = true,
            Document = document,
            Receipt = new DocumentMergeReceipt
            {
                Inputs = result.Receipt!.Inputs,
                PlanSha256 = result.Receipt.PlanSha256,
                OutputSha256 = result.Receipt.OutputSha256,
                TimestampUtc = result.Receipt.TimestampUtc,
                Actor = result.Receipt.Actor,
                OutputDocument = document
            }
        };
    }

    /// <summary>Produces assembled bytes only after validating the plan and every input hash.</summary>
    public DocumentMergeResult CommitMerge(DocumentMergePlan plan, IReadOnlyList<byte[]> sources,
        CancellationToken cancellationToken = default)
    {
        ValidateMergePlan(plan);
        CheckMergeBytes(sources);
        cancellationToken.ThrowIfCancellationRequested();
        if (sources.Count != plan.Inputs.Count || sources.Where((bytes, index) =>
            !string.Equals(Sha256(bytes), plan.Inputs[index].Sha256, StringComparison.Ordinal)).Any())
            return new() { Diagnostics = new[] { Diagnostic(AssemblyDiagnosticCodes.StaleMergeSource, "An assembly input changed; preview again.", "sources") } };
        var candidate = AssembleMerge(sources, plan.Options, cancellationToken);
        if (candidate.Content is null || candidate.Diagnostics.Count != 0)
            return new() { Diagnostics = candidate.Diagnostics };
        return new()
        {
            Committed = true,
            Content = candidate.Content,
            Receipt = MergeReceipt(plan, candidate.Content, null)
        };
    }

    private DocumentAssemblyCandidate AssembleMerge(IReadOnlyList<byte[]> sources, DocumentMergeOptions options,
        CancellationToken cancellationToken) =>
        (_blankDocumentFactories.OfType<IDocumentAssembler>().SingleOrDefault()
            ?? throw new NotSupportedException("Register a Word assembly-capable format module."))
        .Assemble(sources, options, MergeLimits, cancellationToken);

    private DocumentMergeReceipt MergeReceipt(DocumentMergePlan plan, byte[] bytes, DocumentReference? document) => new()
    {
        Inputs = plan.Inputs,
        PlanSha256 = plan.PlanSha256,
        OutputSha256 = Sha256(bytes),
        TimestampUtc = DateTimeOffset.UtcNow,
        Actor = ResolveActor(null),
        OutputDocument = document
    };

    private void ValidateMergePlan(DocumentMergePlan plan)
    {
        if (plan is null) throw new ArgumentNullException(nameof(plan));
        if (plan.Inputs is null || plan.Options is null)
            throw new ArgumentException("Merge plan inputs and options are required.", nameof(plan));
        CheckMergeCount(plan.Inputs.Count);
        if (plan.Version != DocumentMergePlan.CurrentVersion ||
            plan.Inputs.Any(input => input is null || input.Sha256.Length != 64) ||
            plan.Inputs.Where((input, index) => input.Index != index).Any() ||
            plan.PlanSha256 != MergePlanHash(plan.Inputs, plan.Options))
            throw new ArgumentException(
                $"The merge plan is invalid or changed. Version must be \"{DocumentMergePlan.CurrentVersion}\"; preview again.",
                nameof(plan));
    }

    private static string MergePlanHash(IReadOnlyList<DocumentMergeInput> inputs, DocumentMergeOptions options) =>
        Sha256(Encoding.UTF8.GetBytes(JsonSerializer.Serialize(new
        {
            Version = DocumentMergePlan.CurrentVersion,
            Inputs = inputs,
            Options = options
        })));

    private void CheckMergeCount(int count)
    {
        if (MergeLimits.MaximumSources < 2 || MergeLimits.MaximumInputBytes <= 0 ||
            MergeLimits.MaximumTotalInputBytes <= 0 || MergeLimits.MaximumExpandedBytes <= 0 ||
            MergeLimits.MaximumParts <= 0 || MergeLimits.MaximumOutputBytes <= 0)
            throw new ArgumentException("Merge limits must be positive and allow at least two sources.");
        if (count < 2 || count > MergeLimits.MaximumSources)
            throw new ArgumentException($"Merge requires 2 to {MergeLimits.MaximumSources} ordered sources.");
    }

    private void CheckMergeBytes(IReadOnlyList<byte[]> sources)
    {
        CheckMergeCount(sources.Count);
        if (sources.Any(bytes => bytes is null || bytes.LongLength > MergeLimits.MaximumInputBytes) ||
            sources.Sum(bytes => bytes.LongLength) > MergeLimits.MaximumTotalInputBytes)
            throw new ArgumentException("Assembly compressed input limit exceeded.", nameof(sources));
    }

    private async Task<(IReadOnlyList<byte[]> Bytes, IReadOnlyList<DocumentReference> References)> ReadMergeSources(
        IReadOnlyList<DocumentReference> sources, CancellationToken cancellationToken)
    {
        var bytes = new List<byte[]>();
        var references = new List<DocumentReference>();
        long total = 0;
        foreach (var source in sources)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var provider = string.IsNullOrEmpty(source.Provider)
                ? _providers.ResolveConnection(source.ConnectionId) : _providers.Resolve(source);
            var reference = DocumentReference.For(provider.Provider, source.ConnectionId, source.ItemId);
            using var content = await provider.OpenReadAsync(reference, cancellationToken).ConfigureAwait(false);
            var data = await ReadBoundedAsync(content.Stream,
                Math.Min(MergeLimits.MaximumInputBytes, MergeLimits.MaximumTotalInputBytes - total), cancellationToken).ConfigureAwait(false);
            total += data.LongLength;
            bytes.Add(data);
            references.Add(content.Reference);
        }
        return (bytes, references);
    }
}
