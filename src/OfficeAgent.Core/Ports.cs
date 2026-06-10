using OfficeAgent.Abstractions;
using DocFormat = OfficeAgent.Abstractions.DocumentFormat;

namespace OfficeAgent.Core;

/// <summary>
/// Defines the contract implemented by a document format module.
/// </summary>
public interface IFormatModule
{
    /// <summary>Gets the document format served by the module.</summary>
    DocFormat Format { get; }

    /// <summary>Returns whether the module can process the open package.</summary>
    /// <param name="package">The open package.</param>
    /// <returns><see langword="true"/> when the package can be processed; otherwise, <see langword="false"/>.</returns>
    bool CanHandle(IOpenXmlPackage package);

    /// <summary>Inspects the package and returns an anchored document model.</summary>
    /// <param name="package">The open package.</param>
    /// <param name="options">Inspection options.</param>
    /// <returns>The inspection result.</returns>
    InspectResult Inspect(IOpenXmlPackage package, InspectOptions options);

    /// <summary>Finds text in the package.</summary>
    /// <param name="package">The open package.</param>
    /// <param name="query">The find query.</param>
    /// <returns>Matching hits with content-verified anchors.</returns>
    IReadOnlyList<FindHit> Find(IOpenXmlPackage package, FindQuery query);

    /// <summary>Gets operation handlers registered by the module.</summary>
    IReadOnlyList<IOperationHandler> Handlers { get; }

    /// <summary>
    /// Ensures the package carries stable identifiers for the entities a plan may address
    /// (e.g. assigns <c>w14:paraId</c> to paragraphs that lack one) and returns a map from
    /// the positional id surfaced by <see cref="Inspect"/> (e.g. <c>auto-0007</c>) to the
    /// now-stable id (e.g. <c>w14:…</c>). Handlers translate anchor ids through this map via
    /// <see cref="ApplyContext.ResolveAlias"/>, so an earlier offset-shifting op cannot
    /// redirect a later op's target. Mutates the package; idempotent.
    /// </summary>
    IReadOnlyDictionary<string, string> Stabilize(IOpenXmlPackage package);
}

/// <summary>
/// Defines preview and apply behavior for one or more plan operations.
/// </summary>
public interface IOperationHandler
{
    /// <summary>Returns whether the handler supports the operation.</summary>
    /// <param name="operation">The plan operation.</param>
    /// <returns><see langword="true"/> when the operation is supported; otherwise, <see langword="false"/>.</returns>
    bool CanHandle(PlanOperation operation);

    /// <summary>Previews the operation without mutating the package.</summary>
    /// <param name="context">The apply context.</param>
    /// <param name="operation">The operation to preview.</param>
    /// <returns>The preview result.</returns>
    OperationPreview Preview(ApplyContext context, PlanOperation operation);

    /// <summary>Applies the operation to the package.</summary>
    /// <param name="context">The apply context.</param>
    /// <param name="operation">The operation to apply.</param>
    void Apply(ApplyContext context, PlanOperation operation);
}

/// <summary>
/// Contains the result of previewing one operation.
/// </summary>
public sealed class OperationPreview
{
    /// <summary>Gets the proposed change when preview succeeds.</summary>
    public ProposedChange? Change { get; init; }

    /// <summary>Gets the validation error when preview fails.</summary>
    public ValidationError? Error { get; init; }

    /// <summary>Creates a successful operation preview.</summary>
    /// <param name="change">The proposed change.</param>
    /// <returns>A successful preview.</returns>
    public static OperationPreview Ok(ProposedChange change) => new() { Change = change };

    /// <summary>Creates a failed operation preview.</summary>
    /// <param name="error">The validation error.</param>
    /// <returns>A failed preview.</returns>
    public static OperationPreview Fail(ValidationError error) => new() { Error = error };
}

/// <summary>
/// Carries the open package, the inspection captured when the context was created, and
/// the anchor-alias map produced by <see cref="IFormatModule.Stabilize"/>.
/// </summary>
public sealed class ApplyContext
{
    private readonly IReadOnlyDictionary<string, string> _aliases;

    /// <summary>Gets the open package being edited.</summary>
    public IOpenXmlPackage Package { get; }

    /// <summary>Gets the inspection captured when the context was created.</summary>
    public InspectResult Inspection { get; }

    /// <summary>Initializes a new instance of the <see cref="ApplyContext"/> class.</summary>
    /// <param name="package">The open package.</param>
    /// <param name="inspection">The inspection captured for the package.</param>
    /// <param name="aliases">The positional-to-stable anchor id map, or <see langword="null"/> for none.</param>
    public ApplyContext(
        IOpenXmlPackage package,
        InspectResult inspection,
        IReadOnlyDictionary<string, string>? aliases = null)
    {
        Package = package;
        Inspection = inspection;
        _aliases = aliases ?? EmptyAliases;
    }

    /// <summary>
    /// Translates a plan anchor id through the stabilization alias map. Returns the input
    /// unchanged when no alias exists, so anchors that already use stable ids pass through.
    /// </summary>
    public string ResolveAlias(string anchorId) =>
        _aliases.TryGetValue(anchorId, out var stable) ? stable : anchorId;

    private static readonly IReadOnlyDictionary<string, string> EmptyAliases =
        new Dictionary<string, string>(0);
}

/// <summary>
/// Resolves document handles to readable streams.
/// </summary>
public interface IHandleResolver
{
    /// <summary>Returns whether this resolver can open the handle.</summary>
    /// <param name="handle">The document handle.</param>
    /// <returns><see langword="true"/> when the handle can be resolved; otherwise, <see langword="false"/>.</returns>
    bool CanResolve(DocumentHandle handle);

    /// <summary>Opens the document represented by the handle.</summary>
    /// <param name="handle">The document handle.</param>
    /// <returns>A readable document stream.</returns>
    Stream Resolve(DocumentHandle handle);
}
