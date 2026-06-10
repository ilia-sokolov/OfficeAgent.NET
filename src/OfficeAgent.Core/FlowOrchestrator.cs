using System.Diagnostics;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using OfficeAgent.Abstractions;

namespace OfficeAgent.Core;

/// <summary>
/// Wires the flow: resolve a handle, open a package, route to a format module,
/// and run inspect / find / validate / apply. Apply validates first, then commits
/// on a separate editable package only when valid and not a dry run. The module
/// stabilises anchor ids before apply and each op is re-validated against live
/// state immediately before it is applied; if any op fails, the source document is
/// left untouched and the failure is returned as a structured <see cref="ValidationError"/>.
/// </summary>
internal sealed class FlowOrchestrator
{
    private readonly OpenXmlPackageService _opc = new();
    private readonly FormatRouter _router;
    private readonly PlanValidator _validator = new();
    private readonly TransactionManager _transaction = new();
    private readonly IReadOnlyList<IHandleResolver> _resolvers;
    private readonly ILogger _logger;

    public FlowOrchestrator(IEnumerable<IFormatModule> modules, IEnumerable<IHandleResolver> resolvers)
        : this(modules, resolvers, NullLogger.Instance)
    {
    }

    public FlowOrchestrator(IEnumerable<IFormatModule> modules, IEnumerable<IHandleResolver> resolvers, ILogger logger)
    {
        _router = new FormatRouter(modules);
        _resolvers = resolvers.ToList();
        _logger = logger;
    }

    public InspectResult Inspect(DocumentHandle handle, InspectOptions options) =>
        InspectCore(ReadAll(handle), options);

    public IReadOnlyList<FindHit> Find(DocumentHandle handle, FindQuery query) =>
        FindCore(ReadAll(handle), query);

    public ChangeReport Validate(DocumentHandle handle, DocumentPlan plan) =>
        ValidateBytes(ReadAll(handle), plan).Report;

    public ApplyResult Apply(DocumentHandle handle, DocumentPlan plan, ApplyOptions options) =>
        ApplyCore(ReadAll(handle), plan, options, handle, CancellationToken.None);

    public async Task<InspectResult> InspectAsync(DocumentHandle handle, InspectOptions options, CancellationToken cancellationToken) =>
        InspectCore(await ReadAllAsync(handle, cancellationToken).ConfigureAwait(false), options);

    public async Task<IReadOnlyList<FindHit>> FindAsync(DocumentHandle handle, FindQuery query, CancellationToken cancellationToken) =>
        FindCore(await ReadAllAsync(handle, cancellationToken).ConfigureAwait(false), query);

    public async Task<ChangeReport> ValidateAsync(DocumentHandle handle, DocumentPlan plan, CancellationToken cancellationToken) =>
        ValidateBytes(await ReadAllAsync(handle, cancellationToken).ConfigureAwait(false), plan).Report;

    public async Task<ApplyResult> ApplyAsync(DocumentHandle handle, DocumentPlan plan, ApplyOptions options, CancellationToken cancellationToken)
    {
        var bytes = await ReadAllAsync(handle, cancellationToken).ConfigureAwait(false);
        return ApplyCore(bytes, plan, options, handle, cancellationToken);
    }

    private InspectResult InspectCore(byte[] bytes, InspectOptions options)
    {
        using var activity = OfficeAgentTelemetry.ActivitySource.StartActivity("OfficeAgent.Inspect");
        activity?.SetTag("officeagent.bytes", bytes.Length);
        activity?.SetTag("officeagent.fidelity", options.Fidelity.ToString());

        var sw = Stopwatch.StartNew();
        using var package = _opc.Open(bytes, editable: false);
        var module = _router.Route(package);
        var result = module.Inspect(package, options);

        _logger.LogDebug("Inspect {Format} ({Bytes} B, fidelity={Fidelity}) → {Paragraphs} paragraphs, snapshot={Snapshot} in {Elapsed} ms",
            module.Format, bytes.Length, options.Fidelity, result.Paragraphs.Count, result.Snapshot.ETag, sw.ElapsedMilliseconds);
        return result;
    }

    private IReadOnlyList<FindHit> FindCore(byte[] bytes, FindQuery query)
    {
        using var activity = OfficeAgentTelemetry.ActivitySource.StartActivity("OfficeAgent.Find");
        activity?.SetTag("officeagent.bytes", bytes.Length);

        var sw = Stopwatch.StartNew();
        using var package = _opc.Open(bytes, editable: false);
        var module = _router.Route(package);
        var hits = module.Find(package, query);

        _logger.LogDebug("Find {Format} pattern='{Pattern}' → {Hits} hits in {Elapsed} ms",
            module.Format, query.Pattern, hits.Count, sw.ElapsedMilliseconds);
        return hits;
    }

    private ApplyResult ApplyCore(byte[] bytes, DocumentPlan plan, ApplyOptions options, DocumentHandle handle, CancellationToken cancellationToken)
    {
        using var activity = OfficeAgentTelemetry.ActivitySource.StartActivity("OfficeAgent.Apply");
        activity?.SetTag("officeagent.bytes", bytes.Length);
        activity?.SetTag("officeagent.operations", plan.Operations.Count);
        activity?.SetTag("officeagent.dryrun", options.DryRun);

        var sw = Stopwatch.StartNew();
        var (report, _) = ValidateBytes(bytes, plan);

        if (options.DryRun || !report.IsValid)
        {
            if (!report.IsValid)
                _logger.LogInformation("Apply rejected: {ErrorCount} validation error(s), first={FirstCode}",
                    report.Errors.Count, report.Errors[0].Code);
            else
                _logger.LogDebug("Apply dry-run produced {ChangeCount} proposed change(s) in {Elapsed} ms",
                    report.Changes.Count, sw.ElapsedMilliseconds);
            return new ApplyResult { Report = report, Committed = false, Output = null };
        }

        using var commitPackage = _opc.Open(bytes, editable: true);
        var commitModule = _router.Route(commitPackage);

        var aliases = commitModule.Stabilize(commitPackage);
        var commitSnapshot = commitModule.Inspect(commitPackage, InspectOptions.Default);
        var context = new ApplyContext(commitPackage, commitSnapshot, aliases);

        var outcome = _transaction.ApplyAtomic(context, plan, commitModule, cancellationToken);
        if (!outcome.Success)
        {
            _logger.LogWarning("Apply aborted mid-plan: {Code}: {Message}",
                outcome.Error!.Code, outcome.Error.Message);
            activity?.SetStatus(ActivityStatusCode.Error, outcome.Error.Code);
            return new ApplyResult
            {
                Report = new ChangeReport
                {
                    IsValid = false,
                    Changes = report.Changes,
                    Errors = new[] { outcome.Error! }
                },
                Committed = false,
                Output = null
            };
        }

        var outputBytes = commitPackage.ToBytes();
        var output = new StreamHandle(new MemoryStream(outputBytes), OutputName(handle));

        _logger.LogInformation("Apply committed {Operations} op(s), {InputBytes} B → {OutputBytes} B in {Elapsed} ms",
            plan.Operations.Count, bytes.Length, outputBytes.Length, sw.ElapsedMilliseconds);
        return new ApplyResult { Report = report, Committed = true, Output = output };
    }

    private (ChangeReport Report, IFormatModule Module) ValidateBytes(byte[] bytes, DocumentPlan plan)
    {
        using var package = _opc.Open(bytes, editable: false);
        var module = _router.Route(package);
        var snapshot = module.Inspect(package, InspectOptions.Default);
        var context = new ApplyContext(package, snapshot);
        var report = _validator.Validate(context, plan, module);
        return (report, module);
    }

    private Stream ResolveStream(DocumentHandle handle)
    {
        var resolver = _resolvers.FirstOrDefault(r => r.CanResolve(handle))
            ?? throw new NotSupportedException($"No handle resolver for {handle}.");
        return resolver.Resolve(handle);
    }

    private byte[] ReadAll(DocumentHandle handle)
    {
        using var source = ResolveStream(handle);
        if (source is MemoryStream ms)
            return ms.ToArray();
        using var copy = new MemoryStream();
        source.CopyTo(copy);
        return copy.ToArray();
    }

    private async Task<byte[]> ReadAllAsync(DocumentHandle handle, CancellationToken cancellationToken)
    {
        if (handle is FileHandle file)
        {
#if NET8_0_OR_GREATER
            return await File.ReadAllBytesAsync(file.Path, cancellationToken).ConfigureAwait(false);
#else
            cancellationToken.ThrowIfCancellationRequested();
            return await Task.Run(() => File.ReadAllBytes(file.Path), cancellationToken).ConfigureAwait(false);
#endif
        }

        using var source = ResolveStream(handle);
        if (source is MemoryStream ms)
            return ms.ToArray();
        using var copy = new MemoryStream();
        await source.CopyToAsync(copy, 81920, cancellationToken).ConfigureAwait(false);
        return copy.ToArray();
    }

    private static string OutputName(DocumentHandle handle) => handle switch
    {
        FileHandle f => Path.GetFileName(f.Path),
        StreamHandle s => s.Name ?? "output.docx",
        _ => "output.docx"
    };
}
