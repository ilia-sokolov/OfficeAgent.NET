using System.Diagnostics;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using OfficeAgent.Abstractions;
using OfficeAgent.Core.DocumentProviders;

namespace OfficeAgent.Core;

/// <summary>
/// The public entry point. Composes the engine over one or more format modules and
/// offers the inspect → find → preview → commit surface.
/// </summary>
/// <remarks>
/// <para>
/// Documents are addressed by an opaque, provider-assigned id - never by a filesystem
/// path. Register an existing document with a provider connection via
/// <see cref="RegisterAsync"/>, which returns the document's id, then drive
/// inspect/find/preview/commit with <c>(connectionId, documentId)</c>. The provider
/// stores only the reference (path, URL, …); the host owns the underlying file's
/// lifecycle. In-memory <see cref="StreamHandle"/> content remains supported for
/// callers that already hold bytes and don't need provider-backed storage.
/// </para>
/// <para>
/// Instances are safe for concurrent use; share one per host. Inspect/Find/Validate are
/// pure reads; Apply opens a fresh editable package per call so concurrent edits on
/// different documents do not interfere.
/// </para>
/// </remarks>
public sealed class OfficeAgentClient
{
    private readonly IDocumentService _service;
    private readonly DocumentProviderRegistry _providers;
    private readonly ILogger _logger;

    public OfficeAgentClient(params IFormatModule[] modules)
        : this(new OfficeAgentEngine(modules))
    {
    }

    public OfficeAgentClient(IDocumentService service)
        : this(service, new DocumentProviderRegistry(Array.Empty<IDocumentProvider>()), NullLoggerFactory.Instance)
    {
    }

    /// <summary>Initializes a client over format modules with provider-backed document access.</summary>
    public OfficeAgentClient(DocumentProviderRegistry providers, params IFormatModule[] modules)
        : this(new OfficeAgentEngine(modules), providers, NullLoggerFactory.Instance)
    {
    }

    /// <summary>Initializes a client with provider-backed document access.</summary>
    public OfficeAgentClient(
        IDocumentService service,
        DocumentProviderRegistry providers,
        ILoggerFactory? loggerFactory = null)
    {
        _service = service ?? throw new ArgumentNullException(nameof(service));
        _providers = providers ?? throw new ArgumentNullException(nameof(providers));
        _logger = (loggerFactory ?? NullLoggerFactory.Instance).CreateLogger(OfficeAgentTelemetry.LogCategory);
    }

    // ---- Core surface (DocumentHandle) ----

    public InspectResult Inspect(DocumentHandle handle, InspectOptions? options = null) =>
        _service.Inspect(handle, options ?? InspectOptions.Default);

    public IReadOnlyList<FindHit> Find(DocumentHandle handle, FindQuery query) =>
        _service.Find(handle, query);

    public ChangeReport Preview(DocumentHandle handle, DocumentPlan plan) =>
        _service.Validate(handle, plan);

    public ApplyResult Commit(DocumentHandle handle, DocumentPlan plan) =>
        _service.Apply(handle, plan, ApplyOptions.Commit);

    public ApplyResult Apply(DocumentHandle handle, DocumentPlan plan, ApplyOptions? options = null) =>
        _service.Apply(handle, plan, options ?? ApplyOptions.Preview);

    public Task<InspectResult> InspectAsync(DocumentHandle handle, InspectOptions? options = null, CancellationToken cancellationToken = default) =>
        _service.InspectAsync(handle, options ?? InspectOptions.Default, cancellationToken);

    public Task<IReadOnlyList<FindHit>> FindAsync(DocumentHandle handle, FindQuery query, CancellationToken cancellationToken = default) =>
        _service.FindAsync(handle, query, cancellationToken);

    public Task<ChangeReport> PreviewAsync(DocumentHandle handle, DocumentPlan plan, CancellationToken cancellationToken = default) =>
        _service.ValidateAsync(handle, plan, cancellationToken);

    public Task<ApplyResult> CommitAsync(DocumentHandle handle, DocumentPlan plan, CancellationToken cancellationToken = default) =>
        _service.ApplyAsync(handle, plan, ApplyOptions.Commit, cancellationToken);

    public Task<ApplyResult> ApplyAsync(DocumentHandle handle, DocumentPlan plan, ApplyOptions? options = null, CancellationToken cancellationToken = default) =>
        _service.ApplyAsync(handle, plan, options ?? ApplyOptions.Preview, cancellationToken);

    // ---- In-memory byte overload ----

    public InspectResult Inspect(byte[] document, InspectOptions? options = null) =>
        Inspect(new StreamHandle(new MemoryStream(document, writable: false)), options);

    // ---- Provider-backed surface ----

    /// <summary>
    /// Registers an existing document with a configured provider connection and returns
    /// its canonical reference, including the provider-assigned opaque
    /// <see cref="DocumentReference.ItemId"/>. The provider stores only the reference
    /// (path, URL, drive id, …), not the bytes - the host owns the underlying file's
    /// lifecycle. <paramref name="source"/> is provider-specific: a filesystem path for
    /// the filesystem provider, a sharing URL or drive id for cloud providers. Defaults
    /// to the filesystem provider.
    /// </summary>
    public async Task<DocumentReference> RegisterAsync(
        string connectionId,
        string source,
        CancellationToken cancellationToken = default)
    {
        var provider = _providers.Resolve(DocumentReference.ForFileSystem(connectionId, string.Empty));
        var reference = await provider.RegisterAsync(source, cancellationToken).ConfigureAwait(false);
        _logger.LogInformation(
            "Provider register {Provider}:{ConnectionId} '{Source}' → {ItemId}",
            provider.Provider, provider.ConnectionId, source, reference.ItemId);
        return reference;
    }

    /// <summary>Removes a document from a registered provider connection.</summary>
    public Task RemoveAsync(
        DocumentReference reference,
        CancellationToken cancellationToken = default)
    {
        var provider = _providers.Resolve(reference);
        return provider.RemoveAsync(reference, cancellationToken);
    }

    /// <summary>Removes a filesystem-provider document by its opaque id.</summary>
    public Task RemoveAsync(
        string connectionId,
        string documentId,
        CancellationToken cancellationToken = default) =>
        RemoveAsync(DocumentReference.ForFileSystem(connectionId, documentId), cancellationToken);

    /// <summary>Opens a document and returns its canonical current reference plus bytes.</summary>
    public Task<DocumentContent> OpenReadAsync(
        DocumentReference reference,
        CancellationToken cancellationToken = default) =>
        OpenWithTelemetryAsync(reference, cancellationToken);

    /// <summary>Inspects a provider document.</summary>
    public async Task<InspectResult> InspectAsync(
        DocumentReference reference,
        InspectOptions? options = null,
        CancellationToken cancellationToken = default)
    {
        using var content = await OpenWithTelemetryAsync(reference, cancellationToken).ConfigureAwait(false);
        return await InspectAsync(
            new StreamHandle(content.Stream, content.Reference.Name), options, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Finds content in a provider document.</summary>
    public async Task<IReadOnlyList<FindHit>> FindAsync(
        DocumentReference reference,
        FindQuery query,
        CancellationToken cancellationToken = default)
    {
        using var content = await OpenWithTelemetryAsync(reference, cancellationToken).ConfigureAwait(false);
        return await FindAsync(
            new StreamHandle(content.Stream, content.Reference.Name), query, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Previews a plan against a provider document without saving.</summary>
    public async Task<ChangeReport> PreviewAsync(
        DocumentReference reference,
        DocumentPlan plan,
        CancellationToken cancellationToken = default)
    {
        plan = await ResolveImageReferencesAsync(plan, cancellationToken).ConfigureAwait(false);
        using var content = await OpenWithTelemetryAsync(reference, cancellationToken).ConfigureAwait(false);
        return await PreviewAsync(
            new StreamHandle(content.Stream, content.Reference.Name), plan, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Commits a plan and saves the result through the selected provider.</summary>
    public async Task<ProviderApplyResult> CommitAsync(
        DocumentReference reference,
        DocumentPlan plan,
        SaveDocumentOptions? options = null,
        CancellationToken cancellationToken = default)
    {
        plan = await ResolveImageReferencesAsync(plan, cancellationToken).ConfigureAwait(false);
        var provider = _providers.Resolve(reference);
        using var content = await OpenWithTelemetryAsync(provider, reference, cancellationToken).ConfigureAwait(false);
        using var result = await CommitAsync(
            new StreamHandle(content.Stream, content.Reference.Name), plan, cancellationToken).ConfigureAwait(false);

        if (!result.Committed)
        {
            _logger.LogInformation(
                "Provider commit rejected for {Provider}:{ConnectionId}/{ItemId} - plan invalid",
                provider.Provider, provider.ConnectionId, reference.ItemId);
            return new ProviderApplyResult { Report = result.Report, Committed = false };
        }

        var bytes = result.ToBytes();
        using var output = new MemoryStream(bytes, writable: false);
        var saveOpts = options ?? new SaveDocumentOptions();

        using var saveActivity = OfficeAgentTelemetry.ActivitySource.StartActivity("OfficeAgent.Provider.Save");
        saveActivity?.SetTag("officeagent.provider", provider.Provider);
        saveActivity?.SetTag("officeagent.connectionId", provider.ConnectionId);
        saveActivity?.SetTag("officeagent.itemId", reference.ItemId);
        saveActivity?.SetTag("officeagent.save_mode", saveOpts.Mode.ToString());
        saveActivity?.SetTag("officeagent.bytes", bytes.Length);

        var sw = Stopwatch.StartNew();
        var saved = await provider.SaveAsync(content.Reference, output, saveOpts, cancellationToken).ConfigureAwait(false);

        _logger.LogInformation(
            "Provider save {Provider}:{ConnectionId} {Source} → {Destination} ({Bytes} B, mode={Mode}) in {Elapsed} ms",
            provider.Provider, provider.ConnectionId, reference.ItemId, saved.ItemId, bytes.Length, saveOpts.Mode, sw.ElapsedMilliseconds);

        return new ProviderApplyResult
        {
            Report = result.Report,
            Committed = true,
            Document = saved
        };
    }

    // ── Opaque-id overloads (default to the filesystem provider) ──────────────

    /// <summary>Inspects a provider document by <c>(connectionId, documentId)</c>, where <paramref name="documentId"/> is the opaque id returned by <see cref="RegisterAsync"/> or a save.</summary>
    public Task<InspectResult> InspectAsync(string connectionId, string documentId, InspectOptions? options = null, CancellationToken cancellationToken = default) =>
        InspectAsync(DocumentReference.ForFileSystem(connectionId, documentId), options, cancellationToken);

    /// <summary>Finds content in a provider document by <c>(connectionId, documentId)</c>.</summary>
    public Task<IReadOnlyList<FindHit>> FindAsync(string connectionId, string documentId, FindQuery query, CancellationToken cancellationToken = default) =>
        FindAsync(DocumentReference.ForFileSystem(connectionId, documentId), query, cancellationToken);

    /// <summary>Previews a plan against a provider document by <c>(connectionId, documentId)</c>.</summary>
    public Task<ChangeReport> PreviewAsync(string connectionId, string documentId, DocumentPlan plan, CancellationToken cancellationToken = default) =>
        PreviewAsync(DocumentReference.ForFileSystem(connectionId, documentId), plan, cancellationToken);

    /// <summary>Commits a plan against a provider document by <c>(connectionId, documentId)</c>.</summary>
    public Task<ProviderApplyResult> CommitAsync(string connectionId, string documentId, DocumentPlan plan, SaveDocumentOptions? options = null, CancellationToken cancellationToken = default) =>
        CommitAsync(DocumentReference.ForFileSystem(connectionId, documentId), plan, options, cancellationToken);

    // ── Internal helpers ──────────────────────────────────────────────────────

    private Task<DocumentContent> OpenWithTelemetryAsync(DocumentReference reference, CancellationToken cancellationToken) =>
        OpenWithTelemetryAsync(_providers.Resolve(reference), reference, cancellationToken);

    private async Task<DocumentContent> OpenWithTelemetryAsync(IDocumentProvider provider, DocumentReference reference, CancellationToken cancellationToken)
    {
        using var activity = OfficeAgentTelemetry.ActivitySource.StartActivity("OfficeAgent.Provider.Open");
        activity?.SetTag("officeagent.provider", provider.Provider);
        activity?.SetTag("officeagent.connectionId", provider.ConnectionId);
        activity?.SetTag("officeagent.itemId", reference.ItemId);

        var sw = Stopwatch.StartNew();
        try
        {
            var content = await provider.OpenReadAsync(reference, cancellationToken).ConfigureAwait(false);
            _logger.LogDebug(
                "Provider open {Provider}:{ConnectionId}/{ItemId} → {Bytes} B, version={Version} in {Elapsed} ms",
                provider.Provider, provider.ConnectionId, reference.ItemId,
                content.Stream.CanSeek ? content.Stream.Length : -1,
                content.Reference.Version, sw.ElapsedMilliseconds);
            return content;
        }
        catch (DocumentProviderException ex)
        {
            activity?.SetStatus(ActivityStatusCode.Error, ex.Code.ToString());
            _logger.LogWarning(
                "Provider open failed for {Provider}:{ConnectionId}/{ItemId}: {Code} - {Message}",
                provider.Provider, provider.ConnectionId, reference.ItemId, ex.Code, ex.Message);
            throw;
        }
    }

    // ---- Plan preprocessing ----

    /// <summary>
    /// Walks the plan and resolves any <see cref="InsertImageOp"/> that references
    /// an image by <c>(ImageConnectionId, ImageDocumentId)</c> into its base64 form
    /// by reading the bytes from the configured provider. The original plan is
    /// returned unchanged when no image ops reference a provider id.
    /// </summary>
    private async Task<DocumentPlan> ResolveImageReferencesAsync(DocumentPlan plan, CancellationToken cancellationToken)
    {
        if (plan.Operations.Count == 0) return plan;
        if (!plan.Operations.OfType<InsertImageOp>().Any(o => !string.IsNullOrEmpty(o.ImageDocumentId)))
            return plan;

        var rewritten = new List<PlanOperation>(plan.Operations.Count);
        foreach (var op in plan.Operations)
        {
            if (op is InsertImageOp image && !string.IsNullOrEmpty(image.ImageDocumentId))
            {
                var bytes = await OpenImageBytesAsync(image.ImageConnectionId!, image.ImageDocumentId!, cancellationToken).ConfigureAwait(false);
                rewritten.Add(new InsertImageOp
                {
                    Target = image.Target,
                    Base64Bytes = Convert.ToBase64String(bytes),
                    ImageType = image.ImageType,
                    WidthPx = image.WidthPx,
                    HeightPx = image.HeightPx,
                    Position = image.Position,
                    AltText = image.AltText
                });
            }
            else
            {
                rewritten.Add(op);
            }
        }

        return new DocumentPlan
        {
            ContractVersion = plan.ContractVersion,
            Format = plan.Format,
            Snapshot = plan.Snapshot,
            Operations = rewritten
        };
    }

    private async Task<byte[]> OpenImageBytesAsync(string connectionId, string documentId, CancellationToken cancellationToken)
    {
        var reference = DocumentReference.ForFileSystem(connectionId, documentId);
        using var content = await OpenWithTelemetryAsync(reference, cancellationToken).ConfigureAwait(false);
        if (content.Stream is MemoryStream ms)
            return ms.ToArray();
        using var copy = new MemoryStream();
        await content.Stream.CopyToAsync(copy, 81920, cancellationToken).ConfigureAwait(false);
        return copy.ToArray();
    }

    // ---- Back-compat statics (delegate to ApplyResult instance methods) ----

    /// <summary>
    /// Deprecated. Prefer <see cref="ApplyResult.ToBytes"/>.
    /// </summary>
    public static byte[] ToBytes(ApplyResult result) => result.ToBytes();

    /// <summary>
    /// Deprecated. Prefer <see cref="ApplyResult.Save"/>.
    /// </summary>
    public static void Save(ApplyResult result, string path) => result.Save(path);
}
