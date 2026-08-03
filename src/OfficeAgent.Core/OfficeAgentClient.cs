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
    private readonly IReadOnlyList<IBlankDocumentFactory> _blankDocumentFactories;
    private readonly ILogger _logger;

    public OfficeAgentClient(params IFormatModule[] modules)
        : this(
            new OfficeAgentEngine(modules),
            new DocumentProviderRegistry(Array.Empty<IDocumentProvider>()),
            modules.OfType<IBlankDocumentFactory>().ToArray(),
            loggerFactory: null)
    {
    }

    public OfficeAgentClient(IDocumentService service)
        : this(service, new DocumentProviderRegistry(Array.Empty<IDocumentProvider>()), NullLoggerFactory.Instance)
    {
    }

    /// <summary>Initializes a client over format modules with provider-backed document access.</summary>
    public OfficeAgentClient(DocumentProviderRegistry providers, params IFormatModule[] modules)
        : this(new OfficeAgentEngine(modules), providers, modules.OfType<IBlankDocumentFactory>().ToArray(), loggerFactory: null)
    {
    }

    /// <summary>Initializes a client with provider-backed document access.</summary>
    public OfficeAgentClient(
        IDocumentService service,
        DocumentProviderRegistry providers,
        ILoggerFactory? loggerFactory = null)
        : this(service, providers, Array.Empty<IBlankDocumentFactory>(), loggerFactory)
    {
    }

    /// <summary>
    /// Initializes a client over a custom document service with factories for formats
    /// that support creating blank documents.
    /// </summary>
    public OfficeAgentClient(
        IDocumentService service,
        DocumentProviderRegistry providers,
        ILoggerFactory? loggerFactory,
        IEnumerable<IBlankDocumentFactory> blankDocumentFactories)
        : this(
            service,
            providers,
            (blankDocumentFactories ?? throw new ArgumentNullException(nameof(blankDocumentFactories))).ToArray(),
            loggerFactory)
    {
    }

    private OfficeAgentClient(
        IDocumentService service,
        DocumentProviderRegistry providers,
        IReadOnlyList<IBlankDocumentFactory> blankDocumentFactories,
        ILoggerFactory? loggerFactory)
    {
        _service = service ?? throw new ArgumentNullException(nameof(service));
        _providers = providers ?? throw new ArgumentNullException(nameof(providers));
        _logger = (loggerFactory ?? NullLoggerFactory.Instance).CreateLogger(OfficeAgentTelemetry.LogCategory);
        _blankDocumentFactories = blankDocumentFactories;
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
    /// the filesystem provider, a drive-relative path for the SharePoint provider. The
    /// connection id alone selects the provider, whatever its type.
    /// </summary>
    public async Task<DocumentReference> RegisterAsync(
        string connectionId,
        string source,
        CancellationToken cancellationToken = default)
    {
        var provider = _providers.ResolveConnection(connectionId);
        var reference = await provider.RegisterAsync(source, cancellationToken).ConfigureAwait(false);
        _logger.LogInformation(
            "Provider register {Provider}:{ConnectionId} '{Source}' → {ItemId}",
            provider.Provider, provider.ConnectionId, source, reference.ItemId);
        return reference;
    }

    /// <summary>
    /// Creates a new document inside a configured connection: mints an empty but valid
    /// package for the format implied by <paramref name="name"/>'s extension, optionally
    /// applies <paramref name="plan"/> to it, and writes the result through the provider,
    /// which registers it and returns its opaque id.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The plan is applied in memory <em>before</em> anything is written, so a plan that
    /// fails validation leaves no trace: the returned result has
    /// <see cref="ProviderApplyResult.Committed"/> <see langword="false"/>, a
    /// <see langword="null"/> <see cref="ProviderApplyResult.Document"/>, and the errors
    /// to act on. Nothing is ever overwritten - a name already in use fails with
    /// <see cref="ProviderErrorCode.AlreadyExists"/>.
    /// </para>
    /// <para>
    /// On the result, <see cref="ProviderApplyResult.Committed"/> means "the document was
    /// created", and when <paramref name="plan"/> is <see langword="null"/> the
    /// <see cref="ProviderApplyResult.Report"/> is a synthetic empty valid report - the
    /// type is shared with <see cref="CommitAsync(DocumentReference, DocumentPlan, SaveDocumentOptions?, CancellationToken)"/>
    /// so both reach an agent as the same JSON shape.
    /// </para>
    /// <para>
    /// A blank Word document holds a single empty body paragraph, which inspection
    /// addresses as paragraph id <c>auto-0000</c>; an initial plan targets that anchor.
    /// </para>
    /// <para>
    /// Only the format-independent part of <paramref name="name"/> is checked before the
    /// package is minted; connection-specific naming rules (the extension allow-list and
    /// bare file name) are the provider's and are applied last, when the result
    /// is written.
    /// </para>
    /// </remarks>
    /// <exception cref="DocumentProviderException">
    /// The connection is unknown, its provider cannot create documents
    /// (<see cref="IDocumentCreatingProvider"/>), no format module can mint the requested
    /// format, or the provider refused the write.
    /// </exception>
    public async Task<ProviderApplyResult> CreateAsync(
        string connectionId,
        string name,
        DocumentPlan? plan = null,
        CancellationToken cancellationToken = default)
    {
        var provider = _providers.ResolveConnection(connectionId);
        if (provider is not IDocumentCreatingProvider creator)
            throw new DocumentProviderException(
                ProviderErrorCode.ConfigurationError,
                $"The '{provider.Provider}' connection '{connectionId}' does not support creating documents; " +
                "register an existing document instead.",
                provider.Provider, connectionId, itemId: null);

        var bytes = CreateBlankDocument(name, provider);
        var report = new ChangeReport { IsValid = true };

        if (plan is not null)
        {
            // Every supplied plan is validated, including an empty one: its Format and
            // Snapshot are part of the contract, and a plan that disagrees with the
            // document should fail before anything is written rather than after.
            plan = await ResolveImageReferencesAsync(plan, cancellationToken).ConfigureAwait(false);
            using var applied = await CommitAsync(
                new StreamHandle(new MemoryStream(bytes, writable: false), name), plan, cancellationToken).ConfigureAwait(false);

            if (!applied.Committed)
            {
                _logger.LogInformation(
                    "Provider create rejected for {Provider}:{ConnectionId} '{Name}' - initial plan invalid",
                    provider.Provider, provider.ConnectionId, name);
                return new ProviderApplyResult { Report = applied.Report, Committed = false };
            }

            bytes = applied.ToBytes();
            report = applied.Report;
        }

        using var output = new MemoryStream(bytes, writable: false);
        using var activity = OfficeAgentTelemetry.ActivitySource.StartActivity("OfficeAgent.Provider.Create");
        activity?.SetTag("officeagent.provider", provider.Provider);
        activity?.SetTag("officeagent.connectionId", provider.ConnectionId);
        activity?.SetTag("officeagent.bytes", bytes.Length);

        var sw = Stopwatch.StartNew();
        var created = await creator.CreateAsync(name, output, cancellationToken).ConfigureAwait(false);

        _logger.LogInformation(
            "Provider create {Provider}:{ConnectionId} '{Name}' → {ItemId} ({Bytes} B, {Operations} initial op(s)) in {Elapsed} ms",
            provider.Provider, provider.ConnectionId, name, created.ItemId, bytes.Length,
            plan?.Operations.Count ?? 0, sw.ElapsedMilliseconds);

        return new ProviderApplyResult { Report = report, Committed = true, Document = created };
    }

    /// <summary>
    /// Mints the empty package a new document starts from, choosing the format module
    /// from the requested file name's extension.
    /// </summary>
    private byte[] CreateBlankDocument(string name, IDocumentProvider provider)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new DocumentProviderException(
                ProviderErrorCode.InvalidArgument,
                "A document name is required, including its extension (for example 'report.docx').",
                provider.Provider, provider.ConnectionId, itemId: null);

        // Control characters and path separators are refused up front, before the name can
        // reach a log line or an error message: the connection-specific rules run later, at
        // the provider, and by then a name carrying an escape sequence has already been
        // written to the host's log.
        if (name.Any(character => char.IsControl(character)) ||
            name.IndexOfAny(PortableInvalidDocumentNameChars) >= 0)
            throw new DocumentProviderException(
                ProviderErrorCode.InvalidArgument,
                "A document name must be a bare file name without path separators, control characters, or other invalid filename characters.",
                provider.Provider, provider.ConnectionId, itemId: null);

        // Surrounding whitespace and a trailing dot are silently stripped by Windows, so a
        // name carrying them would register under one spelling and land on disk under
        // another. Reject them here rather than let the extension lookup fail obscurely.
        if (!string.Equals(name, name.Trim(), StringComparison.Ordinal) ||
            name.EndsWith(".", StringComparison.Ordinal))
            throw new DocumentProviderException(
                ProviderErrorCode.InvalidArgument,
                $"The document name '{name}' must not begin or end with whitespace or end with a dot.",
                provider.Provider, provider.ConnectionId, itemId: null);

        var extension = Path.GetExtension(name);
        var factory = _blankDocumentFactories.FirstOrDefault(candidate =>
            !string.IsNullOrWhiteSpace(candidate.Extension) &&
            string.Equals(extension, candidate.Extension, StringComparison.OrdinalIgnoreCase));
        if (extension.Length == 0 || factory is null)
            throw new DocumentProviderException(
                ProviderErrorCode.InvalidArgument,
                $"No registered format module can create a document named '{name}'. " +
                $"Creatable extensions: {(_blankDocumentFactories.Count == 0 ? "none (register a format module such as AddWordFormat())" : string.Join(", ", _blankDocumentFactories.Select(item => item.Extension)))}.",
                provider.Provider, provider.ConnectionId, itemId: null);

        return factory.CreateBlank();
    }

    private static readonly char[] PortableInvalidDocumentNameChars =
        { '<', '>', ':', '"', '|', '?', '*', '/', '\\', '\0' };

    /// <summary>Removes a document from a registered provider connection.</summary>
    public Task RemoveAsync(
        DocumentReference reference,
        CancellationToken cancellationToken = default)
    {
        var provider = _providers.Resolve(reference);
        return provider.RemoveAsync(reference, cancellationToken);
    }

    /// <summary>Removes a provider document registration by its opaque id.</summary>
    public Task RemoveAsync(
        string connectionId,
        string documentId,
        CancellationToken cancellationToken = default) =>
        RemoveAsync(ReferenceFor(connectionId, documentId), cancellationToken);

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

    // ── Opaque-id overloads (connection id selects the provider) ──────────────

    /// <summary>Inspects a provider document by <c>(connectionId, documentId)</c>, where <paramref name="documentId"/> is the opaque id returned by <see cref="RegisterAsync"/> or a save.</summary>
    public Task<InspectResult> InspectAsync(string connectionId, string documentId, InspectOptions? options = null, CancellationToken cancellationToken = default) =>
        InspectAsync(ReferenceFor(connectionId, documentId), options, cancellationToken);

    /// <summary>Finds content in a provider document by <c>(connectionId, documentId)</c>.</summary>
    public Task<IReadOnlyList<FindHit>> FindAsync(string connectionId, string documentId, FindQuery query, CancellationToken cancellationToken = default) =>
        FindAsync(ReferenceFor(connectionId, documentId), query, cancellationToken);

    /// <summary>Previews a plan against a provider document by <c>(connectionId, documentId)</c>.</summary>
    public Task<ChangeReport> PreviewAsync(string connectionId, string documentId, DocumentPlan plan, CancellationToken cancellationToken = default) =>
        PreviewAsync(ReferenceFor(connectionId, documentId), plan, cancellationToken);

    /// <summary>Commits a plan against a provider document by <c>(connectionId, documentId)</c>.</summary>
    public Task<ProviderApplyResult> CommitAsync(string connectionId, string documentId, DocumentPlan plan, SaveDocumentOptions? options = null, CancellationToken cancellationToken = default) =>
        CommitAsync(ReferenceFor(connectionId, documentId), plan, options, cancellationToken);

    /// <summary>
    /// Returns the change mode a connection applies to an operation that does not state
    /// one. Hosts that build plans in code set <see cref="ChangeTextOp.Mode"/> themselves;
    /// this is for surfaces that accept a plan from an agent, where an absent <c>mode</c>
    /// should follow the connection's review policy rather than a single global rule.
    /// Unknown connections yield <see cref="ChangeMode.Tracked"/>.
    /// </summary>
    public ChangeMode DefaultChangeModeFor(string connectionId) =>
        _providers.DefaultChangeModeFor(connectionId);

    // ── Internal helpers ──────────────────────────────────────────────────────

    private DocumentReference ReferenceFor(string connectionId, string documentId)
    {
        var provider = _providers.ResolveConnection(connectionId);
        return DocumentReference.For(provider.Provider, connectionId, documentId);
    }

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
        var reference = ReferenceFor(connectionId, documentId);
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
