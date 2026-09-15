namespace OfficeAgent.Abstractions;

/// <summary>
/// What this engine supports, as configured right now: wire versions, formats and their
/// verbs, the host's ingestion ceilings, and optional features.
/// </summary>
/// <remarks>
/// <para>
/// Every field is derived from live registrations rather than from a maintained list, so
/// a module that is not registered contributes nothing and a verb no handler accepts is
/// not advertised. The point is that an agent can stop guessing: if a verb is absent
/// here, sending it fails.
/// </para>
/// <para>
/// This describes engine support. It does not describe what a particular caller may do
/// with a particular document. Connection access is reported separately and only for
/// connections the caller may actually use, and constraints that depend on a document's
/// own content are discoverable only by inspecting that document.
/// </para>
/// </remarks>
public sealed class EngineCapabilities
{
    /// <summary>Gets the wire contract versions this build accepts and emits.</summary>
    public ContractVersions Contracts { get; init; } = new();

    /// <summary>Gets the registered formats and what each supports.</summary>
    public IReadOnlyList<FormatCapabilities> Formats { get; init; } = Array.Empty<FormatCapabilities>();

    /// <summary>Gets the host ingestion ceilings in force.</summary>
    public OpenXmlIngestionLimits Limits { get; init; } = OpenXmlIngestionLimits.Default;

    /// <summary>
    /// Gets whether an optional page renderer is registered. Rendering is not part of the
    /// core engine and is absent unless a host wires it up.
    /// </summary>
    public bool RenderingAvailable { get; init; }

    /// <summary>
    /// Gets the connections this caller may address, with the capabilities allowed on
    /// each. Connections the caller cannot use are absent rather than listed as denied,
    /// so the result never discloses that they exist.
    /// </summary>
    public IReadOnlyList<ConnectionCapabilities> Connections { get; init; } =
        Array.Empty<ConnectionCapabilities>();

    /// <summary>
    /// Gets the things this result deliberately cannot answer, so a reader is not left to
    /// assume silence means support.
    /// </summary>
    public IReadOnlyList<string> RequiresInspection { get; init; } = Array.Empty<string>();
}

/// <summary>The wire contract versions this build accepts and emits.</summary>
public sealed class ContractVersions
{
    /// <summary>Gets the only edit-plan contract version accepted.</summary>
    public string EditPlan { get; init; } = DocumentPlan.CurrentContractVersion;

    /// <summary>Gets the apply-receipt schema version emitted.</summary>
    public string ApplyReceipt { get; init; } = OfficeAgent.Abstractions.ApplyReceipt.CurrentReceiptVersion;

    /// <summary>Gets the Word assembly receipt schema version.</summary>
    public string MergeReceipt { get; init; } = DocumentMergeReceipt.CurrentReceiptVersion;
}

/// <summary>What one registered format supports.</summary>
public sealed class FormatCapabilities
{
    /// <summary>Gets the document format.</summary>
    public DocumentFormat Format { get; init; }

    /// <summary>
    /// Gets the plan verbs at least one registered handler for this format accepts,
    /// sorted. Derived by asking the handlers, so it cannot drift from what validation
    /// actually does.
    /// </summary>
    public IReadOnlyList<string> Operations { get; init; } = Array.Empty<string>();

    /// <summary>
    /// Gets the change modes this format can honour. PowerPoint has no tracked-changes
    /// representation, so a tracked request there is refused rather than downgraded.
    /// </summary>
    public IReadOnlyList<ChangeMode> ChangeModes { get; init; } = Array.Empty<ChangeMode>();

    /// <summary>Gets the node kinds this format's inspection surfaces, sorted.</summary>
    public IReadOnlyList<string> NodeKinds { get; init; } = Array.Empty<string>();
}

/// <summary>One connection the caller may use, and what it may do there.</summary>
public sealed class ConnectionCapabilities
{
    /// <summary>Gets the host-configured connection identifier.</summary>
    public string ConnectionId { get; init; } = string.Empty;

    /// <summary>Gets the provider type, for example <c>filesystem</c> or <c>sharepoint</c>.</summary>
    public string Provider { get; init; } = string.Empty;

    /// <summary>Gets the capabilities this caller is allowed on this connection, sorted.</summary>
    public IReadOnlyList<string> Allowed { get; init; } = Array.Empty<string>();
}
