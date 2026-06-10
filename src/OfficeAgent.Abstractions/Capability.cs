namespace OfficeAgent.Abstractions;

/// <summary>
/// Describes whether a proposed change can be completed by the deterministic document engine.
/// </summary>
public enum Capability
{
    /// <summary>
    /// The change can be completed with deterministic package edits.
    /// </summary>
    Deterministic,

    /// <summary>
    /// The package edit is deterministic, but the displayed value is refreshed by Word when the document opens.
    /// </summary>
    DeferredToWordOnOpen,

    /// <summary>
    /// The change requires a renderer, layout engine, or calculation engine and is rejected by validation.
    /// </summary>
    NeedsRenderer
}
