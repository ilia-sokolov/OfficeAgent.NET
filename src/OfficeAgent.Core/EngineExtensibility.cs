namespace OfficeAgent.Core;

/// <summary>
/// The diagnostic reported for a public type that exists to build format modules and operation
/// handlers rather than to host the engine.
/// </summary>
/// <remarks>
/// These types stay public so a format module can live in its own assembly, but they are the
/// engine's internal architecture, and 1.x may add members to them. On netstandard2.0 an
/// interface cannot gain a member without breaking every implementer, so freezing them would
/// freeze the engine. Hosts never need them: registering formats, providers, access policies and
/// audit actors uses only the stable surface. See docs/compatibility.md.
/// </remarks>
internal static class EngineExtensibility
{
    /// <summary>The diagnostic id a caller suppresses to build against the engine seams.</summary>
    public const string DiagnosticId = "OFFICEAGENT001";

    /// <summary>Where the diagnostic points.</summary>
    public const string UrlFormat =
        "https://github.com/ilia-sokolov/OfficeAgent.NET/blob/main/docs/compatibility.md#engine-extensibility";
}
