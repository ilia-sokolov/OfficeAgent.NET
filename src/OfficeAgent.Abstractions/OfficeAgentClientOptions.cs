namespace OfficeAgent.Abstractions;

/// <summary>Client-wide settings a host chooses once.</summary>
public sealed class OfficeAgentClientOptions
{
    /// <summary>
    /// Gets whether tracked changes and comments are authored with the authenticated actor's
    /// display name instead of the plan's author, which stays on the receipt. With no actor,
    /// or an actor with no usable name, the plan's author is written as usual. Off by default,
    /// because <c>OfficeAgent</c> is the documented default author in 1.x.
    /// </summary>
    public bool RevisionAuthorFromActor { get; init; }
}
