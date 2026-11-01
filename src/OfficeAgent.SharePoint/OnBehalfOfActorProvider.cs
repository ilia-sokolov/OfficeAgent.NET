using OfficeAgent.Abstractions;
using OfficeAgent.Core;

namespace OfficeAgent.SharePoint;

/// <summary>
/// Names the signed-in user on audit receipts, once an On-Behalf-Of exchange has proved who
/// they are. Stdio, filesystem and app-only calls have no such exchange, so they have no actor.
/// The user's token must arrive through <see cref="GraphUserContext.Push"/>: a host that supplies
/// it through its own accessor has no request scope to record the user in, and gets no actor.
/// </summary>
public sealed class OnBehalfOfActorProvider : IAuditActorProvider
{
    /// <inheritdoc />
    public AuditActor? GetCurrentActor() => GraphUserContext.ValidatedActor;
}
