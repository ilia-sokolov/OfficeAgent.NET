using System.Security.Claims;

namespace OfficeAgent.AgentFramework;

/// <summary>The connection capability requested by an agent-facing document operation.</summary>
public enum ConnectionCapability
{
    /// <summary>Read document metadata or content, including previewing a plan.</summary>
    Read,
    /// <summary>Register an existing provider item as an addressable document.</summary>
    Register,
    /// <summary>Create a new document or session item.</summary>
    Create,
    /// <summary>Modify an existing document.</summary>
    Edit,
    /// <summary>Remove a registration. The provider document itself is not deleted.</summary>
    Delete
}

/// <summary>
/// Authorizes one trusted caller principal to use one capability on one configured
/// connection. Hosted applications should implement this boundary with their tenant and
/// resource policy; local applications may retain the allow-all default.
/// </summary>
public interface IConnectionAccessPolicy
{
    /// <summary>Returns whether the caller may use the requested connection capability.</summary>
    ValueTask<bool> IsAllowedAsync(
        ClaimsPrincipal principal,
        string connectionId,
        ConnectionCapability capability,
        CancellationToken cancellationToken = default);
}

/// <summary>Provides the principal authenticated by the host rather than by the model.</summary>
public interface ITrustedPrincipalAccessor
{
    /// <summary>Gets the trusted principal for the current tool invocation.</summary>
    ClaimsPrincipal Principal { get; }
}

/// <summary>Preserves the pre-policy behavior for local and stdio hosts.</summary>
public sealed class AllowAllConnectionAccessPolicy : IConnectionAccessPolicy
{
    /// <inheritdoc />
    public ValueTask<bool> IsAllowedAsync(
        ClaimsPrincipal principal,
        string connectionId,
        ConnectionCapability capability,
        CancellationToken cancellationToken = default) => new(true);
}

/// <summary>Provides an unauthenticated principal to local allow-all hosts.</summary>
public sealed class AnonymousPrincipalAccessor : ITrustedPrincipalAccessor
{
    private static readonly ClaimsPrincipal Anonymous = new(new ClaimsIdentity());

    /// <inheritdoc />
    public ClaimsPrincipal Principal => Anonymous;
}

internal sealed class ConnectionForbiddenException : Exception
{
    public ConnectionForbiddenException() : base("The caller is not allowed to use this connection.") { }
}
