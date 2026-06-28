using System.Threading;

namespace OfficeAgent.SharePoint;

/// <summary>
/// Flows the caller's (end user's) access token to the SharePoint provider for the
/// duration of one request, so the On-Behalf-Of flow can exchange it for a Graph
/// token that carries the user's identity and delegated permissions.
/// </summary>
/// <remarks>
/// The value is ambient (an <see cref="AsyncLocal{T}"/>): a host captures the
/// inbound token once - typically ASP.NET middleware reading the request's
/// <c>Authorization</c> header - and the provider's token exchange reads it without
/// any change to the provider or client call signatures. The token is never stored
/// beyond the request scope and never written to a document reference.
/// </remarks>
public static class GraphUserContext
{
    private static readonly AsyncLocal<string?> Current = new();

    /// <summary>
    /// Gets the end user's access token for the current logical call, or
    /// <see langword="null"/> when none was captured (for example a stdio host with
    /// no inbound HTTP request).
    /// </summary>
    public static string? CurrentUserAccessToken => Current.Value;

    /// <summary>
    /// Sets the end user's access token for the current logical call and returns a
    /// scope that restores the previous value when disposed. Wrap a request in
    /// <c>using (GraphUserContext.Push(token)) { ... }</c>.
    /// </summary>
    public static IDisposable Push(string? userAccessToken)
    {
        var previous = Current.Value;
        Current.Value = userAccessToken;
        return new Scope(previous);
    }

    private sealed class Scope : IDisposable
    {
        private readonly string? _previous;
        private bool _disposed;

        public Scope(string? previous) => _previous = previous;

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            Current.Value = _previous;
        }
    }
}
