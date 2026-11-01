using System.Text;
using System.Text.Json;
using System.Threading;
using OfficeAgent.Abstractions;

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
    // A mutable frame rather than the value itself: the token exchange runs in a nested async
    // call, and a value set there would never reach the code that writes the receipt.
    private static readonly AsyncLocal<Frame?> Current = new();

    /// <summary>
    /// Gets the end user's access token for the current logical call, or
    /// <see langword="null"/> when none was captured (for example a stdio host with
    /// no inbound HTTP request).
    /// </summary>
    public static string? CurrentUserAccessToken => Current.Value?.Token;

    /// <summary>
    /// Gets the signed-in user for the current logical call, or <see langword="null"/> until
    /// Entra has accepted their token in an On-Behalf-Of exchange. An inbound header proves
    /// nothing by itself, so this stays empty without that exchange.
    /// </summary>
    public static AuditActor? ValidatedActor => Current.Value?.Actor;

    /// <summary>
    /// Sets the end user's access token for the current logical call and returns a
    /// scope that restores the previous value when disposed. Wrap a request in
    /// <c>using (GraphUserContext.Push(token)) { ... }</c>.
    /// </summary>
    public static IDisposable Push(string? userAccessToken)
    {
        var previous = Current.Value;
        Current.Value = new Frame(userAccessToken);
        return new Scope(previous);
    }

    /// <summary>
    /// Records that Entra accepted this token, which is what makes its claims trustworthy.
    /// Ignored unless it is the token of the current call.
    /// </summary>
    internal static void MarkValidated(string userAccessToken)
    {
        var frame = Current.Value;
        if (frame is null || !string.Equals(frame.Token, userAccessToken, StringComparison.Ordinal)) return;
        frame.Actor ??= ActorFrom(userAccessToken);
    }

    /// <summary>
    /// Reads <c>oid</c>, <c>iss</c> and <c>name</c> from a token Entra has already accepted, so
    /// the signature is not checked again. <c>oid</c> is used rather than <c>sub</c> because it
    /// is the same across applications in a tenant. Without an <c>oid</c> there is no actor.
    /// </summary>
    private static AuditActor? ActorFrom(string token)
    {
        var parts = token.Split('.');
        if (parts.Length < 2) return null;

        try
        {
            var payload = parts[1].Replace('-', '+').Replace('_', '/');
            payload = payload.PadRight(payload.Length + ((4 - (payload.Length % 4)) % 4), '=');

            using var claims = JsonDocument.Parse(Encoding.UTF8.GetString(Convert.FromBase64String(payload)));
            var root = claims.RootElement;
            if (Claim(root, "oid") is not { Length: > 0 } subject) return null;
            return new AuditActor { Subject = subject, Issuer = Claim(root, "iss"), DisplayName = Claim(root, "name") };
        }
        catch (Exception ex) when (ex is FormatException or JsonException or InvalidOperationException)
        {
            return null;
        }
    }

    private static string? Claim(JsonElement root, string name) =>
        root.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String
            ? value.GetString()
            : null;

    private sealed class Frame
    {
        public Frame(string? token) => Token = token;

        public string? Token { get; }
        public AuditActor? Actor { get; set; }
    }

    private sealed class Scope : IDisposable
    {
        private readonly Frame? _previous;
        private bool _disposed;

        public Scope(Frame? previous) => _previous = previous;

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            Current.Value = _previous;
        }
    }
}
