using System.Security.Claims;
using System.Text.Encodings.Web;
using Microsoft.AspNetCore.Authentication;
using Microsoft.Extensions.Options;
using OfficeAgent.AgentFramework;

namespace OfficeAgent.Samples.HostedGateway;

/// <summary>Maps authenticated sample users to their one permitted connection.</summary>
public sealed class HostedConnectionAccessPolicy : IConnectionAccessPolicy
{
    /// <inheritdoc />
    public ValueTask<bool> IsAllowedAsync(
        ClaimsPrincipal principal,
        string connectionId,
        ConnectionCapability capability,
        CancellationToken cancellationToken = default)
    {
        var subject = principal.FindFirstValue(ClaimTypes.NameIdentifier);
        var allowed = subject switch
        {
            "alice" => "connection-a",
            "bob" => "connection-b",
            _ => null
        };
        return new ValueTask<bool>(string.Equals(allowed, connectionId, StringComparison.Ordinal));
    }
}

/// <summary>Reads the principal established by ASP.NET Core authentication.</summary>
public sealed class HttpContextPrincipalAccessor : ITrustedPrincipalAccessor
{
    private static readonly ClaimsPrincipal Anonymous = new(new ClaimsIdentity());
    private readonly IHttpContextAccessor _httpContext;

    /// <summary>Creates the accessor over the host's current HTTP request.</summary>
    public HttpContextPrincipalAccessor(IHttpContextAccessor httpContext) =>
        _httpContext = httpContext ?? throw new ArgumentNullException(nameof(httpContext));

    /// <inheritdoc />
    public ClaimsPrincipal Principal => _httpContext.HttpContext?.User ?? Anonymous;
}

/// <summary>
/// Demonstration authentication only: bearer token <c>alice-token</c> or <c>bob-token</c> becomes the
/// corresponding trusted principal. Replace this handler with validated JWT authentication
/// in a real deployment.
/// </summary>
public sealed class DemoBearerHandler : AuthenticationHandler<AuthenticationSchemeOptions>
{
    /// <summary>The sample authentication scheme.</summary>
    public new const string Scheme = "DemoBearer";

    /// <summary>Initializes the handler.</summary>
    public DemoBearerHandler(
        IOptionsMonitor<AuthenticationSchemeOptions> options,
        ILoggerFactory logger,
        UrlEncoder encoder) : base(options, logger, encoder) { }

    /// <inheritdoc />
    protected override Task<AuthenticateResult> HandleAuthenticateAsync()
    {
        var header = Request.Headers.Authorization.ToString();
        if (!header.StartsWith("Bearer ", StringComparison.OrdinalIgnoreCase))
            return Task.FromResult(AuthenticateResult.NoResult());

        var token = header.Substring("Bearer ".Length).Trim();
        var subject = token switch
        {
            "alice-token" => "alice",
            "bob-token" => "bob",
            _ => null
        };
        if (subject is null)
            return Task.FromResult(AuthenticateResult.Fail("Invalid demonstration token."));

        var identity = new ClaimsIdentity(new[]
        {
            new Claim(ClaimTypes.NameIdentifier, subject),
            new Claim(ClaimTypes.Name, subject)
        }, Scheme);
        return Task.FromResult(AuthenticateResult.Success(
            new AuthenticationTicket(new ClaimsPrincipal(identity), Scheme)));
    }
}
