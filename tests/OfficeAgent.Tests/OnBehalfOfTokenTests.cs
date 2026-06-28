using System.Net;
using System.Text;
using System.Text.Json;
using OfficeAgent.SharePoint;

namespace OfficeAgent.Tests;

/// <summary>
/// On-Behalf-Of token-exchange tests over a fake Entra token endpoint: the provider
/// exchanges the captured user assertion for a Graph token, caches per user, surfaces
/// AADSTS errors, and refuses to run without an inbound user token.
/// </summary>
public class OnBehalfOfTokenTests
{
    [Fact]
    public async Task Exchanges_the_user_assertion_for_a_graph_token()
    {
        using var entra = new FakeEntra();
        string? userToken = "user-token-alice";
        var provider = new OnBehalfOfAccessTokenProvider(Options(), entra.Client, () => userToken);

        var graphToken = await provider.GetAccessTokenAsync();

        Assert.Equal("graph-for(user-token-alice)", graphToken);
        var exchange = Assert.Single(entra.Requests);
        Assert.Equal("urn:ietf:params:oauth:grant-type:jwt-bearer", exchange["grant_type"]);
        Assert.Equal("on_behalf_of", exchange["requested_token_use"]);
        Assert.Equal("user-token-alice", exchange["assertion"]);
        Assert.Equal("https://graph.microsoft.com/.default", exchange["scope"]);
    }

    [Fact]
    public async Task Caches_per_user_and_re_exchanges_for_a_different_user()
    {
        using var entra = new FakeEntra();
        string? userToken = "user-token-alice";
        var provider = new OnBehalfOfAccessTokenProvider(Options(), entra.Client, () => userToken);

        await provider.GetAccessTokenAsync();
        await provider.GetAccessTokenAsync();           // served from cache
        Assert.Single(entra.Requests);

        userToken = "user-token-bob";                   // different user → new exchange
        var bobToken = await provider.GetAccessTokenAsync();
        Assert.Equal("graph-for(user-token-bob)", bobToken);
        Assert.Equal(2, entra.Requests.Count);
    }

    [Fact]
    public async Task Fails_clearly_when_no_user_token_is_present()
    {
        using var entra = new FakeEntra();
        var provider = new OnBehalfOfAccessTokenProvider(Options(), entra.Client, () => null);

        var ex = await Assert.ThrowsAsync<InvalidOperationException>(() => provider.GetAccessTokenAsync());
        Assert.Contains("On-Behalf-Of", ex.Message);
        Assert.Empty(entra.Requests);
    }

    [Fact]
    public async Task Surfaces_consent_required_errors_from_entra()
    {
        using var entra = new FakeEntra { FailWith = (HttpStatusCode.BadRequest, "invalid_grant", "AADSTS65001: The user has not consented.") };
        var provider = new OnBehalfOfAccessTokenProvider(Options(), entra.Client, () => "user-token-alice");

        var ex = await Assert.ThrowsAsync<InvalidOperationException>(() => provider.GetAccessTokenAsync());
        Assert.Contains("AADSTS65001", ex.Message);
    }

    [Fact]
    public void Ambient_user_context_flows_and_restores()
    {
        Assert.Null(GraphUserContext.CurrentUserAccessToken);
        using (GraphUserContext.Push("outer"))
        {
            Assert.Equal("outer", GraphUserContext.CurrentUserAccessToken);
            using (GraphUserContext.Push("inner"))
                Assert.Equal("inner", GraphUserContext.CurrentUserAccessToken);
            Assert.Equal("outer", GraphUserContext.CurrentUserAccessToken);
        }
        Assert.Null(GraphUserContext.CurrentUserAccessToken);
    }

    [Fact]
    public async Task Default_constructor_reads_the_ambient_context()
    {
        using var entra = new FakeEntra();
        var provider = new OnBehalfOfAccessTokenProvider(Options(), entra.Client);

        using (GraphUserContext.Push("user-token-carol"))
        {
            var token = await provider.GetAccessTokenAsync();
            Assert.Equal("graph-for(user-token-carol)", token);
        }
    }

    private static OnBehalfOfOptions Options() => new()
    {
        TenantId = "00000000-0000-0000-0000-000000000000",
        ClientId = "api-client-id",
        ClientSecret = "api-secret",
        Authority = "https://login.fake"
    };

    private sealed class FakeEntra : IDisposable
    {
        public FakeEntra() => Client = new HttpClient(new Handler(this));

        public HttpClient Client { get; }
        public List<Dictionary<string, string>> Requests { get; } = new();
        public (HttpStatusCode Status, string Error, string Description)? FailWith { get; set; }

        public void Dispose() => Client.Dispose();

        private sealed class Handler : HttpMessageHandler
        {
            private readonly FakeEntra _entra;
            public Handler(FakeEntra entra) => _entra = entra;

            protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
            {
                var body = await request.Content!.ReadAsStringAsync().ConfigureAwait(false);
                var form = body.Split('&')
                    .Select(p => p.Split('='))
                    .ToDictionary(p => Uri.UnescapeDataString(p[0]), p => Uri.UnescapeDataString(p[1]));
                _entra.Requests.Add(form);

                if (_entra.FailWith is { } failure)
                    return Json(failure.Status, new { error = failure.Error, error_description = failure.Description });

                return Json(HttpStatusCode.OK, new
                {
                    access_token = $"graph-for({form["assertion"]})",
                    expires_in = 3600
                });
            }

            private static HttpResponseMessage Json(HttpStatusCode status, object payload) => new(status)
            {
                Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json")
            };
        }
    }
}
