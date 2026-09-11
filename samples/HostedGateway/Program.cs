using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authorization;
using ModelContextProtocol.Server;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Mcp;
using OfficeAgent.Samples.HostedGateway;

var app = HostedGatewayApp.Build(args);
await app.RunAsync();

/// <summary>Entry point marker used by integration tests.</summary>
public partial class Program;

namespace OfficeAgent.Samples.HostedGateway
{
    /// <summary>Builds the authenticated reference gateway.</summary>
    public static class HostedGatewayApp
    {
        /// <summary>Creates the application without starting it.</summary>
        public static WebApplication Build(
            string[] args,
            Action<WebApplicationBuilder>? configureBuilder = null)
        {
            var builder = WebApplication.CreateBuilder(args);
            configureBuilder?.Invoke(builder);

            builder.Logging.AddFilter("Microsoft.AspNetCore", LogLevel.Warning);
            builder.Logging.AddFilter("Microsoft.Hosting", LogLevel.Warning);
            builder.Logging.AddFilter("ModelContextProtocol", LogLevel.Warning);

            builder.Services
                .AddAuthentication(DemoBearerHandler.Scheme)
                .AddScheme<AuthenticationSchemeOptions, DemoBearerHandler>(DemoBearerHandler.Scheme, _ => { });
            builder.Services.AddAuthorization();
            var httpContextAccessor = new HttpContextAccessor();
            builder.Services.AddSingleton<IHttpContextAccessor>(httpContextAccessor);

            var connectionARoot = Root(builder.Configuration, "HostedGateway:ConnectionARoot", "connection-a");
            var connectionBRoot = Root(builder.Configuration, "HostedGateway:ConnectionBRoot", "connection-b");
            Directory.CreateDirectory(connectionARoot);
            Directory.CreateDirectory(connectionBRoot);

            var options = new OfficeAgentMcpOptions
            {
                AllowRegistration = true,
                AllowCreation = true,
                FileSystemConnections =
                {
                    Connection("connection-a", connectionARoot),
                    Connection("connection-b", connectionBRoot)
                }
            };
            var accessPolicy = new HostedConnectionAccessPolicy();
            var principalAccessor = new HttpContextPrincipalAccessor(httpContextAccessor);
            var auditActorProvider = new HttpContextAuditActorProvider(httpContextAccessor);
            builder.Services.AddSingleton<IConnectionAccessPolicy>(accessPolicy);
            builder.Services.AddSingleton<ITrustedPrincipalAccessor>(principalAccessor);
            builder.Services.AddSingleton<IAuditActorProvider>(auditActorProvider);
            builder.Services.AddSingleton(options);

            builder.Services
                .AddMcpServer(server =>
                {
                    server.ServerInfo = new() { Name = "officeagent-hosted-gateway", Version = "1.0.0" };
                    // Per-caller connection ids cannot be placed in static initialization text.
                    // list_connections is the policy-filtered discovery surface.
                    server.ServerInstructions = OfficeAgentMcpServer.InstructionsFor(
                        options, includeConnectionInventory: false);
                })
                .WithHttpTransport()
                .WithTools(OfficeAgentMcpServer.BuildToolset(
                    options, accessPolicy, principalAccessor, auditActorProvider));

            var app = builder.Build();
            app.UseAuthentication();
            app.UseAuthorization();
            app.MapMcp().RequireAuthorization();
            app.MapGet("/healthz", () => Results.Ok(new { status = "ok" })).AllowAnonymous();
            return app;
        }

        private static FileSystemConnectionOptions Connection(string id, string root) => new()
        {
            ConnectionId = id,
            RootPath = root,
            AllowedExtensions = new List<string> { ".docx", ".pptx", ".xlsx" },
            DefaultChangeMode = "Direct"
        };

        private static string Root(IConfiguration configuration, string key, string directory) =>
            Path.GetFullPath(configuration[key] ?? Path.Combine(AppContext.BaseDirectory, "documents", directory));
    }
}
