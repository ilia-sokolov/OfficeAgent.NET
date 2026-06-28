// OfficeAgent MCP server. One binary, two transports:
//
//   officeagent-mcp --stdio       local hosting as a child process of an MCP
//                                 client (Claude Desktop, VS Code, an agent SDK)
//   officeagent-mcp               streamable HTTP on ASP.NET Core for cloud or
//                                 shared hosting; MCP endpoint at /, health at /healthz
//
// Configuration comes from appsettings.json, environment variables prefixed
// OfficeAgent__, and the command line. See docs/mcp-server.md.

using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using OfficeAgent.Mcp;
using OfficeAgent.SharePoint;

var version = typeof(OfficeAgentMcpServer).Assembly.GetName().Version?.ToString(3) ?? "0.0.0";

if (UseStdio(args))
{
    var builder = Host.CreateApplicationBuilder(args);
    var options = Bind(builder.Configuration);

    // stdout carries JSON-RPC frames in stdio mode; logs must go to stderr.
    builder.Logging.AddConsole(o => o.LogToStandardErrorThreshold = LogLevel.Trace);

    builder.Services
        .AddMcpServer(o =>
        {
            o.ServerInfo = new() { Name = OfficeAgentMcpServer.ServerName, Version = version };
            o.ServerInstructions = OfficeAgentMcpServer.InstructionsFor(options);
        })
        .WithStdioServerTransport()
        .WithTools(OfficeAgentMcpServer.BuildToolset(options));

    await builder.Build().RunAsync();
}
else
{
    var builder = WebApplication.CreateBuilder(args);
    var options = Bind(builder.Configuration);

    builder.Services
        .AddMcpServer(o =>
        {
            o.ServerInfo = new() { Name = OfficeAgentMcpServer.ServerName, Version = version };
            o.ServerInstructions = OfficeAgentMcpServer.InstructionsFor(options);
        })
        .WithHttpTransport()
        .WithTools(OfficeAgentMcpServer.BuildToolset(options));

    var app = builder.Build();

    // Capture the caller's inbound bearer token per request so SharePoint connections
    // configured for the On-Behalf-Of flow can exchange it for a user-scoped Graph
    // token. The token is held only for the duration of the request.
    if (UsesOnBehalfOf(options))
    {
        app.Use(async (context, next) =>
        {
            using (GraphUserContext.Push(BearerToken(context)))
                await next().ConfigureAwait(false);
        });
    }

    app.MapMcp();
    app.MapGet("/healthz", () => Results.Ok(new { status = "ok", server = OfficeAgentMcpServer.ServerName, version }));
    await app.RunAsync();
}

static bool UsesOnBehalfOf(OfficeAgentMcpOptions options) =>
    options.SharePointConnections.Any(c =>
        c.AuthMode?.Trim().ToLowerInvariant() is "onbehalfof" or "on-behalf-of" or "obo");

static string? BearerToken(HttpContext context)
{
    var header = context.Request.Headers.Authorization.ToString();
    return header.StartsWith("Bearer ", StringComparison.OrdinalIgnoreCase)
        ? header.Substring("Bearer ".Length).Trim()
        : null;
}

static bool UseStdio(string[] args)
{
    if (args.Contains("--stdio", StringComparer.OrdinalIgnoreCase)) return true;
    var probe = new ConfigurationBuilder()
        .AddJsonFile("appsettings.json", optional: true)
        .AddEnvironmentVariables()
        .AddCommandLine(args)
        .Build();
    return string.Equals(
        probe[$"{OfficeAgentMcpOptions.SectionName}:Transport"], "stdio", StringComparison.OrdinalIgnoreCase);
}

static OfficeAgentMcpOptions Bind(IConfiguration configuration) =>
    configuration.GetSection(OfficeAgentMcpOptions.SectionName).Get<OfficeAgentMcpOptions>()
    ?? new OfficeAgentMcpOptions();
