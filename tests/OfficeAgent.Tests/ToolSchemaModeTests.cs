using System.Text.Json;
using Microsoft.Extensions.AI;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Excel;
using OfficeAgent.Mcp;

namespace OfficeAgent.Tests;

/// <summary>
/// Strict schemas list every parameter as required, which OpenAI strict function calling
/// needs. Microsoft Copilot Studio asks its user for every required input the model leaves
/// out and cannot send an empty string, so under strict schemas it asked for a new file name
/// before every in-place save and never completed one. Relaxed schemas are the opt-in for
/// such clients: the same schemas, except that a parameter with a default is optional.
/// </summary>
public class ToolSchemaModeTests
{
    [Fact]
    public void Schemas_stay_strict_unless_the_host_turns_it_off()
    {
        Assert.True(new OfficeAgentToolsOptions().StrictToolSchemas);
        Assert.True(new OfficeAgentMcpOptions().StrictToolSchemas);
    }

    [Fact]
    public void Relaxed_schemas_require_only_the_parameters_without_a_default()
    {
        using var workspace = new Workspace();
        var strict = Schemas(workspace.Tools.AsAIFunctions(WidestSurface(strict: true)));
        var relaxed = Schemas(workspace.Tools.AsAIFunctions(WidestSurface(strict: false)));

        Assert.Equal(strict.Keys.OrderBy(k => k), relaxed.Keys.OrderBy(k => k));
        foreach (var (name, strictSchema) in strict)
        {
            var relaxedSchema = relaxed[name];

            // Same properties with the same types and defaults; unknown properties still refused.
            Assert.Equal(
                strictSchema.GetProperty("properties").GetRawText(),
                relaxedSchema.GetProperty("properties").GetRawText());
            Assert.Equal(JsonValueKind.False, relaxedSchema.GetProperty("additionalProperties").ValueKind);

            var withoutDefault = strictSchema.GetProperty("properties").EnumerateObject()
                .Where(property => !property.Value.TryGetProperty("default", out _))
                .Select(property => property.Name)
                .OrderBy(n => n, StringComparer.Ordinal);
            Assert.Equal(withoutDefault, Required(relaxedSchema).OrderBy(n => n, StringComparer.Ordinal));
        }

        Assert.Equal(new[] { "connectionId", "documentId", "planJson" }, Required(relaxed["apply_plan"]).OrderBy(n => n));
        Assert.Equal(new[] { "connectionId", "documentId" }, Required(relaxed["inspect_document"]).OrderBy(n => n));
        Assert.Equal(new[] { "connectionId", "documentId", "pattern" }, Required(relaxed["find_in_document"]).OrderBy(n => n));
    }

    [Fact]
    public async Task A_call_that_omits_a_defaulted_parameter_gets_the_default()
    {
        using var workspace = new Workspace();
        var functions = workspace.Tools.AsAIFunctions(new OfficeAgentToolsOptions { StrictToolSchemas = false })
            .ToDictionary(f => f.Name);
        var registered = await workspace.RegisterRegister();

        // Only the required arguments: no fidelity, range, maximumCells, saveMode or newName.
        var inspection = await Invoke(functions["inspect_document"], new()
        {
            ["connectionId"] = "registers",
            ["documentId"] = registered
        });
        var snapshot = inspection.GetProperty("snapshot").GetString();

        var applied = await Invoke(functions["apply_plan"], new()
        {
            ["connectionId"] = "registers",
            ["documentId"] = registered,
            ["planJson"] = Workspace.AppendRowPlan(snapshot)
        });

        Assert.True(applied.GetProperty("committed").GetBoolean(), applied.GetRawText());
        // Replace, the default: the same document, updated where it is.
        Assert.Equal(registered, applied.GetProperty("outputDocumentId").GetString());
        Assert.Equal(1, await workspace.RowsWith(registered, "APAC"));
    }

    [Fact]
    public void The_mcp_server_publishes_relaxed_schemas_when_configured()
    {
        using var workspace = new Workspace();
        var options = new OfficeAgentMcpOptions
        {
            StrictToolSchemas = false,
            FileSystemConnections =
            {
                new FileSystemConnectionOptions { ConnectionId = "registers", RootPath = workspace.Root }
            }
        };

        var applyPlan = OfficeAgentMcpServer.BuildToolset(options)
            .Single(tool => tool.ProtocolTool.Name == "apply_plan").ProtocolTool.InputSchema;

        Assert.Equal(new[] { "connectionId", "documentId", "planJson" }, Required(applyPlan).OrderBy(n => n));
        Assert.True(applyPlan.GetProperty("properties").TryGetProperty("newName", out _));
    }

    [Fact]
    public void The_mcp_server_keeps_strict_schemas_by_default()
    {
        using var workspace = new Workspace();
        var options = new OfficeAgentMcpOptions
        {
            FileSystemConnections =
            {
                new FileSystemConnectionOptions { ConnectionId = "registers", RootPath = workspace.Root }
            }
        };

        var applyPlan = OfficeAgentMcpServer.BuildToolset(options)
            .Single(tool => tool.ProtocolTool.Name == "apply_plan").ProtocolTool.InputSchema;

        Assert.Equal(
            new[] { "connectionId", "documentId", "newName", "planJson", "saveMode" },
            Required(applyPlan).OrderBy(n => n));
    }

    [Fact]
    public void Both_settings_bind_from_environment_variables()
    {
        const string schemas = "OfficeAgent__StrictToolSchemas";
        const string snapshot = "OfficeAgent__RequirePlanSnapshot";
        Environment.SetEnvironmentVariable(schemas, "false");
        Environment.SetEnvironmentVariable(snapshot, "true");
        try
        {
            var configuration = new ConfigurationBuilder().AddEnvironmentVariables().Build();

            var options = OfficeAgentConfiguration.Bind(configuration);

            Assert.False(options.StrictToolSchemas);
            Assert.True(options.RequirePlanSnapshot);
        }
        finally
        {
            Environment.SetEnvironmentVariable(schemas, null);
            Environment.SetEnvironmentVariable(snapshot, null);
        }
    }

    private static OfficeAgentToolsOptions WidestSurface(bool strict) => new()
    {
        AllowRegistration = true,
        AllowCreation = true,
        AllowInlineContent = true,
        AllowEphemeralDocuments = true,
        StrictToolSchemas = strict
    };

    private static Dictionary<string, JsonElement> Schemas(IEnumerable<AIFunction> functions) =>
        functions.ToDictionary(f => f.Name, f => f.JsonSchema, StringComparer.Ordinal);

    private static IEnumerable<string> Required(JsonElement schema) =>
        schema.TryGetProperty("required", out var required)
            ? required.EnumerateArray().Select(e => e.GetString()!)
            : Array.Empty<string>();

    private static async Task<JsonElement> Invoke(AIFunction function, AIFunctionArguments arguments)
    {
        var result = await function.InvokeAsync(arguments);
        var json = result switch
        {
            JsonElement { ValueKind: JsonValueKind.String } element => element.GetString(),
            JsonElement element => element.GetRawText(),
            _ => result?.ToString()
        };
        return JsonDocument.Parse(json!).RootElement.Clone();
    }

    private sealed class Workspace : IDisposable
    {
        private readonly ServiceProvider _services;

        public Workspace()
        {
            Root = Path.Combine(Path.GetTempPath(), $"officeagent-schema-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Root);
            var services = new ServiceCollection();
            services.AddExcelFormat();
            services.AddFileSystemDocumentProvider("registers", Root, o => o.AllowedExtensions = new[] { ".xlsx" });
            services.AddOfficeAgent();
            _services = services.BuildServiceProvider();
            Client = _services.GetRequiredService<OfficeAgentClient>();
            Tools = new OfficeAgentTools(Client);
        }

        public string Root { get; }
        public OfficeAgentClient Client { get; }
        public OfficeAgentTools Tools { get; }

        public async Task<string> RegisterRegister()
        {
            var reference = await Client.RegisterBytesAsync("registers", Root, XlsxFactory.WorkbookWithTable(), "register.xlsx");
            return reference.ItemId;
        }

        public static string AppendRowPlan(string? snapshot) => JsonSerializer.Serialize(new
        {
            snapshot = new { eTag = snapshot },
            operations = new object[]
            {
                new
                {
                    op = "appendTableRows",
                    target = new { kind = "spreadsheetTable", path = "table#7/Sales" },
                    rows = new[] { new[] { "APAC", "15" } }
                }
            }
        });

        public async Task<int> RowsWith(string documentId, string value)
        {
            var inspection = await Client.InspectAsync("registers", documentId);
            return inspection.Cells.Count(cell => cell.DisplayValue == value);
        }

        public void Dispose()
        {
            _services.Dispose();
            if (Directory.Exists(Root)) Directory.Delete(Root, recursive: true);
        }
    }
}
