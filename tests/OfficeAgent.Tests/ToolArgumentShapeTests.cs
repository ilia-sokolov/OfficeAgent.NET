using System.Text.Json;
using Microsoft.Extensions.AI;
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Excel;

namespace OfficeAgent.Tests;

/// <summary>
/// Agents do not always send a string parameter as a string. A blind run sent create_document's
/// planJson as the plan object itself, the default binder threw before the tool ran, and the MCP
/// client saw only "An error occurred invoking 'create_document'." A plan sent as an object or
/// array is now read as its JSON text, and a value of the wrong shape gets a diagnostic.
/// </summary>
public class ToolArgumentShapeTests
{
    private const string TitlePlan =
        """{"operations":[{"op":"setProperty","target":{"kind":"docProperty","path":"core/title"},"value":"Q3 review"}]}""";

    /// <summary>A plan sent as an object is applied as if it had been sent as a string.</summary>
    [Fact]
    public async Task A_plan_sent_as_an_object_is_applied()
    {
        using var workspace = new Workspace();

        var result = await workspace.Invoke("create_document",
            $$"""{"connectionId":"documents","name":"object.xlsx","planJson":{{TitlePlan}}}""");

        Assert.True(result.GetProperty("committed").GetBoolean(), result.GetRawText());
        Assert.Equal("Q3 review", await workspace.TitleOf(result.GetProperty("outputDocumentId").GetString()!));
    }

    /// <summary>A bare operations array, which planJson accepts as a string, is accepted unquoted too.</summary>
    [Fact]
    public async Task A_plan_sent_as_a_bare_array_is_applied()
    {
        using var workspace = new Workspace();
        var operations = JsonDocument.Parse(TitlePlan).RootElement.GetProperty("operations").GetRawText();

        var result = await workspace.Invoke("create_document",
            $$"""{"connectionId":"documents","name":"array.xlsx","planJson":{{operations}}}""");

        Assert.True(result.GetProperty("committed").GetBoolean(), result.GetRawText());
    }

    /// <summary>The string form, the documented one, still works and gives the same document.</summary>
    [Fact]
    public async Task A_plan_sent_as_a_string_still_works()
    {
        using var workspace = new Workspace();

        var result = await workspace.Invoke("create_document",
            JsonSerializer.Serialize(new { connectionId = "documents", name = "string.xlsx", planJson = TitlePlan }));

        Assert.True(result.GetProperty("committed").GetBoolean(), result.GetRawText());
        Assert.Equal("Q3 review", await workspace.TitleOf(result.GetProperty("outputDocumentId").GetString()!));
    }

    /// <summary>A value that is no plan at all is refused with the tool's own diagnostic, and nothing is written.</summary>
    [Fact]
    public async Task A_plan_of_the_wrong_shape_gets_a_diagnostic_not_an_exception()
    {
        using var workspace = new Workspace();

        var result = await workspace.Invoke("create_document",
            """{"connectionId":"documents","name":"number.xlsx","planJson":42}""");

        Assert.False(result.GetProperty("committed").GetBoolean());
        Assert.Equal("notWritten", result.GetProperty("writeOutcome").GetString());
        Assert.False(File.Exists(Path.Combine(workspace.Root, "number.xlsx")));
    }

    /// <summary>The schema still advertises a string: strict function calling clients rely on it.</summary>
    [Fact]
    public void The_schema_still_says_planJson_is_a_string()
    {
        using var workspace = new Workspace();

        var schema = workspace.Function("create_document").JsonSchema;

        Assert.Equal("string", schema.GetProperty("properties").GetProperty("planJson").GetProperty("type").GetString());
    }

    private sealed class Workspace : IDisposable
    {
        private readonly ServiceProvider _services;
        private readonly OfficeAgentTools _tools;

        public Workspace()
        {
            Root = Path.Combine(Path.GetTempPath(), $"officeagent-args-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Root);
            var services = new ServiceCollection();
            services.AddExcelFormat();
            services.AddFileSystemDocumentProvider("documents", Root, o => o.AllowedExtensions = new[] { ".xlsx" });
            services.AddOfficeAgent();
            _services = services.BuildServiceProvider();
            Client = _services.GetRequiredService<OfficeAgentClient>();
            _tools = new OfficeAgentTools(Client);
        }

        public string Root { get; }
        public OfficeAgentClient Client { get; }

        public AIFunction Function(string name) =>
            _tools.AsAIFunctions(new OfficeAgentToolsOptions { AllowRegistration = true, AllowCreation = true }).Single(f => f.Name == name);

        /// <summary>Invokes a tool with arguments as an MCP server hands them over: JSON elements, not CLR strings.</summary>
        public async Task<JsonElement> Invoke(string name, string argumentsJson)
        {
            var arguments = new AIFunctionArguments(JsonDocument.Parse(argumentsJson).RootElement
                .EnumerateObject().ToDictionary(p => p.Name, p => (object?)p.Value.Clone()));
            var result = await Function(name).InvokeAsync(arguments);
            var json = result switch
            {
                JsonElement { ValueKind: JsonValueKind.String } element => element.GetString(),
                JsonElement element => element.GetRawText(),
                _ => result?.ToString()
            };
            return JsonDocument.Parse(json!).RootElement.Clone();
        }

        public async Task<string?> TitleOf(string documentId)
        {
            var inspection = await Client.InspectAsync("documents", documentId);
            return inspection.Nodes.FirstOrDefault(n => n.Path == "core/title")?.Summary;
        }

        public void Dispose()
        {
            _services.Dispose();
            if (Directory.Exists(Root)) Directory.Delete(Root, recursive: true);
        }
    }
}
