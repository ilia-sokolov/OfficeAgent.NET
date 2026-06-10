using System.Text.Json;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

public class StrictSchemaTests
{
    [Fact]
    public void Every_tool_emits_strict_compatible_schema()
    {
        // Azure OpenAI (and OpenAI strict mode) requires that every property in a
        // tool's parameters schema is listed under "required" and that
        // "additionalProperties": false is set. Without this we hit:
        //   HTTP 400 invalid_function_parameters:
        //   'required' is required to be supplied and to be an array including
        //   every key in properties. Missing 'fidelity'.
        var registry = new DocumentProviderRegistry(Array.Empty<IDocumentProvider>());
        var client = new OfficeAgentClient(registry, new WordModule());
        var tools = new OfficeAgentTools(client).AsAIFunctions();

        Assert.NotEmpty(tools);
        foreach (var tool in tools)
        {
            var schema = tool.JsonSchema;
            Assert.True(schema.ValueKind == JsonValueKind.Object, $"{tool.Name} has no object schema");

            var props = schema.GetProperty("properties");
            var required = schema.GetProperty("required").EnumerateArray().Select(e => e.GetString()!).ToHashSet();

            foreach (var prop in props.EnumerateObject())
                Assert.Contains(prop.Name, required);

            Assert.True(schema.TryGetProperty("additionalProperties", out var additional)
                        && additional.ValueKind == JsonValueKind.False,
                $"{tool.Name} must set additionalProperties: false");
        }
    }
}
