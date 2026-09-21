using System.Text.Json;
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// Every tool response uses camelCase property names.
/// </summary>
/// <remarks>
/// Before 1.0 the agent adapter serialised with three configurations, and casing depended on how
/// a payload happened to be built: <c>describe_capabilities</c> was PascalCase throughout,
/// <c>inspect_document</c> mixed ten PascalCase and seventeen camelCase names in one payload, and
/// a validation error carried <c>Code</c> and <c>Message</c> among camelCase siblings. The same
/// concept changed case between tools: <c>ParaId</c> from inspect, <c>paraId</c> from find.
/// This fixes the decision in a test, so a new payload built from a typed record cannot quietly
/// reintroduce PascalCase.
/// </remarks>
public sealed class OutputCasingTests : IDisposable
{
    private readonly string _root = Path.Combine(Path.GetTempPath(), $"officeagent-casing-{Guid.NewGuid():N}");
    private readonly ServiceProvider _services;
    private readonly OfficeAgentClient _client;
    private readonly OfficeAgentTools _tools;

    public OutputCasingTests()
    {
        Directory.CreateDirectory(_root);
        var services = new ServiceCollection();
        services.AddWordFormat();
        services.AddFileSystemDocumentProvider("workspace", _root);
        services.AddOfficeAgent();
        _services = services.BuildServiceProvider();
        _client = _services.GetRequiredService<OfficeAgentClient>();
        _tools = new OfficeAgentTools(_client);
    }

    [Fact]
    public async Task Every_tool_response_uses_camel_case_property_names()
    {
        var document = await _client.RegisterBytesAsync("workspace", _root, DocxFactory.Contract(), "contract.docx");
        var inspect = JsonDocument.Parse(await _tools.InspectDocument("workspace", document.ItemId));

        // Read the setup case-insensitively, so that if casing regresses this test fails at the
        // assertion and names every offending field, rather than throwing here first.
        var paragraph = Prop(inspect.RootElement, "paragraphs").EnumerateArray()
            .First(p => Prop(p, "text").GetString() is { Length: > 0 });
        var validPlan = JsonSerializer.Serialize(new
        {
            operations = new[]
            {
                new
                {
                    op = "changeText",
                    target = new
                    {
                        paraId = Prop(paragraph, "paraId").GetString(),
                        expect = Prop(paragraph, "text").GetString()
                    },
                    with = "Revised."
                }
            }
        });

        var responses = new (string Tool, string Json)[]
        {
            ("describe_capabilities", await _tools.DescribeCapabilities()),
            ("inspect_document", inspect.RootElement.GetRawText()),
            ("find_in_document", await _tools.FindInDocument("workspace", document.ItemId, "Acme")),
            ("preview_plan", await _tools.PreviewPlan("workspace", document.ItemId, validPlan)),
            ("preview_plan, invalid json", await _tools.PreviewPlan("workspace", document.ItemId, "{\"bogus\":1}")),
            ("apply_plan, unknown document", await _tools.ApplyPlan("workspace", "no-such-id", validPlan)),
            ("compare_documents", await _tools.CompareDocuments("workspace", document.ItemId, "workspace", document.ItemId)),
            ("apply_plan", await _tools.ApplyPlan("workspace", document.ItemId, validPlan)),
        };

        var offenders = responses
            .SelectMany(response => PascalCaseNames(JsonDocument.Parse(response.Json).RootElement, "$")
                .Select(path => $"{response.Tool}: {path}"))
            .ToList();

        Assert.True(offenders.Count == 0, string.Join(Environment.NewLine, offenders));
    }

    [Fact]
    public async Task The_same_concept_has_the_same_name_in_every_tool()
    {
        // The pre-1.0 symptom: inspect returned ParaId while find returned paraId, so a host
        // copying an anchor out of one response into the next had to know which tool it came from.
        var document = await _client.RegisterBytesAsync("workspace", _root, DocxFactory.Contract(), "contract.docx");
        var inspect = JsonDocument.Parse(await _tools.InspectDocument("workspace", document.ItemId));
        var find = JsonDocument.Parse(await _tools.FindInDocument("workspace", document.ItemId, "Acme"));

        Assert.True(inspect.RootElement.GetProperty("paragraphs")[0].TryGetProperty("paraId", out _));
        Assert.True(find.RootElement[0].TryGetProperty("paraId", out _));
        Assert.False(inspect.RootElement.GetProperty("paragraphs")[0].TryGetProperty("ParaId", out _));
    }

    private static JsonElement Prop(JsonElement element, string name) =>
        element.EnumerateObject().First(property =>
            string.Equals(property.Name, name, StringComparison.OrdinalIgnoreCase)).Value;

    /// <summary>Paths to every property name starting with an upper-case letter.</summary>
    private static IEnumerable<string> PascalCaseNames(JsonElement element, string path)
    {
        switch (element.ValueKind)
        {
            case JsonValueKind.Object:
                foreach (var property in element.EnumerateObject())
                {
                    var here = $"{path}.{property.Name}";
                    if (property.Name.Length > 0 && char.IsUpper(property.Name[0]))
                        yield return here;
                    foreach (var nested in PascalCaseNames(property.Value, here))
                        yield return nested;
                }
                break;
            case JsonValueKind.Array:
                var index = 0;
                foreach (var item in element.EnumerateArray())
                    foreach (var nested in PascalCaseNames(item, $"{path}[{index++}]"))
                        yield return nested;
                break;
        }
    }

    public void Dispose()
    {
        _services.Dispose();
        if (Directory.Exists(_root)) Directory.Delete(_root, recursive: true);
    }
}
