using System.Text.Json;
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.Abstractions;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Excel;
using OfficeAgent.PowerPoint;
using OfficeAgent.Rendering;
using OfficeAgent.SharePoint;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// Rendering is optional: nothing outside OfficeAgent.Rendering depends on it, structural
/// preview works without it, and discovery reports whether a host registered one.
/// </summary>
public sealed class RenderingOptionalTests
{
    [Fact]
    public void No_other_library_depends_on_the_rendering_package()
    {
        var others = new[]
        {
            typeof(DocumentPlan).Assembly, typeof(OfficeAgentClient).Assembly, typeof(WordModule).Assembly,
            typeof(PowerPointModule).Assembly, typeof(ExcelModule).Assembly, typeof(OfficeAgentTools).Assembly,
            typeof(SharePointDocumentProvider).Assembly, typeof(OfficeAgent.Mcp.OfficeAgentMcpOptions).Assembly,
        };
        var rendering = typeof(LibreOfficeDocumentRenderer).Assembly.GetName().Name;
        foreach (var assembly in others)
            Assert.DoesNotContain(assembly.GetReferencedAssemblies(), reference => reference.Name == rendering);
    }

    [Fact]
    public async Task Structural_preview_works_with_no_renderer_and_discovery_says_so()
    {
        var memory = new MemoryDocumentProvider("memory");
        using var services = Services(renderer: null, memory);
        var client = services.GetRequiredService<OfficeAgentClient>();
        Assert.False(client.DescribeCapabilities().RenderingAvailable);

        var tools = new OfficeAgentTools(client);
        var document = memory.Add("contract.docx", DocxFactory.Contract());
        var hit = (await client.FindAsync(document, new FindQuery("Acme Corp"))).First();
        var preview = await client.PreviewAsync(document, new DocumentPlan
        {
            Operations = new PlanOperation[] { new ChangeTextOp { Target = hit.Anchor, With = "Contoso" } }
        });
        Assert.True(preview.IsValid, string.Join("; ", preview.Errors.Select(e => e.Message)));
        Assert.NotEmpty(preview.Changes);

        using var described = JsonDocument.Parse(await tools.DescribeCapabilities());
        Assert.False(described.RootElement.GetProperty("renderingAvailable").GetBoolean());
    }

    [Fact]
    public async Task Registering_a_renderer_makes_discovery_report_it()
    {
        // Before 1.0 nothing set this flag, so discovery said false even with a renderer
        // registered, contradicting its documentation.
        using var services = Services(renderer: new LibreOfficeDocumentRenderer());
        var client = services.GetRequiredService<OfficeAgentClient>();
        Assert.True(client.DescribeCapabilities().RenderingAvailable);

        using var described = JsonDocument.Parse(await new OfficeAgentTools(client).DescribeCapabilities());
        Assert.True(described.RootElement.GetProperty("renderingAvailable").GetBoolean());
    }

    private static ServiceProvider Services(IDocumentRenderer? renderer, MemoryDocumentProvider? memory = null)
    {
        var services = new ServiceCollection();
        services.AddWordFormat();
        services.AddSingleton<IDocumentProvider>(memory ?? new MemoryDocumentProvider("memory"));
        if (renderer is not null) services.AddSingleton(renderer);
        services.AddOfficeAgent();
        return services.BuildServiceProvider();
    }
}
