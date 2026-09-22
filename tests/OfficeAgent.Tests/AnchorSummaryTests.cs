using System.Text.Json;
using OfficeAgent.Abstractions;
using OfficeAgent.AgentFramework;

namespace OfficeAgent.Tests;

/// <summary>
/// A change or error target names its anchor the way a plan does.
/// </summary>
/// <remarks>
/// Before 1.0 cell and shape anchors fell through to a fallback that reported the C# class name
/// (<c>kind: "CellAnchor"</c>) and dropped the sheet, cell and slide, so an agent reading a
/// change report could not tell which cell or shape had changed. Every concrete anchor type is
/// enumerated, so a new one that lands in the fallback fails here.
/// </remarks>
public sealed class AnchorSummaryTests
{
    public static IEnumerable<object[]> ConcreteAnchors() =>
        typeof(Anchor).Assembly.ExportedTypes
            .Where(type => typeof(Anchor).IsAssignableFrom(type) && !type.IsAbstract)
            .OrderBy(type => type.Name, StringComparer.Ordinal)
            .Select(type => new object[] { type });

    [Theory]
    [MemberData(nameof(ConcreteAnchors))]
    public void The_summary_kind_is_the_plan_discriminator(Type type)
    {
        var anchor = (Anchor)Activator.CreateInstance(type)!;
        using var plan = JsonDocument.Parse(JsonSerializer.Serialize(anchor, OfficeAgentTools.PlanJson));
        using var summary = JsonDocument.Parse(JsonSerializer.Serialize(OfficeAgentTools.SummariseAnchor(anchor), OfficeAgentTools.Json));

        Assert.Equal(plan.RootElement.GetProperty("$anchor").GetString(), summary.RootElement.GetProperty("kind").GetString());
    }

    [Fact]
    public void Cell_and_shape_targets_carry_what_locates_them()
    {
        using var cell = JsonDocument.Parse(JsonSerializer.Serialize(
            OfficeAgentTools.SummariseAnchor(new CellAnchor { SheetId = 2, Address = "B7" }), OfficeAgentTools.Json));
        Assert.Equal(2u, cell.RootElement.GetProperty("sheetId").GetUInt32());
        Assert.Equal("B7", cell.RootElement.GetProperty("address").GetString());

        using var shape = JsonDocument.Parse(JsonSerializer.Serialize(
            OfficeAgentTools.SummariseAnchor(new ShapeAnchor { SlideId = "256", ShapeId = "4" }), OfficeAgentTools.Json));
        Assert.Equal("256", shape.RootElement.GetProperty("slideId").GetString());
        Assert.Equal("4", shape.RootElement.GetProperty("shapeId").GetString());
    }
}
