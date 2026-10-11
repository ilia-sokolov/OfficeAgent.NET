using DocumentFormat.OpenXml.Packaging;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Excel;
using OfficeAgent.PowerPoint;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// Every Office package carries the same core properties in <c>docProps/core.xml</c>, so
/// <c>setProperty</c> on <c>core/*</c> works on decks and workbooks as it does on Word
/// documents (#30). Search indexes, SharePoint columns and Office's title bar show them.
/// </summary>
public class CorePropertyTests
{
    public static TheoryData<string> Formats => new() { "pptx", "xlsx" };

    [Theory]
    [MemberData(nameof(Formats))]
    public void Inspection_lists_the_core_properties(string format)
    {
        var (client, blank) = Setup(format);

        var paths = client.Inspect(blank).Nodes.Where(n => n.Kind == "docProperty").Select(n => n.Path).ToList();

        Assert.Contains("core/title", paths);
        Assert.Contains("core/creator", paths);
        Assert.Equal(8, paths.Count);
    }

    [Theory]
    [MemberData(nameof(Formats))]
    public void Setting_the_title_writes_it_to_the_package(string format)
    {
        var (client, blank) = Setup(format);

        using var applied = client.Commit(Handle(blank), Plan("core/title", "Q2 Board Review"));
        Assert.True(applied.Committed, string.Join("; ", applied.Report.Errors.Select(e => $"{e.Code}: {e.Message}")));

        var result = applied.ToBytes();
        Assert.Equal("Q2 Board Review", Title(format, result));
        Assert.Equal("Q2 Board Review",
            client.Inspect(result).Nodes.Single(n => n.Path == "core/title").Summary);
    }

    [Theory]
    [MemberData(nameof(Formats))]
    public void Preview_shows_the_value_before_and_after(string format)
    {
        var (client, blank) = Setup(format);
        using var titled = client.Commit(Handle(blank), Plan("core/title", "Draft"));

        var change = Assert.Single(client.Preview(Handle(titled.ToBytes()), Plan("core/title", "Final")).Changes);

        Assert.Equal("Draft", change.Before);
        Assert.Equal("Final", change.After);
    }

    [Theory]
    [MemberData(nameof(Formats))]
    public void An_unknown_property_is_refused_as_not_found(string format)
    {
        var (client, blank) = Setup(format);

        var report = client.Preview(Handle(blank), Plan("core/colour", "blue"));

        Assert.False(report.IsValid);
        Assert.Equal(ValidationErrorCodes.AnchorNotFound, Assert.Single(report.Errors).Code);
    }

    [Theory]
    [MemberData(nameof(Formats))]
    public void A_field_target_stays_word_only(string format)
    {
        var (client, blank) = Setup(format);

        var report = client.Preview(Handle(blank), new DocumentPlan
        {
            Operations = new PlanOperation[] { new SetPropertyOp { Target = new NodeAnchor { Kind = "field" }, Name = "updateOnOpen" } }
        });

        Assert.False(report.IsValid);
        Assert.Equal(ValidationErrorCodes.UnsupportedOperation, Assert.Single(report.Errors).Code);
    }

    [Fact]
    public void Discovery_advertises_set_property_for_every_format()
    {
        var capabilities = CapabilityDiscovery.Describe(
            new IFormatModule[] { new WordModule(), new PowerPointModule(), new ExcelModule() });

        Assert.All(capabilities.Formats, format =>
        {
            Assert.Contains("setProperty", format.Operations);
            Assert.Contains("docProperty", format.NodeKinds);
        });
    }

    private static (OfficeAgentClient Client, byte[] Blank) Setup(string format) => format switch
    {
        "pptx" => (new OfficeAgentClient(new PowerPointModule()), new PowerPointModule().CreateBlank()),
        _ => (new OfficeAgentClient(new ExcelModule()), new ExcelModule().CreateBlank())
    };

    private static DocumentPlan Plan(string path, string value) => new()
    {
        Operations = new PlanOperation[]
        {
            new SetPropertyOp { Target = new NodeAnchor { Kind = "docProperty", Path = path }, Value = value }
        }
    };

    private static string? Title(string format, byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using OpenXmlPackage package = format == "pptx"
            ? PresentationDocument.Open(stream, false)
            : SpreadsheetDocument.Open(stream, false);
        return package.PackageProperties.Title;
    }

    private static StreamHandle Handle(byte[] bytes) => new(new MemoryStream(bytes));
}
