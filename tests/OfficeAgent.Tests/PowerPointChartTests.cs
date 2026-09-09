using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.PowerPoint;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;
using DocFormat = OfficeAgent.Abstractions.DocumentFormat;

namespace OfficeAgent.Tests;

public class PowerPointChartTests
{
    [Theory]
    [InlineData(ChartKind.ClusteredColumn)]
    [InlineData(ChartKind.Bar)]
    [InlineData(ChartKind.Line)]
    [InlineData(ChartKind.Pie)]
    public void Insert_chart_creates_native_chart_and_embedded_workbook(ChartKind kind)
    {
        var module = new PowerPointModule();
        var client = new OfficeAgentClient(module);
        using var applied = client.Commit(Handle(module.CreateBlank()), Plan(new InsertChartOp
        {
            Target = new NodeAnchor { Kind = "slide", Path = "slide#256" },
            Kind = kind,
            Categories = new[] { "Q1", "Q2" },
            Series = new[] { new ChartSeries { Name = "Revenue", Values = new double?[] { 10, 12 } } },
            Title = "Revenue",
            Description = "Revenue by quarter"
        }));
        Assert.True(applied.Committed, string.Join("; ", applied.Report.Errors.Select(e => e.Message)));

        using var stream = new MemoryStream(applied.ToBytes());
        using var document = PresentationDocument.Open(stream, false);
        var slide = document.PresentationPart!.SlideParts.Single();
        var chart = Assert.Single(slide.ChartParts);
        Assert.NotNull(chart.ChartSpace);
        Assert.Single(chart.Parts.Select(p => p.OpenXmlPart).OfType<EmbeddedPackagePart>());
        Assert.Equal("Revenue by quarter", slide.Slide.Descendants<P.NonVisualDrawingProperties>()
            .Single(p => p.Name?.Value?.StartsWith("OfficeAgent Chart ") == true).Description?.Value);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(document).ToList();
        Assert.True(errors.Count == 0, string.Join("; ", errors.Select(e => e.Description)));
    }

    [Fact]
    public void Update_chart_rewrites_data_but_keeps_the_frame()
    {
        var module = new PowerPointModule();
        var client = new OfficeAgentClient(module);
        using var inserted = client.Commit(Handle(module.CreateBlank()), Plan(new InsertChartOp
        {
            Target = new NodeAnchor { Kind = "slide", Path = "slide#256" },
            Categories = new[] { "A" },
            Series = new[] { new ChartSeries { Name = "Old", Values = new double?[] { 1 } } }
        }));
        var chartPath = Assert.Single(client.Inspect(inserted.ToBytes()).Nodes, n => n.Kind == "chart").Path;
        using var updated = client.Commit(Handle(inserted.ToBytes()), Plan(new UpdateChartOp
        {
            Target = new NodeAnchor { Kind = "chart", Path = chartPath },
            Kind = ChartKind.Line,
            Categories = new[] { "A", "B" },
            Series = new[] { new ChartSeries { Name = "New", Values = new double?[] { 2, 3 } } },
            Description = "Updated chart"
        }));
        Assert.True(updated.Committed, string.Join("; ", updated.Report.Errors.Select(e => e.Message)));

        using var stream = new MemoryStream(updated.ToBytes());
        using var document = PresentationDocument.Open(stream, false);
        var slide = document.PresentationPart!.SlideParts.Single();
        Assert.Single(slide.Slide.Descendants<P.GraphicFrame>());
        Assert.Single(slide.ChartParts.Single().ChartSpace!.Descendants<C.LineChart>());
    }

    private static DocumentPlan Plan(PlanOperation operation) => new()
    {
        Format = DocFormat.PowerPoint,
        Operations = new[] { operation }
    };

    private static StreamHandle Handle(byte[] bytes) => new(new MemoryStream(bytes, writable: false));
}
