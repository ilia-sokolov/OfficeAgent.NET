using OfficeAgent.Abstractions;
using OfficeAgent.Rendering;
using OfficeAgent.RenderWorker;

namespace OfficeAgent.Tests;

public sealed class RenderingTests
{
    [Fact]
    public async Task Out_of_process_renderer_returns_ordered_page_images()
    {
        var renderer = Renderer(pdfArguments: new[] { "--pages=2" });

        var result = await renderer.RenderAsync(
            new MemoryStream(new byte[] { 1, 2, 3 }), Options());

        Assert.True(result.Succeeded, result.Message);
        Assert.Equal(new[] { 1, 2 }, result.Pages.Select(page => page.PageNumber));
        Assert.All(result.Pages, page => Assert.Equal("image/png", page.ContentType));
    }

    [Fact]
    public async Task Renderer_stops_the_process_tree_at_the_page_limit()
    {
        var renderer = Renderer(pdfArguments: new[] { "--pages=2" });

        var result = await renderer.RenderAsync(
            new MemoryStream(new byte[] { 1 }), Options(maximumPages: 1));

        Assert.False(result.Succeeded);
        Assert.Equal("page-limit-exceeded", result.FailureCode);
        Assert.Empty(result.Pages);
    }

    [Fact]
    public async Task Renderer_stops_the_process_tree_at_the_total_timeout()
    {
        var renderer = Renderer(libreOfficeArguments: new[] { "--delay-ms=1000" });

        var result = await renderer.RenderAsync(
            new MemoryStream(new byte[] { 1 }), Options(timeout: TimeSpan.FromMilliseconds(250)));

        Assert.False(result.Succeeded);
        Assert.Equal("render-timeout", result.FailureCode);
        Assert.Empty(result.Pages);
    }

    [Fact]
    public async Task Renderer_rejects_output_that_exceeds_the_byte_limit()
    {
        var result = await Renderer().RenderAsync(
            new MemoryStream(new byte[] { 1 }),
            Options(maximumOutputBytes: 32));

        Assert.False(result.Succeeded);
        Assert.Equal("output-limit-exceeded", result.FailureCode);
        Assert.Empty(result.Pages);
    }

    [Fact]
    public async Task Renderer_stops_a_process_that_exceeds_its_working_set()
    {
        var renderer = Renderer(libreOfficeArguments: new[] { "--allocate-mb=64" });

        var result = await renderer.RenderAsync(
            new MemoryStream(new byte[] { 1 }),
            Options(maximumWorkingSetBytes: 32L * 1024 * 1024));

        Assert.False(result.Succeeded);
        Assert.Equal("memory-limit-exceeded", result.FailureCode);
        Assert.Empty(result.Pages);
    }

    [Fact]
    public void Structural_preview_has_no_renderer_dependency()
    {
        Assert.DoesNotContain(
            typeof(IDocumentRenderer),
            typeof(OfficeAgent.Core.OfficeAgentClient).GetConstructors()
                .SelectMany(constructor => constructor.GetParameters())
                .Select(parameter => parameter.ParameterType));
    }

    private static LibreOfficeDocumentRenderer Renderer(
        IEnumerable<string>? libreOfficeArguments = null,
        IEnumerable<string>? pdfArguments = null)
    {
        var worker = typeof(RenderWorkerMarker).Assembly.Location;
        return new LibreOfficeDocumentRenderer(new LibreOfficeRendererOptions
        {
            LibreOfficeExecutable = "dotnet",
            LibreOfficePrefixArguments = new[] { worker }.Concat(libreOfficeArguments ?? Array.Empty<string>()).ToList(),
            PdfToPpmExecutable = "dotnet",
            PdfToPpmPrefixArguments = new[] { worker }.Concat(pdfArguments ?? Array.Empty<string>()).ToList()
        });
    }

    private static RenderOptions Options(
        int maximumPages = 10,
        TimeSpan? timeout = null,
        long maximumOutputBytes = 1024 * 1024,
        long maximumWorkingSetBytes = 256L * 1024 * 1024) => new()
    {
        FileName = "sample.docx",
        MaximumPages = maximumPages,
        Timeout = timeout ?? TimeSpan.FromSeconds(10),
        MaximumInputBytes = 1024,
        MaximumOutputBytes = maximumOutputBytes,
        MaximumWorkingSetBytes = maximumWorkingSetBytes
    };
}
