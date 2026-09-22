using OfficeAgent.Abstractions;
using OfficeAgent.Rendering;
using OfficeAgent.RenderWorker;
using System.Diagnostics;

namespace OfficeAgent.Tests;

public sealed class RenderingTests
{
    [Fact]
    public async Task Out_of_process_renderer_returns_ordered_page_images()
    {
        var renderer = Renderer(pdfArguments: new[] { "--pages=2" });

        var result = await renderer.RenderAsync(
            new MemoryStream(Package()), Options());

        Assert.True(result.Succeeded, result.Message);
        Assert.Equal(new[] { 1, 2 }, result.Pages.Select(page => page.PageNumber));
        Assert.All(result.Pages, page => Assert.Equal("image/png", page.ContentType));
    }

    [Fact]
    public async Task Renderer_stops_the_process_tree_at_the_page_limit()
    {
        var renderer = Renderer(pdfArguments: new[] { "--pages=2" });

        var result = await renderer.RenderAsync(
            new MemoryStream(Package()), Options(maximumPages: 1));

        Assert.False(result.Succeeded);
        Assert.Equal("page-limit-exceeded", result.FailureCode);
        AssertFailsClosed(result);
    }

    [Fact]
    public async Task Renderer_stops_the_process_tree_at_the_total_timeout()
    {
        var renderer = Renderer(libreOfficeArguments: new[] { "--delay-ms=1000" });

        var result = await renderer.RenderAsync(
            new MemoryStream(Package()), Options(timeout: TimeSpan.FromMilliseconds(250)));

        Assert.False(result.Succeeded);
        Assert.Equal("render-timeout", result.FailureCode);
        AssertFailsClosed(result);
    }

    [Fact]
    public async Task Renderer_timeout_includes_output_pipe_draining()
    {
        var renderer = Renderer(libreOfficeArguments: new[] { "--spawn-child-holding-pipe" });
        var stopwatch = Stopwatch.StartNew();

        var result = await renderer.RenderAsync(
            new MemoryStream(Package()), Options(timeout: TimeSpan.FromMilliseconds(250)));

        Assert.False(result.Succeeded);
        Assert.Equal("render-timeout", result.FailureCode);
        Assert.InRange(stopwatch.Elapsed, TimeSpan.Zero, TimeSpan.FromSeconds(4));
    }

    [Fact]
    public async Task Renderer_rejects_output_that_exceeds_the_byte_limit()
    {
        var result = await Renderer().RenderAsync(
            new MemoryStream(Package()),
            Options(maximumOutputBytes: 32));

        Assert.False(result.Succeeded);
        Assert.Equal("output-limit-exceeded", result.FailureCode);
        AssertFailsClosed(result);
    }

    [Fact]
    public async Task Renderer_stops_a_process_that_exceeds_its_working_set()
    {
        var renderer = Renderer(libreOfficeArguments: new[] { "--allocate-mb=64" });

        var result = await renderer.RenderAsync(
            new MemoryStream(Package()),
            Options(timeout: TimeSpan.FromSeconds(5), maximumWorkingSetBytes: 32L * 1024 * 1024));

        Assert.False(result.Succeeded);
        Assert.Equal("memory-limit-exceeded", result.FailureCode);
        AssertFailsClosed(result);
    }

    [Fact]
    public async Task Missing_renderer_executable_fails_a_page_gate_closed()
    {
        // The renderer is configured to a command that does not exist, as on a container built
        // without LibreOffice. A page-count gate written the obvious way must not read that as
        // a zero-page document and admit the contract.
        var renderer = new LibreOfficeDocumentRenderer(new LibreOfficeRendererOptions
        {
            LibreOfficeExecutable = "officeagent-no-such-renderer",
            PdfToPpmExecutable = "officeagent-no-such-rasterizer"
        });

        var result = await renderer.RenderAsync(new MemoryStream(Package()), Options());

        Assert.False(result.Succeeded);
        Assert.Equal("renderer-unavailable", result.FailureCode);

        var thrown = Assert.Throws<RenderFailedException>(() => result.PageCount);
        Assert.Equal("renderer-unavailable", thrown.FailureCode);
        Assert.False(WithinPageBudget(result, budget: 4));
    }

    [Fact]
    public void A_failed_result_never_reports_pages_and_a_successful_one_needs_no_check()
    {
        var failure = RenderResult.Failure("renderer-failed", "A renderer process exited unsuccessfully.");

        Assert.Throws<RenderFailedException>(() => failure.Pages);
        Assert.Throws<RenderFailedException>(() => failure.PageCount);
        Assert.Throws<RenderFailedException>(() => failure.EnsureSucceeded());
        Assert.False(failure.TryGetPages(out var none));
        Assert.Empty(none);

        var success = RenderResult.Success(new[] { new RenderedPage { PageNumber = 1 } });

        Assert.Equal(1, success.PageCount);
        Assert.True(success.TryGetPages(out var pages));
        Assert.Single(pages);
        Assert.Null(success.FailureCode);
        success.EnsureSucceeded();
    }

    [Fact]
    public void A_failed_result_can_still_be_serialized_for_a_log()
    {
        // Logging a failure must not trip over the guarded members, and page bytes never
        // belong in a log line.
        var json = System.Text.Json.JsonSerializer.Serialize(
            RenderResult.Failure("renderer-unavailable", "A configured renderer executable was not found."));

        Assert.Contains("renderer-unavailable", json);
        Assert.DoesNotContain("Pages", json);
        Assert.DoesNotContain("PageCount", json);
    }

    [Fact]
    public void A_failure_cannot_be_constructed_without_a_code()
    {
        Assert.Throws<ArgumentException>(() => RenderResult.Failure("", "missing code"));
        Assert.Throws<ArgumentException>(() => RenderResult.Failure("renderer-failed", " "));
        Assert.Throws<ArgumentNullException>(() => RenderResult.Success(null!));
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

    /// <summary>A failed render exposes no pages by any route.</summary>
    private static void AssertFailsClosed(RenderResult result)
    {
        Assert.False(result.Succeeded);
        Assert.Throws<RenderFailedException>(() => result.Pages);
        Assert.False(result.TryGetPages(out var pages));
        Assert.Empty(pages);
    }

    /// <summary>A page gate written the way a caller would write it.</summary>
    private static bool WithinPageBudget(RenderResult result, int budget)
    {
        if (!result.TryGetPages(out var pages)) return false;
        return pages.Count <= budget;
    }

    [Fact]
    public async Task Content_that_is_not_the_declared_package_is_refused_before_any_backend_runs()
    {
        // No backend exists at these paths: reaching one would report renderer-unavailable, so
        // renderer-failed proves the refusal came first. LibreOffice imports by content, and a
        // .docx name alone would hand any of these to one of its parsers.
        var renderer = new LibreOfficeDocumentRenderer(new LibreOfficeRendererOptions
        {
            LibreOfficeExecutable = "officeagent-no-such-renderer",
            PdfToPpmExecutable = "officeagent-no-such-rasterizer"
        });
        var samples = new (string Name, byte[] Bytes)[]
        {
            ("plain text", "not a package"u8.ToArray()),
            ("a zip with no content types", Zip(("word/document.xml", "<w:document/>"))),
            ("a deck under a .docx name", Package("application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml")),
            ("a macro-enabled main part", Package("application/vnd.ms-word.document.macroEnabled.main+xml")),
            ("a VBA project beside a normal main part", Package(extra: ("word/vbaProject.bin", "x"))),
        };

        foreach (var (name, bytes) in samples)
        {
            var result = await renderer.RenderAsync(new MemoryStream(bytes), Options());
            Assert.True(result.FailureCode == "renderer-failed", $"{name}: {result.FailureCode} {result.Message}");
        }

        var genuine = await renderer.RenderAsync(new MemoryStream(Package()), Options());
        Assert.Equal("renderer-unavailable", genuine.FailureCode);
    }

    /// <summary>A minimal package declaring a Word main part, small enough for the test input limit.</summary>
    private static byte[] Package(
        string mainContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
        (string Name, string Content)? extra = null)
    {
        var entries = new List<(string, string)>
        {
            ("[Content_Types].xml",
             "<?xml version=\"1.0\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
             $"<Override PartName=\"/word/document.xml\" ContentType=\"{mainContentType}\"/></Types>"),
            ("word/document.xml", "<w:document/>")
        };
        if (extra is { } added) entries.Add(added);
        return Zip(entries.ToArray());
    }

    private static byte[] Zip(params (string Name, string Content)[] entries)
    {
        using var stream = new MemoryStream();
        using (var zip = new System.IO.Compression.ZipArchive(stream, System.IO.Compression.ZipArchiveMode.Create, leaveOpen: true))
            foreach (var (name, content) in entries)
            {
                using var writer = new StreamWriter(zip.CreateEntry(name).Open());
                writer.Write(content);
            }
        return stream.ToArray();
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
