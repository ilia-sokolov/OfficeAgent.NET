using System.Diagnostics;
using System.Text.Json;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.Abstractions;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.PowerPoint;
using OfficeAgent.Word;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeAgent.Tests;

/// <summary>Resource bounds for agent-supplied regular expressions.</summary>
public sealed class RegexSafetyTests
{
    private const string PathologicalPattern = "(a+)+$";

    [Fact]
    public async Task Word_regex_timeout_is_a_stable_tool_error()
    {
        using var workspace = new RegexWorkspace();
        var document = await workspace.Client.RegisterBytesAsync(
            "workspace",
            workspace.Root,
            WordWithText(PathologicalText()),
            "pathological.docx");
        var tools = new OfficeAgentTools(workspace.Client);

        var stopwatch = Stopwatch.StartNew();
        var json = await tools.FindInDocument(
            "workspace",
            document.ItemId,
            PathologicalPattern,
            regex: true);
        stopwatch.Stop();

        using var parsed = JsonDocument.Parse(json);
        var error = Assert.Single(parsed.RootElement.GetProperty("errors").EnumerateArray());
        Assert.Equal("regex-timeout", error.GetProperty("code").GetString());
        Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(10), $"search took {stopwatch.Elapsed}");
    }

    [Fact]
    public void PowerPoint_regex_search_is_bounded()
    {
        var client = new OfficeAgentClient(new PowerPointModule());
        var handle = new StreamHandle(
            new MemoryStream(PowerPointWithText(PathologicalText())),
            "pathological.pptx");
        var query = new FindQuery
        {
            Pattern = PathologicalPattern,
            Options = new MatchOptions { Regex = true },
        };

        var stopwatch = Stopwatch.StartNew();
        Assert.Throws<RegexMatchTimeoutException>(() => client.Find(handle, query));
        stopwatch.Stop();

        Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(10), $"search took {stopwatch.Elapsed}");
    }

    private static string PathologicalText() => new string('a', 100_000) + "!";

    private static byte[] WordWithText(string text)
    {
        using var stream = EditableCopy(DocxFactory.Contract());
        using (var document = WordprocessingDocument.Open(stream, isEditable: true))
        {
            var main = document.MainDocumentPart
                ?? throw new InvalidDataException("The Word fixture has no main document part.");
            var root = main.Document
                ?? throw new InvalidDataException("The Word fixture has no document root.");
            var target = root.Descendants<Text>().First();
            target.Text = text;
            root.Save();
        }

        return stream.ToArray();
    }

    private static byte[] PowerPointWithText(string text)
    {
        using var stream = EditableCopy(PptxFactory.Deck());
        using (var document = PresentationDocument.Open(stream, isEditable: true))
        {
            var presentation = document.PresentationPart
                ?? throw new InvalidDataException("The PowerPoint fixture has no presentation part.");
            var slide = presentation.SlideParts.First().Slide
                ?? throw new InvalidDataException("The PowerPoint fixture has no slide root.");
            var target = slide.Descendants<A.Text>().First();
            target.Text = text;
            presentation.Presentation.Save();
        }

        return stream.ToArray();
    }

    private static MemoryStream EditableCopy(byte[] bytes)
    {
        var stream = new MemoryStream();
        stream.Write(bytes, 0, bytes.Length);
        stream.Position = 0;
        return stream;
    }

    private sealed class RegexWorkspace : IDisposable
    {
        private readonly ServiceProvider _services;

        public RegexWorkspace()
        {
            Root = Path.Combine(Path.GetTempPath(), $"officeagent-regex-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Root);
            _services = new ServiceCollection()
                .AddWordFormat()
                .AddFileSystemDocumentProvider("workspace", Root, options =>
                    options.AllowedExtensions = new[] { ".docx" })
                .AddOfficeAgent()
                .BuildServiceProvider();
            Client = _services.GetRequiredService<OfficeAgentClient>();
        }

        public string Root { get; }

        public OfficeAgentClient Client { get; }

        public void Dispose()
        {
            _services.Dispose();
            try
            {
                Directory.Delete(Root, recursive: true);
            }
            catch (IOException)
            {
                // A locked disposable fixture should not hide the assertion result.
            }
        }
    }
}
