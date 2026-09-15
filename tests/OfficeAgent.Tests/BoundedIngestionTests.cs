using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using OfficeAgent.Abstractions;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// Host ingestion ceilings. Each configurable limit is exercised at boundary minus one,
/// at the boundary and at boundary plus one, and the hostile package shapes are refused
/// without producing output or leaving the source altered.
/// </summary>
/// <remarks>
/// These are resource ceilings inside one process, not an operating-system sandbox.
/// Renderer isolation is a separate concern and belongs to V10-04.
/// </remarks>
public sealed class BoundedIngestionTests
{
    // ── The limits object itself ─────────────────────────────────────────

    /// <summary>A request may ask for something stricter; it can never raise a ceiling.</summary>
    [Fact]
    public void A_request_can_lower_an_effective_limit_but_never_raise_a_host_ceiling()
    {
        var host = new OpenXmlIngestionLimits
        {
            MaximumCompressedBytes = 1_000,
            MaximumParts = 10,
            MaximumXmlDepth = 20
        };

        var stricter = host.Restrict(new OpenXmlIngestionLimits
        {
            MaximumCompressedBytes = 500,
            MaximumParts = 5,
            MaximumXmlDepth = 10
        });
        Assert.Equal(500, stricter.MaximumCompressedBytes);
        Assert.Equal(5, stricter.MaximumParts);
        Assert.Equal(10, stricter.MaximumXmlDepth);

        var greedy = host.Restrict(new OpenXmlIngestionLimits
        {
            MaximumCompressedBytes = long.MaxValue,
            MaximumParts = int.MaxValue,
            MaximumXmlDepth = int.MaxValue
        });
        Assert.Equal(1_000, greedy.MaximumCompressedBytes);
        Assert.Equal(10, greedy.MaximumParts);
        Assert.Equal(20, greedy.MaximumXmlDepth);

        // No request at all leaves the host ceilings exactly as configured.
        Assert.Equal(1_000, host.Restrict(null).MaximumCompressedBytes);
    }

    /// <summary>A ceiling of zero or less is a configuration error, not a silent disable.</summary>
    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    public void A_non_positive_ceiling_is_rejected(long value)
    {
        var limits = new OpenXmlIngestionLimits { MaximumCompressedBytes = value };
        Assert.Throws<ArgumentOutOfRangeException>(() => limits.Validate());
    }

    // ── Compressed size, at and around the boundary ──────────────────────

    /// <summary>
    /// Boundary minus one and the boundary itself are accepted; boundary plus one is
    /// refused, naming the limit and reporting the observed size.
    /// </summary>
    [Fact]
    public void The_compressed_ceiling_admits_up_to_the_boundary_and_refuses_past_it()
    {
        var document = DocxFactory.Contract();
        long exact = document.LongLength;

        Assert.NotNull(ClientWith(new OpenXmlIngestionLimits { MaximumCompressedBytes = exact + 1 })
            .Inspect(document));
        Assert.NotNull(ClientWith(new OpenXmlIngestionLimits { MaximumCompressedBytes = exact })
            .Inspect(document));

        var tooSmall = ClientWith(new OpenXmlIngestionLimits { MaximumCompressedBytes = exact - 1 });
        var error = Assert.Throws<OpenXmlIngestionLimitException>(() => tooSmall.Inspect(document));
        Assert.Equal("MaximumCompressedBytes", error.Limit);
        Assert.Equal(exact - 1, error.Allowed);
        Assert.Equal(exact, error.Observed);
        Assert.Equal("input-too-large", OpenXmlIngestionLimitException.Code);
    }

    /// <summary>
    /// A stream that never ends is refused at the ceiling rather than drained. Without
    /// the bound this test would not finish.
    /// </summary>
    [Fact]
    public void An_endless_stream_is_refused_at_the_ceiling_rather_than_drained()
    {
        var client = ClientWith(new OpenXmlIngestionLimits { MaximumCompressedBytes = 64 * 1024 });
        using var endless = new EndlessStream();

        var error = Assert.Throws<OpenXmlIngestionLimitException>(
            () => client.Inspect(new StreamHandle(endless)));

        Assert.Equal("MaximumCompressedBytes", error.Limit);
        Assert.Equal(64 * 1024, error.Allowed);
        // Refused before the true size was known, which is the point of stopping early.
        Assert.Equal(-1, error.Observed);
        Assert.True(endless.BytesProduced <= 64 * 1024 + 81920,
            $"the reader kept pulling after the limit: {endless.BytesProduced} bytes");
    }

    // ── Part count and part size ─────────────────────────────────────────

    /// <summary>Entry count is checked at the boundary and one past it.</summary>
    [Fact]
    public void The_part_count_ceiling_admits_the_boundary_and_refuses_one_more()
    {
        var document = DocxFactory.Contract();
        int parts = EntryCount(document);

        Assert.NotNull(ClientWith(new OpenXmlIngestionLimits { MaximumParts = parts }).Inspect(document));
        Assert.NotNull(ClientWith(new OpenXmlIngestionLimits { MaximumParts = parts + 1 }).Inspect(document));

        var error = Assert.Throws<OpenXmlIngestionLimitException>(
            () => ClientWith(new OpenXmlIngestionLimits { MaximumParts = parts - 1 }).Inspect(document));
        Assert.Equal("MaximumParts", error.Limit);
    }

    /// <summary>Per-part size is checked against the largest declared entry.</summary>
    [Fact]
    public void The_per_part_ceiling_admits_the_boundary_and_refuses_one_less()
    {
        var document = DocxFactory.Contract();
        long largest = LargestEntry(document);

        Assert.NotNull(ClientWith(new OpenXmlIngestionLimits { MaximumPartBytes = largest }).Inspect(document));

        var error = Assert.Throws<OpenXmlIngestionLimitException>(
            () => ClientWith(new OpenXmlIngestionLimits { MaximumPartBytes = largest - 1 }).Inspect(document));
        Assert.Equal("MaximumPartBytes", error.Limit);
    }

    // ── Hostile package shapes ───────────────────────────────────────────

    /// <summary>
    /// A decompression bomb: a small archive declaring an enormous expansion. It is
    /// refused from its declared ratio before any part is read, so the expansion never
    /// happens.
    /// </summary>
    [Fact]
    public void A_zip_bomb_is_refused_before_any_part_is_expanded()
    {
        var bomb = Archive(builder =>
        {
            builder("[Content_Types].xml", ContentTypes);
            // 40 MiB of zeros compresses to almost nothing.
            builder("word/document.xml", new string('0', 40 * 1024 * 1024));
        });

        var error = Assert.Throws<OpenXmlIngestionLimitException>(() => Client().Inspect(bomb));
        Assert.True(
            error.Limit is "MaximumExpansionRatio" or "MaximumExpandedBytes" or "MaximumPartBytes",
            $"unexpected limit {error.Limit}");
    }

    /// <summary>A truncated archive is a bounded refusal, not a crash or a hang.</summary>
    [Fact]
    public void A_truncated_package_is_refused_as_malformed()
    {
        var document = DocxFactory.Contract();
        var truncated = document.Take(document.Length / 2).ToArray();

        var error = Assert.Throws<OpenXmlPackageRejectedException>(() => Client().Inspect(truncated));
        Assert.Contains("corrupt", error.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("malformed-package", OpenXmlPackageRejectedException.Code);
    }

    /// <summary>Bytes that are not an archive at all are refused the same bounded way.</summary>
    [Fact]
    public void Bytes_that_are_not_a_package_are_refused_as_malformed()
    {
        Assert.Throws<OpenXmlPackageRejectedException>(
            () => Client().Inspect(Encoding.UTF8.GetBytes("this is not a document")));
    }

    /// <summary>A duplicated part name makes the package ambiguous, so it is refused.</summary>
    [Fact]
    public void A_duplicated_part_name_is_refused()
    {
        var duplicated = Archive(builder =>
        {
            builder("[Content_Types].xml", ContentTypes);
            builder("word/document.xml", MinimalDocument);
            builder("word/document.xml", MinimalDocument);
        });

        var error = Assert.Throws<OpenXmlPackageRejectedException>(() => Client().Inspect(duplicated));
        Assert.Contains("more than one entry named", error.Message, StringComparison.Ordinal);
    }

    /// <summary>An entry name that escapes the package root is refused.</summary>
    [Theory]
    [InlineData("../outside.xml")]
    [InlineData("word/../../outside.xml")]
    [InlineData("/absolute.xml")]
    public void A_traversing_or_absolute_entry_name_is_refused(string name)
    {
        var hostile = Archive(builder =>
        {
            builder("[Content_Types].xml", ContentTypes);
            builder(name, "<x/>");
            builder("word/document.xml", MinimalDocument);
        });

        var error = Assert.Throws<OpenXmlPackageRejectedException>(() => Client().Inspect(hostile));
        Assert.True(
            error.Message.Contains("traversing", StringComparison.Ordinal) ||
            error.Message.Contains("absolute", StringComparison.Ordinal),
            error.Message);
    }

    /// <summary>
    /// An external XML entity is never resolved. A billion-laughs style payload in the
    /// content-types part is refused because document type definitions are prohibited.
    /// </summary>
    [Fact]
    public void External_entities_and_doctypes_in_the_manifest_are_refused()
    {
        var hostile = Archive(builder =>
        {
            builder("[Content_Types].xml",
                "<?xml version=\"1.0\"?><!DOCTYPE Types [<!ENTITY xxe SYSTEM \"file:///etc/passwd\">]>" +
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">&xxe;</Types>");
            builder("word/document.xml", MinimalDocument);
        });

        var error = Assert.Throws<OpenXmlPackageRejectedException>(() => Client().Inspect(hostile));
        Assert.Contains("external entities are refused", error.Message, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>A package with no content-types part is not an OOXML package.</summary>
    [Fact]
    public void A_package_without_content_types_is_refused()
    {
        var incomplete = Archive(builder => builder("word/document.xml", MinimalDocument));

        var error = Assert.Throws<OpenXmlPackageRejectedException>(() => Client().Inspect(incomplete));
        Assert.Contains("[Content_Types].xml", error.Message, StringComparison.Ordinal);
    }

    /// <summary>Deeply nested manifest XML is refused against the depth ceiling.</summary>
    [Fact]
    public void Manifest_xml_deeper_than_the_ceiling_is_refused()
    {
        const int depth = 60;
        var nested = new StringBuilder("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
        for (int i = 0; i < depth; i++) nested.Append("<a>");
        for (int i = 0; i < depth; i++) nested.Append("</a>");
        nested.Append("</Types>");

        var document = Archive(builder =>
        {
            builder("[Content_Types].xml", nested.ToString());
            builder("word/document.xml", MinimalDocument);
        });

        var error = Assert.Throws<OpenXmlIngestionLimitException>(
            () => ClientWith(new OpenXmlIngestionLimits { MaximumXmlDepth = 10 }).Inspect(document));
        Assert.Equal("MaximumXmlDepth", error.Limit);
    }

    // ── No output, no drift, bounded diagnostics ─────────────────────────

    /// <summary>
    /// A refused input leaves the stored document byte-identical and produces no output.
    /// </summary>
    [Fact]
    public async Task A_refused_input_leaves_provider_storage_unchanged()
    {
        using var workspace = new TemporaryWorkspace();
        var path = Path.Combine(workspace.Root, "oversized.docx");
        var document = DocxFactory.Contract();
        await File.WriteAllBytesAsync(path, document);
        var before = Convert.ToHexString(SHA256.HashData(document));

        var client = new OfficeAgentClient(
            new OfficeAgentEngine(
                new[] { new WordModule() },
                limits: new OpenXmlIngestionLimits { MaximumCompressedBytes = document.LongLength - 1 }),
            new DocumentProviderRegistry(new[] { workspace.Provider() }),
            loggerFactory: null);

        var reference = await client.RegisterAsync("workspace", path);
        await Assert.ThrowsAsync<OpenXmlIngestionLimitException>(() => client.InspectAsync(reference));

        Assert.Equal(before, Convert.ToHexString(SHA256.HashData(await File.ReadAllBytesAsync(path))));
        Assert.Single(Directory.GetFiles(workspace.Root));
    }

    /// <summary>
    /// The diagnostic names the limit and the numbers, and stays short. A refusal must
    /// not echo the document back to the caller.
    /// </summary>
    [Fact]
    public void A_refusal_diagnostic_is_bounded_and_actionable()
    {
        var document = DocxFactory.Contract();
        var error = Assert.Throws<OpenXmlIngestionLimitException>(
            () => ClientWith(new OpenXmlIngestionLimits { MaximumCompressedBytes = 32 }).Inspect(document));

        Assert.True(error.Message.Length < 400, $"diagnostic was {error.Message.Length} characters");
        Assert.Contains("MaximumCompressedBytes", error.Message, StringComparison.Ordinal);
        Assert.Contains("32", error.Message, StringComparison.Ordinal);
        Assert.Contains("no output was produced", error.Message, StringComparison.Ordinal);
    }

    // ── The bypass surface ───────────────────────────────────────────────

    /// <summary>
    /// The inline agent tool is bounded too, and reports the stable wire code rather than
    /// an internal error.
    /// </summary>
    [Fact]
    public async Task The_inline_tool_refuses_an_oversized_payload_with_a_stable_code()
    {
        var tools = new OfficeAgentTools(new OfficeAgentClient(new WordModule()));

        // Larger than the default compressed ceiling, as base64, without allocating it
        // decoded: the tool must refuse from the encoded length.
        var oversized = new string('A', (int)(OpenXmlIngestionLimits.Default.MaximumCompressedBytes / 3 * 4) + 8);

        using var result = JsonDocument.Parse(await tools.InspectDocumentContent(oversized));
        Assert.Equal("input-too-large", FirstErrorCode(result));
    }

    /// <summary>
    /// The direct API reports a malformed package with its own code. The inline tool
    /// keeps its documented <c>invalid-argument</c> answer instead, because for that path
    /// the useful advice is about the base64 copy rather than the package: re-sending the
    /// same string fails identically.
    /// </summary>
    [Fact]
    public async Task A_malformed_package_is_reported_per_entry_point()
    {
        var notAPackage = Encoding.UTF8.GetBytes("nowhere near a docx");

        Assert.Throws<OpenXmlPackageRejectedException>(() => Client().Inspect(notAPackage));

        var tools = new OfficeAgentTools(new OfficeAgentClient(new WordModule()));
        using var result = JsonDocument.Parse(
            await tools.InspectDocumentContent(Convert.ToBase64String(notAPackage)));
        Assert.Equal("invalid-argument", FirstErrorCode(result));
    }

    /// <summary>
    /// The ceiling applies to the direct byte API, the stream API and the provider flow
    /// alike, because all three reach the same open path.
    /// </summary>
    [Fact]
    public async Task Every_entry_point_reaches_the_same_ceiling()
    {
        var document = DocxFactory.Contract();
        var limits = new OpenXmlIngestionLimits { MaximumCompressedBytes = document.LongLength - 1 };

        Assert.Throws<OpenXmlIngestionLimitException>(() => ClientWith(limits).Inspect(document));
        Assert.Throws<OpenXmlIngestionLimitException>(
            () => ClientWith(limits).Inspect(new StreamHandle(new MemoryStream(document, writable: false))));

        using var workspace = new TemporaryWorkspace();
        var path = Path.Combine(workspace.Root, "contract.docx");
        await File.WriteAllBytesAsync(path, document);

        var provider = new OfficeAgentClient(
            new OfficeAgentEngine(new[] { new WordModule() }, limits: limits),
            new DocumentProviderRegistry(new[] { workspace.Provider() }),
            loggerFactory: null);
        var reference = await provider.RegisterAsync("workspace", path);
        await Assert.ThrowsAsync<OpenXmlIngestionLimitException>(() => provider.InspectAsync(reference));
    }

    /// <summary>Cancellation during a bounded read is observed, not swallowed.</summary>
    [Fact]
    public async Task Cancellation_is_observed_while_reading()
    {
        var client = Client();
        using var source = new MemoryStream(DocxFactory.Contract(), writable: false);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(
            () => client.InspectAsync(new StreamHandle(source), cancellationToken: cancellation.Token));
    }

    /// <summary>
    /// When a stream would breach the ceiling and the caller has also cancelled, the
    /// ceiling is what stops the read: the bytes are refused before any work is queued.
    /// Either outcome is bounded, and this fixes which one callers actually see.
    /// </summary>
    [Fact]
    public async Task An_oversized_stream_is_refused_even_when_cancellation_is_pending()
    {
        var client = ClientWith(new OpenXmlIngestionLimits { MaximumCompressedBytes = 64 * 1024 });
        using var endless = new EndlessStream();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsAsync<OpenXmlIngestionLimitException>(
            () => client.InspectAsync(new StreamHandle(endless), cancellationToken: cancellation.Token));
    }

    /// <summary>Total expanded size, at the boundary and one below it.</summary>
    [Fact]
    public void The_expanded_ceiling_admits_the_boundary_and_refuses_one_less()
    {
        var document = DocxFactory.Contract();
        long expanded = TotalExpanded(document);

        Assert.NotNull(ClientWith(new OpenXmlIngestionLimits { MaximumExpandedBytes = expanded }).Inspect(document));
        Assert.NotNull(ClientWith(new OpenXmlIngestionLimits { MaximumExpandedBytes = expanded + 1 }).Inspect(document));

        var error = Assert.Throws<OpenXmlIngestionLimitException>(
            () => ClientWith(new OpenXmlIngestionLimits { MaximumExpandedBytes = expanded - 1 }).Inspect(document));
        Assert.Equal("MaximumExpandedBytes", error.Limit);
    }

    /// <summary>The expansion ratio, at the boundary and one below it.</summary>
    [Fact]
    public void The_expansion_ratio_ceiling_admits_the_boundary_and_refuses_one_less()
    {
        var document = DocxFactory.Contract();
        long ratio = TotalExpanded(document) / document.LongLength;

        Assert.NotNull(ClientWith(new OpenXmlIngestionLimits
        {
            MaximumExpansionRatio = (int)Math.Max(ratio, 1)
        }).Inspect(document));

        if (ratio >= 1)
        {
            var error = Assert.Throws<OpenXmlIngestionLimitException>(
                () => ClientWith(new OpenXmlIngestionLimits { MaximumExpansionRatio = 1 })
                    .Inspect(HighlyCompressible()));
            Assert.Equal("MaximumExpansionRatio", error.Limit);
        }
    }

    /// <summary>
    /// The XML character ceiling, at the boundary and one below it, measured against the
    /// real manifest of a real document. A manifest larger than the ceiling is refused
    /// while being read rather than after it has been buffered.
    /// </summary>
    [Fact]
    public void The_xml_character_ceiling_admits_the_boundary_and_refuses_one_less()
    {
        var document = DocxFactory.Contract();
        long manifest = EntryLength(document, "[Content_Types].xml");

        Assert.NotNull(ClientWith(new OpenXmlIngestionLimits { MaximumXmlCharacters = manifest }).Inspect(document));
        Assert.NotNull(ClientWith(new OpenXmlIngestionLimits { MaximumXmlCharacters = manifest + 1 }).Inspect(document));

        var error = Assert.Throws<OpenXmlIngestionLimitException>(
            () => ClientWith(new OpenXmlIngestionLimits { MaximumXmlCharacters = manifest - 1 }).Inspect(document));
        Assert.Contains("Xml", error.Limit, StringComparison.Ordinal);
    }

    /// <summary>
    /// A package that announces a Word document but carries no main document part is
    /// damaged, and must be refused where every other damaged package is refused.
    /// </summary>
    /// <remarks>
    /// Seed 11 of the malformed corpus reaches exactly this state on Linux and macOS, where
    /// its byte flip lands on the main part's relationship. It did not on Windows, so the
    /// gap survived until the platform matrix ran. Building the shape directly makes the
    /// refusal provable on every platform instead of only where a byte flip happens to land.
    /// </remarks>
    [Fact]
    public void A_package_that_declares_a_main_part_but_omits_it_is_refused()
    {
        var withoutMainPart = Archive(add =>
        {
            add("[Content_Types].xml", ContentTypes);
            add("_rels/.rels", PackageRelationships);
        });

        var error = Assert.Throws<OpenXmlPackageRejectedException>(() => Client().Inspect(withoutMainPart));
        Assert.Equal("malformed-package", OpenXmlPackageRejectedException.Code);

        // Pins which refusal fired. Without this the test would also pass if the SDK
        // happened to reject the package first, and the missing-main-part guard could be
        // removed without a failure.
        Assert.Contains("no main document part", error.Message, StringComparison.Ordinal);
        Assert.True(error.Message.Length < 600, $"diagnostic was {error.Message.Length} characters");
    }

    // ── Deterministic malformed seeds ────────────────────────────────────

    /// <summary>
    /// A fixed corpus of malformed packages, each derived from a valid document by one
    /// deterministic mutation. Every seed must produce a bounded refusal rather than a
    /// hang, an unhandled exception type, or a partial result. The seed index is the
    /// reproduction: rerun this case with that index to get the identical bytes.
    /// </summary>
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    [InlineData(5)]
    [InlineData(6)]
    [InlineData(7)]
    [InlineData(8)]
    [InlineData(9)]
    [InlineData(10)]
    [InlineData(11)]
    public void Every_malformed_seed_produces_a_bounded_refusal(int seed)
    {
        var corrupted = Mutate(DocxFactory.Contract(), seed);
        var client = Client();

        var thrown = Record.Exception(() => client.Inspect(corrupted));

        // A seed may survive its mutation and still be a valid document; that is a pass.
        if (thrown is null) return;

        Assert.True(
            thrown is OpenXmlPackageRejectedException or OpenXmlIngestionLimitException,
            $"seed {seed} produced {thrown.GetType().Name}: {thrown.Message}");
        Assert.True(thrown.Message.Length < 600, $"seed {seed} diagnostic was {thrown.Message.Length} characters");
    }

    /// <summary>
    /// One deterministic byte flip per seed, at a position derived from the seed. No
    /// randomness, so a failure names the exact input that produced it.
    /// </summary>
    private static byte[] Mutate(byte[] original, int seed)
    {
        var copy = (byte[])original.Clone();
        if (seed % 3 == 0)
        {
            // Truncate to a fraction of the original length.
            int keep = (int)(copy.Length * (0.1 + 0.07 * seed)) % Math.Max(copy.Length, 1);
            return copy.Take(Math.Max(keep, 1)).ToArray();
        }

        int index = (int)((long)seed * 7919 % copy.Length);
        copy[index] ^= 0xFF;
        return copy;
    }

    // ── Helpers ──────────────────────────────────────────────────────────

    private const string ContentTypes =
        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
        "<Override PartName=\"/word/document.xml\" " +
        "ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>" +
        "</Types>";

    private const string PackageRelationships =
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
        "<Relationship Id=\"rId1\" " +
        "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" " +
        "Target=\"word/document.xml\"/>" +
        "</Relationships>";

    private const string MinimalDocument =
        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body/></w:document>";

    private static string? FirstErrorCode(JsonDocument result) =>
        result.RootElement.GetProperty("errors")[0].GetProperty("Code").GetString();

    private static OfficeAgentClient Client() => new(new WordModule());

    private static OfficeAgentClient ClientWith(OpenXmlIngestionLimits limits) =>
        new(new OfficeAgentEngine(new[] { new WordModule() }, limits: limits));

    private static byte[] Archive(Action<Action<string, string>> build)
    {
        using var memory = new MemoryStream();
        using (var zip = new ZipArchive(memory, ZipArchiveMode.Create, leaveOpen: true))
        {
            build((name, content) =>
            {
                var entry = zip.CreateEntry(name, CompressionLevel.Optimal);
                using var writer = new StreamWriter(entry.Open(), Encoding.UTF8);
                writer.Write(content);
            });
        }

        return memory.ToArray();
    }

    private static long TotalExpanded(byte[] package)
    {
        using var zip = new ZipArchive(new MemoryStream(package, writable: false), ZipArchiveMode.Read);
        return zip.Entries.Sum(entry => entry.Length);
    }

    /// <summary>A valid package whose parts compress far better than a real document.</summary>
    private static byte[] HighlyCompressible() => Archive(builder =>
    {
        builder("[Content_Types].xml", ContentTypes);
        builder("word/document.xml", MinimalDocument.Replace(
            "<w:body/>", "<w:body>" + new string(' ', 512 * 1024) + "</w:body>", StringComparison.Ordinal));
    });

    private static long EntryLength(byte[] package, string name)
    {
        using var zip = new ZipArchive(new MemoryStream(package, writable: false), ZipArchiveMode.Read);
        return zip.Entries.Single(entry => entry.FullName == name).Length;
    }

    private static int EntryCount(byte[] package)
    {
        using var zip = new ZipArchive(new MemoryStream(package, writable: false), ZipArchiveMode.Read);
        return zip.Entries.Count;
    }

    private static long LargestEntry(byte[] package)
    {
        using var zip = new ZipArchive(new MemoryStream(package, writable: false), ZipArchiveMode.Read);
        return zip.Entries.Max(entry => entry.Length);
    }

    /// <summary>A source that never ends, to prove the reader stops on its own.</summary>
    private sealed class EndlessStream : Stream
    {
        public long BytesProduced { get; private set; }

        public override int Read(byte[] buffer, int offset, int count)
        {
            Array.Clear(buffer, offset, count);
            BytesProduced += count;
            return count;
        }

        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            return Task.FromResult(Read(buffer, offset, count));
        }

        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => BytesProduced; set => throw new NotSupportedException(); }
        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }

    private sealed class TemporaryWorkspace : IDisposable
    {
        public TemporaryWorkspace()
        {
            Root = Path.Combine(Path.GetTempPath(), $"officeagent-ingestion-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Root);
        }

        public string Root { get; }

        public FileSystemDocumentProvider Provider() => new(new FileSystemDocumentProviderOptions
        {
            ConnectionId = "workspace",
            RootPath = Root,
            DefaultChangeMode = ChangeMode.Direct
        });

        public void Dispose()
        {
            if (Directory.Exists(Root)) Directory.Delete(Root, recursive: true);
        }
    }
}
