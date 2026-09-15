using System.Diagnostics;
using System.IO.Compression;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.Json.Serialization;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Excel;
using OfficeAgent.PowerPoint;
using OfficeAgent.Word;

namespace OfficeAgent.Benchmarks;

/// <summary>
/// Measures the release candidate against the tracked v0.8 acceptance corpus and writes raw,
/// machine-readable results.
/// </summary>
/// <remarks>
/// <para>
/// This harness measures and records. It does not summarise, decide, or gate: the statistics,
/// the report and the threshold policy live in <c>scripts/benchmark_report.py</c>, where they
/// can be tested against fixed inputs without running a benchmark at all. Anything that both
/// produced numbers and judged them would be able to hide a bad run inside its own arithmetic.
/// </para>
/// <para>
/// Every scenario asserts its own correctness, and a failed assertion is recorded as a failed
/// iteration rather than throwing the run away. A benchmark that silently drops the iterations
/// it did not like reports the speed of the cases that happened to work.
/// </para>
/// </remarks>
internal static class Program
{
    private static readonly JsonSerializerOptions Json = new()
    {
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    public static async Task<int> Main(string[] args)
    {
        try
        {
            var corpus = Argument(args, "--corpus")
                ?? Path.Combine(RepositoryRoot(), "tests", "OfficeAgent.Tests", "Corpus", "v0.8.0");
            var output = Argument(args, "--output") ?? "benchmark-results.json";
            var repetitions = int.Parse(Argument(args, "--repetitions") ?? "15", CultureInfo.InvariantCulture);
            var warmup = int.Parse(Argument(args, "--warmup") ?? "3", CultureInfo.InvariantCulture);
            var only = Argument(args, "--scenario");

            if (repetitions < 1) throw new ArgumentException("--repetitions must be at least 1.");
            if (warmup < 0) throw new ArgumentException("--warmup cannot be negative.");
            if (!Directory.Exists(corpus)) throw new DirectoryNotFoundException($"Corpus not found: {corpus}");

            var run = await RunAsync(corpus, repetitions, warmup, only).ConfigureAwait(false);

            var path = Path.GetFullPath(output);
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            await File.WriteAllTextAsync(path, JsonSerializer.Serialize(run, Json)).ConfigureAwait(false);

            var failed = run.Scenarios.Sum(scenario => scenario.FailedIterations);
            Console.WriteLine($"benchmark-scenarios={run.Scenarios.Count} failed-iterations={failed}");
            Console.WriteLine($"raw-results={path}");
            return failed == 0 ? 0 : 1;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"ERROR: {ex.Message}");
            return 2;
        }
    }

    private static async Task<BenchmarkRun> RunAsync(string corpus, int repetitions, int warmup, string? only)
    {
        var workspace = Path.Combine(
            Path.GetTempPath(), $"officeagent-benchmark-{Guid.NewGuid():N}");
        Directory.CreateDirectory(workspace);
        try
        {
            var scenarios = new List<ScenarioResult>();
            foreach (var scenario in Scenarios())
            {
                if (only is not null && !string.Equals(only, scenario.Name, StringComparison.Ordinal)) continue;
                scenarios.Add(await MeasureAsync(scenario, corpus, workspace, repetitions, warmup)
                    .ConfigureAwait(false));
            }

            if (scenarios.Count == 0)
                throw new ArgumentException($"No scenario matched '{only}'.");

            using var process = Process.GetCurrentProcess();
            return new BenchmarkRun
            {
                SchemaVersion = "1.0",
                Environment = Describe(corpus, repetitions, warmup, process),
                Scenarios = scenarios
            };
        }
        finally
        {
            try { Directory.Delete(workspace, recursive: true); } catch (IOException) { }
        }
    }

    /// <summary>
    /// Runs one scenario: warmup iterations that are measured and discarded, then the
    /// recorded repetitions.
    /// </summary>
    /// <remarks>
    /// Warmup exists because the first call through a code path pays for JIT compilation and
    /// for the Open XML SDK's static initialisation, which is real work but not the work
    /// being measured. The warmup count is recorded so a reader can see what was excluded.
    /// </remarks>
    private static async Task<ScenarioResult> MeasureAsync(
        Scenario scenario, string corpus, string workspaceRoot, int repetitions, int warmup)
    {
        var workspace = Path.Combine(workspaceRoot, scenario.Name);
        Directory.CreateDirectory(workspace);

        var samples = new List<Iteration>();
        for (var index = 0; index < warmup + repetitions; index++)
        {
            var recorded = index >= warmup;
            var iterationRoot = Path.Combine(workspace, index.ToString(CultureInfo.InvariantCulture));
            Directory.CreateDirectory(iterationRoot);

            var client = NewClient(iterationRoot, out var provider);
            _ = provider;
            var context = new ScenarioContext(client, corpus, iterationRoot, scenario.PreservedParts);

            // Collect before timing so an earlier iteration's garbage is not billed to this
            // one, and so the allocation figure covers this iteration alone.
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            var allocatedBefore = GC.GetTotalAllocatedBytes(precise: true);

            var stopwatch = Stopwatch.StartNew();
            ScenarioOutcome outcome;
            try
            {
                outcome = await scenario.RunAsync(context).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                outcome = ScenarioOutcome.Failed($"{ex.GetType().Name}: {ex.Message}");
            }

            stopwatch.Stop();
            var allocated = GC.GetTotalAllocatedBytes(precise: true) - allocatedBefore;

            if (recorded)
                samples.Add(new Iteration
                {
                    Index = samples.Count,
                    ElapsedMilliseconds = stopwatch.Elapsed.TotalMilliseconds,
                    AllocatedBytes = allocated,
                    OutputBytes = outcome.OutputBytes,
                    WorkUnits = outcome.WorkUnits,
                    InputBytes = context.InputBytes,
                    Succeeded = outcome.Succeeded,
                    Failure = outcome.Failure
                });

            try { Directory.Delete(iterationRoot, recursive: true); } catch (IOException) { }
        }

        using var process = Process.GetCurrentProcess();
        process.Refresh();
        return new ScenarioResult
        {
            Name = scenario.Name,
            Description = scenario.Description,
            Inputs = scenario.Inputs,
            WarmupIterations = warmup,
            Iterations = samples,
            FailedIterations = samples.Count(sample => !sample.Succeeded),
            ConcurrencyLevel = scenario.ConcurrencyLevel,
            PreservedParts = scenario.PreservedParts,
            // Process-level and cumulative for the whole run to this point. It is not this
            // scenario's own peak and must not be reported as one; the per-iteration
            // allocation figure is the attributable number.
            ProcessPeakWorkingSetBytes = process.PeakWorkingSet64
        };
    }

    private static OfficeAgentClient NewClient(string root, out FileSystemDocumentProvider provider)
    {
        provider = new FileSystemDocumentProvider(new FileSystemDocumentProviderOptions
        {
            ConnectionId = "bench",
            RootPath = root,
            DefaultChangeMode = ChangeMode.Direct,
            // The default allowlist is Word only; this harness measures all three formats.
            AllowedExtensions = new[] { ".docx", ".pptx", ".xlsx" }
        });

        return new OfficeAgentClient(
            new DocumentProviderRegistry(new[] { provider }),
            new WordModule(), new PowerPointModule(), new ExcelModule());
    }

    private static IEnumerable<Scenario> Scenarios()
    {
        yield return new Scenario(
            "word-contract-tracked-edit",
            "Inspect a contract, then apply one tracked text change through a provider.",
            new[] { "complex-contract.docx" },
            async context =>
            {
                var reference = await context.StageAsync("complex-contract.docx").ConfigureAwait(false);
                var inspected = await context.Client.InspectAsync(reference).ConfigureAwait(false);
                var paragraph = inspected.Paragraphs.FirstOrDefault(p => !string.IsNullOrWhiteSpace(p.Text))
                    ?? throw new InvalidOperationException("the contract fixture has no text paragraph");

                var plan = new DocumentPlan
                {
                    Operations = new PlanOperation[]
                    {
                        new ChangeTextOp
                        {
                            Target = new TextSpanAnchor { ParaId = paragraph.ParaId, Expect = paragraph.Text },
                            With = paragraph.Text + " (reviewed)",
                            Mode = ChangeMode.Tracked
                        }
                    }
                };

                var result = await context.Client.CommitAsync(reference, plan).ConfigureAwait(false);
                if (!result.Committed || result.Document is null)
                    return ScenarioOutcome.Failed("the tracked edit did not commit");

                var bytes = await context.ReadAsync(result.Document).ConfigureAwait(false);
                if (context.CheckPreserved(bytes) is { } damage) return ScenarioOutcome.Failed(damage);
                return ScenarioOutcome.Ok(bytes.LongLength, workUnits: inspected.Paragraphs.Count);
            })
        {
            // The parts the v0.8 corpus manifest declares protected for this workflow.
            PreservedParts = new[] { "word/comments.xml", "word/footnotes.xml", "word/styles.xml" }
        };

        yield return new Scenario(
            "word-comparison",
            "Compare two revisions of a Word document and produce a tracked-change plan.",
            new[] { "comparison-original.docx", "comparison-revised.docx" },
            async context =>
            {
                var original = await context.StageAsync("comparison-original.docx").ConfigureAwait(false);
                var revised = await context.StageAsync("comparison-revised.docx").ConfigureAwait(false);

                var comparison = await context.Client
                    .CompareDocumentsAsync(original, revised, new DocumentComparisonOptions())
                    .ConfigureAwait(false);

                if (comparison.Differences.Count == 0)
                    return ScenarioOutcome.Failed("the comparison found no differences between two revisions");

                return ScenarioOutcome.Ok(0, workUnits: comparison.Differences.Count);
            });

        yield return new Scenario(
            "template-batch-repeated",
            "Populate one Word template into eight independent outputs through a preflighted batch.",
            new[] { "template.docx" },
            async context =>
            {
                var template = await context.StageAsync("template.docx").ConfigureAwait(false);
                var discovery = await context.Client.DiscoverTemplateAsync(template).ConfigureAwait(false);
                var slot = discovery.Slots.FirstOrDefault()?.Name
                    ?? throw new InvalidOperationException("the template fixture exposes no slot");

                var request = new TemplateBatchRequest
                {
                    Items = Enumerable.Range(0, 8).Select(index => new TemplateBatchItem
                    {
                        OutputName = $"quote-{index}.docx",
                        Binding = new TemplateBinding
                        {
                            Values = new Dictionary<string, string?> { [slot] = $"Customer {index}" },
                            MissingValueBehavior = MissingTemplateValueBehavior.Ignore
                        }
                    }).ToArray()
                };

                var preview = await context.Client.PreviewTemplateBatchAsync(template, request)
                    .ConfigureAwait(false);
                if (!preview.IsValid) return ScenarioOutcome.Failed("the batch preflight refused a valid batch");

                var result = await context.Client
                    .PopulateTemplateBatchAsync(template, request, preview.Token)
                    .ConfigureAwait(false);
                if (!result.Committed)
                    return ScenarioOutcome.Failed("the token-bound batch commit was refused");

                long bytes = 0;
                foreach (var item in result.Items.Where(item => item.Document is not null))
                    bytes += (await context.ReadAsync(item.Document!).ConfigureAwait(false)).LongLength;

                return ScenarioOutcome.Ok(bytes, workUnits: result.Items.Count);
            });

        yield return new Scenario(
            "word-assembly",
            "Assemble two Word documents into one through preview and commit.",
            new[] { "assembly-source-a.docx", "assembly-source-b.docx" },
            async context =>
            {
                var first = await context.StageAsync("assembly-source-a.docx").ConfigureAwait(false);
                var second = await context.StageAsync("assembly-source-b.docx").ConfigureAwait(false);

                var preview = await context.Client.PreviewMergeAsync(new DocumentMergeRequest
                {
                    Sources = new[] { first, second }
                }).ConfigureAwait(false);

                if (preview.Plan is null)
                    return ScenarioOutcome.Failed("the assembly preview produced no plan");

                var result = await context.Client
                    .CommitMergeAsync(preview.Plan, "bench", "assembled.docx")
                    .ConfigureAwait(false);
                if (result.Document is null)
                    return ScenarioOutcome.Failed("the assembly commit produced no document");

                var bytes = await context.ReadAsync(result.Document).ConfigureAwait(false);
                return ScenarioOutcome.Ok(bytes.LongLength, workUnits: 2);
            });

        yield return new Scenario(
            "powerpoint-chart-binding",
            "Discover a deck's native chart slot and rewrite the chart's data through a bound batch.",
            new[] { "deck-with-chart.pptx" },
            async context =>
            {
                var deck = await context.StageAsync("deck-with-chart.pptx").ConfigureAwait(false);
                var discovered = await context.Client.DiscoverTemplateAsync(deck).ConfigureAwait(false);
                var slot = discovered.MediaSlots.FirstOrDefault(media => media.MediaKind == "chart");
                if (slot is null)
                    return ScenarioOutcome.Failed("the deck fixture exposes no chart slot to bind");

                var request = new TemplateBatchRequest
                {
                    Items = new[]
                    {
                        new TemplateBatchItem
                        {
                            OutputName = "bound-deck.pptx",
                            Binding = new TemplateBinding
                            {
                                TypedValues = new Dictionary<string, TemplateValue>
                                {
                                    [slot.Name] = new TemplateChartValue
                                    {
                                        Kind = ChartKind.Bar,
                                        Categories = new[] { "North", "South", "East", "West" },
                                        Series = new[]
                                        {
                                            new ChartSeries
                                            {
                                                Name = "Bookings",
                                                Values = new double?[] { 41, 58, 33, 27 }
                                            }
                                        },
                                        Title = "Bookings by region"
                                    }
                                },
                                MissingValueBehavior = MissingTemplateValueBehavior.Ignore
                            }
                        }
                    }
                };

                var result = await context.Client
                    .PopulateTemplateBatchAsync(deck, request)
                    .ConfigureAwait(false);
                if (!result.Committed)
                    return ScenarioOutcome.Failed("the chart binding did not commit");

                var document = result.Items.Select(item => item.Document).FirstOrDefault(d => d is not null);
                if (document is null) return ScenarioOutcome.Failed("the chart binding produced no document");

                var bytes = await context.ReadAsync(document).ConfigureAwait(false);
                if (context.CheckPreserved(bytes) is { } damage) return ScenarioOutcome.Failed(damage);
                return ScenarioOutcome.Ok(bytes.LongLength, workUnits: 4);
            })
        {
            PreservedParts = new[]
            {
                "ppt/slideMasters/slideMaster1.xml", "ppt/notesSlides/notesSlide1.xml"
            }
        };

        yield return new Scenario(
            "powerpoint-image-heavy-deck",
            "Grow a deck to eight slides and place a background image on every one of them.",
            new[] { "deck-with-chart.pptx" },
            async context =>
            {
                var deck = await context.StageAsync("deck-with-chart.pptx").ConfigureAwait(false);

                var grow = new DocumentPlan
                {
                    Operations = Enumerable.Range(0, 6)
                        .Select(index => (PlanOperation)new InsertSlideOp
                        {
                            Position = SlidePosition.End,
                            Slide = new SlideData { Title = $"Image slide {index}" }
                        })
                        .ToArray()
                };

                var grown = await context.Client.CommitAsync(deck, grow).ConfigureAwait(false);
                if (!grown.Committed || grown.Document is null)
                    return ScenarioOutcome.Failed("the deck could not be grown");

                var inspected = await context.Client.InspectAsync(grown.Document).ConfigureAwait(false);
                var slides = inspected.Nodes
                    .Where(node => string.Equals(node.Kind, "slide", StringComparison.Ordinal))
                    .ToList();
                if (slides.Count < 8)
                    return ScenarioOutcome.Failed($"expected at least 8 slides, found {slides.Count}");

                var images = new DocumentPlan
                {
                    Operations = slides
                        .Select(slide => (PlanOperation)new BackgroundImageOp
                        {
                            Target = new NodeAnchor { Kind = "slide", Path = slide.Path },
                            Base64Bytes = Png
                        })
                        .ToArray()
                };

                var result = await context.Client.CommitAsync(grown.Document, images).ConfigureAwait(false);
                if (!result.Committed || result.Document is null)
                    return ScenarioOutcome.Failed("the image placements did not commit");

                var bytes = await context.ReadAsync(result.Document).ConfigureAwait(false);
                return ScenarioOutcome.Ok(bytes.LongLength, workUnits: slides.Count);
            });

        yield return new Scenario(
            "excel-bounded-cell-edit",
            "Edit cells in a styled workbook under a bounded cell budget, then read the output back.",
            new[] { "workbook-styles.xlsx" },
            async context =>
            {
                var workbook = await context.StageAsync("workbook-styles.xlsx").ConfigureAwait(false);
                var inspected = await context.Client
                    .InspectAsync(workbook, new InspectOptions { MaximumCells = 500 })
                    .ConfigureAwait(false);

                var sheet = inspected.Nodes
                    .FirstOrDefault(node => string.Equals(node.Kind, "worksheet", StringComparison.Ordinal));
                if (sheet is null) return ScenarioOutcome.Failed("the workbook fixture reported no sheet");

                // The sheet's durable catalogue id, taken from the inspect rather than
                // assumed: a rename does not change it, but a different fixture would.
                var sheetId = SheetIdOf(sheet.Path);
                if (sheetId is null)
                    return ScenarioOutcome.Failed($"could not read a sheet id from '{sheet.Path}'");

                var plan = new DocumentPlan
                {
                    Operations = Enumerable.Range(1, 8)
                        .Select(row => (PlanOperation)new SetCellOp
                        {
                            Target = new CellAnchor { SheetId = sheetId.Value, Address = $"H{row}" },
                            Value = (row * 11).ToString(CultureInfo.InvariantCulture)
                        })
                        .ToArray()
                };

                var result = await context.Client.CommitAsync(workbook, plan).ConfigureAwait(false);
                if (!result.Committed || result.Document is null)
                    return ScenarioOutcome.Failed("the workbook edit did not commit");

                var bytes = await context.ReadAsync(result.Document).ConfigureAwait(false);
                if (context.CheckPreserved(bytes) is { } damage) return ScenarioOutcome.Failed(damage);
                return ScenarioOutcome.Ok(bytes.LongLength, workUnits: 8);
            })
        {
            // xl/styles.xml only. The corpus manifest also protects xl/tables/table1.xml, and
            // it does survive semantically, but the SDK re-serialises the tables part it read
            // and moves its namespace declaration, so the bytes differ. That difference is
            // pinned by ExcelPreservationTests rather than hidden by widening this list.
            PreservedParts = new[] { "xl/styles.xml" }
        };

        yield return new Scenario(
            "word-concurrent-edits",
            "Apply a tracked edit to each of four documents at once through one client.",
            new[] { "complex-contract.docx" },
            async context =>
            {
                // V09-07 defined the concurrency and isolation semantics; nothing measured
                // them under concurrency until now. Separate documents, so this measures
                // parallel throughput rather than the per-item gate's contention.
                var references = new List<DocumentReference>();
                for (var index = 0; index < ConcurrentDocuments; index++)
                    references.Add(await context
                        .StageAsync("complex-contract.docx", $"contract-{index}.docx")
                        .ConfigureAwait(false));

                var edits = references.Select(async reference =>
                {
                    var inspected = await context.Client.InspectAsync(reference).ConfigureAwait(false);
                    var paragraph = inspected.Paragraphs.First(p => !string.IsNullOrWhiteSpace(p.Text));
                    var plan = new DocumentPlan
                    {
                        Operations = new PlanOperation[]
                        {
                            new ChangeTextOp
                            {
                                Target = new TextSpanAnchor
                                {
                                    ParaId = paragraph.ParaId,
                                    Expect = paragraph.Text
                                },
                                With = paragraph.Text + " (reviewed)",
                                Mode = ChangeMode.Tracked
                            }
                        }
                    };
                    return await context.Client.CommitAsync(reference, plan).ConfigureAwait(false);
                });

                var results = await Task.WhenAll(edits).ConfigureAwait(false);
                if (results.Any(result => !result.Committed))
                    return ScenarioOutcome.Failed(
                        $"{results.Count(r => !r.Committed)} of {results.Length} concurrent edits did not commit");

                long bytes = 0;
                foreach (var result in results.Where(r => r.Document is not null))
                    bytes += (await context.ReadAsync(result.Document!).ConfigureAwait(false)).LongLength;

                return ScenarioOutcome.Ok(bytes, workUnits: results.Length);
            })
        {
            ConcurrencyLevel = ConcurrentDocuments
        };
    }

    /// <summary>Reads the numeric sheet id out of an inspect node path such as "sheet#7".</summary>
    private static uint? SheetIdOf(string path)
    {
        var digits = new string(path.Where(char.IsDigit).ToArray());
        return uint.TryParse(digits, NumberStyles.Integer, CultureInfo.InvariantCulture, out var id)
            ? id
            : null;
    }

    /// <summary>Documents edited simultaneously by the concurrency scenario.</summary>
    private const int ConcurrentDocuments = 4;

    /// <summary>A one-pixel PNG, so image placement is measured without a large asset.</summary>
    private const string Png =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

    private static RunEnvironment Describe(string corpus, int repetitions, int warmup, Process process) => new()
    {
        Commit = Command("git", "rev-parse HEAD"),
        CommitIsClean = string.IsNullOrWhiteSpace(Command("git", "status --porcelain")),
        Corpus = Path.GetFullPath(corpus),
        RepetitionsPerScenario = repetitions,
        WarmupIterations = warmup,
        Runtime = RuntimeInformation.FrameworkDescription,
        OperatingSystem = RuntimeInformation.OSDescription,
        Architecture = RuntimeInformation.ProcessArchitecture.ToString(),
        ProcessorCount = System.Environment.ProcessorCount,
        ServerGarbageCollection = System.Runtime.GCSettings.IsServerGC,
        Configuration =
#if DEBUG
            "Debug",
#else
            "Release",
#endif
        MeasurementMethod =
            "Stopwatch wall clock around one scenario iteration, after a forced blocking collection. " +
            "Allocation is GC.GetTotalAllocatedBytes delta for the iteration. Peak working set is " +
            "process-level and cumulative, not per scenario.",
        StartedUtc = DateTimeOffset.UtcNow
    };

    private static string Command(string file, string arguments)
    {
        try
        {
            using var process = Process.Start(new ProcessStartInfo(file, arguments)
            {
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false
            });
            if (process is null) return string.Empty;
            var text = process.StandardOutput.ReadToEnd().Trim();
            process.WaitForExit(10_000);
            return text;
        }
        catch (Exception)
        {
            return string.Empty;
        }
    }

    private static string? Argument(string[] args, string name)
    {
        var index = Array.IndexOf(args, name);
        return index >= 0 && index + 1 < args.Length ? args[index + 1] : null;
    }

    private static string RepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "OfficeAgent.NET.sln")))
            directory = directory.Parent;
        return directory?.FullName
            ?? throw new DirectoryNotFoundException("Could not locate the repository root.");
    }

}

internal sealed record Scenario(
    string Name,
    string Description,
    IReadOnlyList<string> Inputs,
    Func<ScenarioContext, Task<ScenarioOutcome>> RunAsync)
{
    /// <summary>How many operations run at once. One means sequential.</summary>
    public int ConcurrencyLevel { get; init; } = 1;

    /// <summary>
    /// Parts this scenario requires to survive its edit byte for byte.
    /// </summary>
    /// <remarks>
    /// Correctness and preservation are different claims. A scenario can produce exactly the
    /// edit it promised and still have destroyed a styles part or an image on the way, which
    /// is the failure the corpus manifest's protected parts exist to catch.
    /// </remarks>
    public IReadOnlyList<string> PreservedParts { get; init; } = Array.Empty<string>();
}

internal sealed class ScenarioContext
{
    private readonly string _corpus;
    private readonly string _root;
    private readonly IReadOnlyList<string> _preserved;
    private readonly Dictionary<string, byte[]> _stagedParts = new(StringComparer.Ordinal);

    public ScenarioContext(
        OfficeAgentClient client, string corpus, string root, IReadOnlyList<string> preserved)
    {
        Client = client;
        _corpus = corpus;
        _root = root;
        _preserved = preserved;
    }

    public OfficeAgentClient Client { get; }

    /// <summary>Bytes of corpus input this iteration staged.</summary>
    public long InputBytes { get; private set; }

    /// <summary>Copies a corpus fixture into this iteration's workspace and registers it.</summary>
    public Task<DocumentReference> StageAsync(string fixture) => StageAsync(fixture, fixture);

    /// <summary>Stages a corpus fixture under a chosen name, so one fixture can be staged twice.</summary>
    public async Task<DocumentReference> StageAsync(string fixture, string name)
    {
        var source = Path.Combine(_corpus, fixture);
        if (!File.Exists(source)) throw new FileNotFoundException($"Corpus fixture not found: {source}");
        var destination = Path.Combine(_root, name);
        File.Copy(source, destination, overwrite: true);
        var bytes = await File.ReadAllBytesAsync(destination).ConfigureAwait(false);
        InputBytes += bytes.LongLength;

        // Snapshot the parts this scenario promises to preserve, so the promise can be
        // checked against the output rather than assumed from a successful commit.
        foreach (var part in _preserved)
            if (TryReadPart(bytes, part, out var content))
                _stagedParts[part] = content;

        return await Client.RegisterAsync("bench", destination).ConfigureAwait(false);
    }

    public async Task<byte[]> ReadAsync(DocumentReference reference)
    {
        using var content = await Client.OpenReadAsync(reference).ConfigureAwait(false);
        using var copy = new MemoryStream();
        await content.Stream.CopyToAsync(copy).ConfigureAwait(false);
        return copy.ToArray();
    }

    /// <summary>
    /// Requires every declared part to be byte-identical in <paramref name="output"/>.
    /// </summary>
    /// <remarks>
    /// A scenario that produced exactly the edit it promised can still have rewritten a
    /// styles part, dropped an image, or re-serialised a part it never needed to touch.
    /// Correctness does not catch that; only comparing the bytes does.
    /// </remarks>
    public string? CheckPreserved(byte[] output)
    {
        var damaged = new List<string>();
        foreach (var pair in _stagedParts)
        {
            if (!TryReadPart(output, pair.Key, out var after))
            {
                damaged.Add($"{pair.Key} (missing)");
                continue;
            }

            if (!after.AsSpan().SequenceEqual(pair.Value)) damaged.Add(pair.Key);
        }

        return damaged.Count == 0
            ? null
            : $"parts that had to survive the edit did not: {string.Join(", ", damaged)}";
    }

    private static bool TryReadPart(byte[] package, string part, out byte[] content)
    {
        content = Array.Empty<byte>();
        try
        {
            using var archive = new ZipArchive(
                new MemoryStream(package, writable: false), ZipArchiveMode.Read);
            var entry = archive.GetEntry(part);
            if (entry is null) return false;
            using var stream = entry.Open();
            using var copy = new MemoryStream();
            stream.CopyTo(copy);
            content = copy.ToArray();
            return true;
        }
        catch (InvalidDataException)
        {
            return false;
        }
    }
}

internal sealed class ScenarioOutcome
{
    private ScenarioOutcome(bool succeeded, long outputBytes, int workUnits, string? failure)
    {
        Succeeded = succeeded;
        OutputBytes = outputBytes;
        WorkUnits = workUnits;
        Failure = failure;
    }

    public bool Succeeded { get; }
    public long OutputBytes { get; }
    public int WorkUnits { get; }
    public string? Failure { get; }

    public static ScenarioOutcome Ok(long outputBytes, int workUnits) =>
        new(true, outputBytes, workUnits, null);

    public static ScenarioOutcome Failed(string failure) => new(false, 0, 0, failure);
}

internal sealed class BenchmarkRun
{
    public string SchemaVersion { get; init; } = "1.0";
    public RunEnvironment Environment { get; init; } = new();
    public IReadOnlyList<ScenarioResult> Scenarios { get; init; } = Array.Empty<ScenarioResult>();
}

internal sealed class RunEnvironment
{
    public string Commit { get; init; } = string.Empty;
    public bool CommitIsClean { get; init; }
    public string Corpus { get; init; } = string.Empty;
    public int RepetitionsPerScenario { get; init; }
    public int WarmupIterations { get; init; }
    public string Runtime { get; init; } = string.Empty;
    public string OperatingSystem { get; init; } = string.Empty;
    public string Architecture { get; init; } = string.Empty;
    public int ProcessorCount { get; init; }
    public bool ServerGarbageCollection { get; init; }
    public string Configuration { get; init; } = string.Empty;
    public string MeasurementMethod { get; init; } = string.Empty;
    public DateTimeOffset StartedUtc { get; init; }
}

internal sealed class ScenarioResult
{
    public string Name { get; init; } = string.Empty;
    public string Description { get; init; } = string.Empty;
    public IReadOnlyList<string> Inputs { get; init; } = Array.Empty<string>();
    public int WarmupIterations { get; init; }
    public IReadOnlyList<Iteration> Iterations { get; init; } = Array.Empty<Iteration>();
    public int FailedIterations { get; init; }
    public long ProcessPeakWorkingSetBytes { get; init; }

    /// <summary>
    /// How many operations this scenario ran at once. One means sequential.
    /// </summary>
    /// <remarks>
    /// Recorded on every scenario rather than only the concurrent one, so a reader never has
    /// to assume which it was looking at.
    /// </remarks>
    public int ConcurrencyLevel { get; init; } = 1;

    /// <summary>Package parts this scenario asserted were left byte-identical.</summary>
    public IReadOnlyList<string> PreservedParts { get; init; } = Array.Empty<string>();
}

internal sealed class Iteration
{
    public int Index { get; init; }
    public double ElapsedMilliseconds { get; init; }
    public long AllocatedBytes { get; init; }
    public long OutputBytes { get; init; }
    public int WorkUnits { get; init; }

    /// <summary>Bytes of input this iteration read, so throughput has a denominator.</summary>
    public long InputBytes { get; init; }

    public bool Succeeded { get; init; }
    public string? Failure { get; init; }
}
