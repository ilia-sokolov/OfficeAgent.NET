using System.Diagnostics;
using System.IO.Compression;
using System.Runtime.InteropServices;
using SharpFuzz;

namespace OfficeAgent.Fuzz;

/// <summary>
/// Entry point. Started by libfuzzer-dotnet with no arguments, it runs the ingestion target
/// under libFuzzer. <c>--repro &lt;file&gt;</c> runs one saved input and says what it produced;
/// <c>--check &lt;folder&gt;</c> runs every file in a folder once, which is how the seed corpus is
/// checked before a campaign; <c>--minimize &lt;file&gt; &lt;output&gt;</c> shrinks a finding to the
/// fewest package entries that still reproduce it.
/// </summary>
internal static class Program
{
    private static readonly TimeSpan DefaultTimeout = TimeSpan.FromSeconds(10);

    public static int Main(string[] args)
    {
        if (args.Length == 2 && args[0] == "--repro") return Repro(args[1]);
        if (args.Length == 2 && args[0] == "--check") return Check(args[1]);
        if (args.Length == 3 && args[0] == "--minimize") return Minimize(args[1], args[2]);

        if (args.Length != 0)
        {
            Console.Error.WriteLine("usage: OfficeAgent.Fuzz                   (under libfuzzer-dotnet)");
            Console.Error.WriteLine("       OfficeAgent.Fuzz --repro <file>    (run one input)");
            Console.Error.WriteLine("       OfficeAgent.Fuzz --check <folder>  (run every file in a folder)");
            Console.Error.WriteLine("       OfficeAgent.Fuzz --minimize <file> <output>  (drop entries while it still fails)");
            return 2;
        }

        Fuzzer.LibFuzzer.Run(data => IngestionTarget.Run(data));
        return 0;
    }

    private static int Repro(string path)
    {
        EnableTraceOutsideLibFuzzer();
        var (verdict, finding) = RunOne(File.ReadAllBytes(path));
        Console.WriteLine($"{Path.GetFileName(path)}: {verdict}");
        if (finding is not null) Console.WriteLine(finding);
        return finding is null ? 0 : 1;
    }

    private static int Check(string folder)
    {
        EnableTraceOutsideLibFuzzer();
        var files = Directory.GetFiles(folder).OrderBy(f => f, StringComparer.Ordinal).ToArray();
        var counts = new SortedDictionary<string, int>(StringComparer.Ordinal);
        int findings = 0;
        long slowest = 0, mostAllocated = 0;

        foreach (var file in files)
        {
            var watch = Stopwatch.StartNew();
            var (verdict, finding) = RunOne(File.ReadAllBytes(file));
            slowest = Math.Max(slowest, watch.ElapsedMilliseconds);
            mostAllocated = Math.Max(mostAllocated, IngestionTarget.LastAllocatedBytes);

            var key = finding is null ? verdict : "finding";
            counts[key] = counts.GetValueOrDefault(key) + 1;
            if (finding is not null)
            {
                findings++;
                Console.WriteLine($"FINDING {Path.GetFileName(file)}: {verdict}");
                Console.WriteLine(finding);
            }
        }

        Console.WriteLine($"{files.Length} inputs: " + string.Join(", ", counts.Select(c => $"{c.Key} {c.Value}")));
        Console.WriteLine($"slowest {slowest} ms, most allocated {mostAllocated / (1024 * 1024)} MiB " +
                          $"(bound {IngestionTarget.MaximumAllocatedBytes / (1024 * 1024)} MiB)");
        return findings == 0 ? 0 : 1;
    }

    /// <summary>
    /// Shrinks a finding by package entry. libFuzzer's <c>-minimize_crash</c> removes bytes, and
    /// almost any byte removed from a ZIP breaks the archive before the finding is reached, so it
    /// rarely gets anywhere. This re-writes the package with every entry stored, then drops entries
    /// one at a time while the same exception still escapes from the same OfficeAgent frame.
    /// </summary>
    private static int Minimize(string path, string output)
    {
        EnableTraceOutsideLibFuzzer();
        var original = File.ReadAllBytes(path);
        var signature = Signature(RunOne(original).Finding);
        if (signature is null)
        {
            Console.WriteLine($"{Path.GetFileName(path)} is not a finding; nothing to minimise.");
            return 2;
        }

        var entries = ReadEntries(original);
        var best = original;
        if (entries is not null && Signature(RunOne(Pack(entries)).Finding) == signature)
        {
            for (bool shrunk = true; shrunk;)
            {
                shrunk = false;
                for (int i = entries.Count - 1; i >= 0; i--)
                {
                    var candidate = entries.Where((_, index) => index != i).ToList();
                    if (Signature(RunOne(Pack(candidate)).Finding) != signature) continue;
                    entries = candidate;
                    shrunk = true;
                }
            }
            best = Pack(entries);
        }
        else
        {
            entries = null;
            Console.WriteLine("The finding depends on the archive's own bytes; kept as it is.");
        }

        File.WriteAllBytes(output, best);
        Console.WriteLine(signature);
        Console.WriteLine($"{original.Length} -> {best.Length} bytes" +
                          (entries is null ? string.Empty : $"; entries: {string.Join(", ", entries.Select(e => e.Name))}"));
        return 0;
    }

    /// <summary>The exception type and the first OfficeAgent frame, without line numbers.</summary>
    private static string? Signature(Exception? finding)
    {
        if (finding is null) return null;
        var frame = (finding.StackTrace ?? string.Empty)
            .Split('\n')
            .Select(line => line.Trim())
            .FirstOrDefault(line => line.StartsWith("at OfficeAgent.", StringComparison.Ordinal)
                                    && !line.StartsWith("at OfficeAgent.Fuzz.", StringComparison.Ordinal));
        var cut = frame?.IndexOf(" in ", StringComparison.Ordinal) ?? -1;
        if (cut > 0) frame = frame!.Substring(0, cut);
        return $"{finding.GetType().FullName} {frame}";
    }

    private static List<(string Name, byte[] Data)>? ReadEntries(byte[] package)
    {
        try
        {
            using var zip = new ZipArchive(new MemoryStream(package, writable: false), ZipArchiveMode.Read);
            var entries = new List<(string Name, byte[] Data)>();
            foreach (var entry in zip.Entries)
            {
                using var data = new MemoryStream();
                using (var stream = entry.Open()) stream.CopyTo(data);
                entries.Add((entry.FullName, data.ToArray()));
            }
            return entries;
        }
        catch (Exception ex) when (ex is InvalidDataException or IOException or NotSupportedException)
        {
            return null;
        }
    }

    private static byte[] Pack(IEnumerable<(string Name, byte[] Data)> entries)
    {
        using var buffer = new MemoryStream();
        using (var zip = new ZipArchive(buffer, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var (name, data) in entries)
            {
                using var stream = zip.CreateEntry(name, CompressionLevel.NoCompression).Open();
                stream.Write(data, 0, data.Length);
            }
        }
        return buffer.ToArray();
    }

    /// <summary>Runs one input with a time limit, returning its outcome or the finding.</summary>
    private static (string Verdict, Exception? Finding) RunOne(byte[] input)
    {
        var timeout = int.TryParse(Environment.GetEnvironmentVariable("OFFICEAGENT_FUZZ_TIMEOUT_SECONDS"), out var s) && s > 0
            ? TimeSpan.FromSeconds(s)
            : DefaultTimeout;

        var work = Task.Run(() => IngestionTarget.Run(input));
        try
        {
            if (!work.Wait(timeout))
                return ("hang", new TimeoutException($"No outcome after {timeout.TotalSeconds:0} s."));
            return (work.Result.ToString(), null);
        }
        catch (AggregateException ex) when (ex.InnerExceptions.Count == 1)
        {
            var inner = ex.InnerException!;
            return (inner.GetType().FullName ?? inner.GetType().Name, inner);
        }
    }

    /// <summary>
    /// Instrumented assemblies write coverage through <c>SharpFuzz.Common.Trace.SharedMem</c>,
    /// which libfuzzer-dotnet supplies. Outside it, point that at a scratch buffer so an
    /// instrumented build can still run a single input.
    /// </summary>
    private static unsafe void EnableTraceOutsideLibFuzzer()
    {
        if (SharpFuzz.Common.Trace.SharedMem == null)
            SharpFuzz.Common.Trace.SharedMem = (byte*)Marshal.AllocHGlobal(1 << 16);
    }
}
