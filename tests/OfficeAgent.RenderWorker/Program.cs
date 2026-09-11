namespace OfficeAgent.RenderWorker;

/// <summary>Marker used to locate this deterministic renderer test process.</summary>
public sealed class RenderWorkerMarker;

internal static class Program
{
    public static async Task<int> Main(string[] args)
    {
        var holdPipe = Value(args, "--hold-pipe-ms=");
        if (holdPipe > 0)
        {
            await Task.Delay(holdPipe).ConfigureAwait(false);
            return 0;
        }
        if (args.Contains("--spawn-child-holding-pipe", StringComparer.Ordinal))
        {
            var child = new System.Diagnostics.ProcessStartInfo("dotnet") { UseShellExecute = false, CreateNoWindow = true };
            child.ArgumentList.Add(typeof(RenderWorkerMarker).Assembly.Location);
            child.ArgumentList.Add("--hold-pipe-ms=5000");
            System.Diagnostics.Process.Start(child);
        }

        var delay = Value(args, "--delay-ms=");
        if (delay > 0) await Task.Delay(delay).ConfigureAwait(false);

        var allocation = Value(args, "--allocate-mb=");
        byte[]? memory = null;
        if (allocation > 0)
        {
            memory = new byte[allocation * 1024 * 1024];
            for (var index = 0; index < memory.Length; index += 4096) memory[index] = 1;
            await Task.Delay(500).ConfigureAwait(false);
        }

        var convert = Array.IndexOf(args, "--convert-to");
        if (convert >= 0)
        {
            var output = args[Array.IndexOf(args, "--outdir") + 1];
            var input = args[^1];
            await File.WriteAllBytesAsync(
                Path.Combine(output, Path.GetFileNameWithoutExtension(input) + ".pdf"),
                new byte[] { 0x25, 0x50, 0x44, 0x46 }).ConfigureAwait(false);
            GC.KeepAlive(memory);
            return 0;
        }

        var pages = Math.Max(1, Value(args, "--pages="));
        var prefix = args[^1];
        for (var page = 1; page <= pages; page++)
        {
            await File.WriteAllBytesAsync(
                $"{prefix}-{page}.png",
                Enumerable.Repeat((byte)page, 64).ToArray()).ConfigureAwait(false);
        }
        GC.KeepAlive(memory);
        return 0;
    }

    private static int Value(IEnumerable<string> args, string prefix) =>
        args.Select(argument => argument.StartsWith(prefix, StringComparison.Ordinal)
                ? int.TryParse(argument[prefix.Length..], out var value) ? value : 0
                : 0)
            .FirstOrDefault(value => value > 0);
}
