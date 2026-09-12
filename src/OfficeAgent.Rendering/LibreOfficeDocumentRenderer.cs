using System.Diagnostics;
using OfficeAgent.Abstractions;

namespace OfficeAgent.Rendering;

/// <summary>Configuration for the isolated LibreOffice and PDF-to-PNG processes.</summary>
public sealed class LibreOfficeRendererOptions
{
    /// <summary>Gets or sets the LibreOffice command. Defaults to <c>soffice</c>.</summary>
    public string LibreOfficeExecutable { get; set; } = "soffice";

    /// <summary>Gets arguments inserted before LibreOffice's standard arguments.</summary>
    public IList<string> LibreOfficePrefixArguments { get; set; } = new List<string>();

    /// <summary>Gets or sets the Poppler PDF rasterizer command. Defaults to <c>pdftoppm</c>.</summary>
    public string PdfToPpmExecutable { get; set; } = "pdftoppm";

    /// <summary>Gets arguments inserted before the rasterizer's standard arguments.</summary>
    public IList<string> PdfToPpmPrefixArguments { get; set; } = new List<string>();
}

/// <summary>
/// Renders Word, PowerPoint, and Excel packages to PNG pages using two child processes:
/// LibreOffice creates a PDF and Poppler's <c>pdftoppm</c> rasterizes it. No renderer code
/// or document content runs inside OfficeAgent.Core.
/// </summary>
public sealed class LibreOfficeDocumentRenderer : IDocumentRenderer
{
    private static readonly HashSet<string> Extensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".docx", ".pptx", ".xlsx"
    };

    private readonly LibreOfficeRendererOptions _renderer;

    /// <summary>Creates a renderer with the supplied executable locations.</summary>
    public LibreOfficeDocumentRenderer(LibreOfficeRendererOptions? options = null) =>
        _renderer = options ?? new LibreOfficeRendererOptions();

    /// <inheritdoc />
    public async Task<RenderResult> RenderAsync(
        Stream document,
        RenderOptions options,
        CancellationToken cancellationToken = default)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (options is null) throw new ArgumentNullException(nameof(options));
        var invalid = Validate(options);
        if (invalid is not null) return Failure("invalid-render-options", invalid);

        var root = Path.Combine(Path.GetTempPath(), "officeagent-render-" + Guid.NewGuid().ToString("N"));
        var inputDirectory = Path.Combine(root, "input");
        var outputDirectory = Path.Combine(root, "output");
        var profileDirectory = Path.Combine(root, "profile");
        Directory.CreateDirectory(inputDirectory);
        Directory.CreateDirectory(outputDirectory);
        Directory.CreateDirectory(profileDirectory);

        try
        {
            var fileName = Path.GetFileName(options.FileName);
            var inputPath = Path.Combine(inputDirectory, fileName);
            var copied = await CopyBoundedAsync(
                document, inputPath, options.MaximumInputBytes, cancellationToken).ConfigureAwait(false);
            if (!copied) return Failure("input-limit-exceeded", "The document exceeded the configured input-size limit.");

            var clock = Stopwatch.StartNew();
            var convertArguments = _renderer.LibreOfficePrefixArguments.Concat(new[]
            {
                "--headless", "--nologo", "--nodefault", "--nofirststartwizard",
                "-env:UserInstallation=" + new Uri(profileDirectory + Path.DirectorySeparatorChar).AbsoluteUri,
                "--convert-to", "pdf", "--outdir", outputDirectory, inputPath
            });
            var converted = await RunBoundedAsync(
                _renderer.LibreOfficeExecutable, convertArguments, root, outputDirectory,
                options, clock, countPages: false, cancellationToken).ConfigureAwait(false);
            if (converted is not null) return converted;

            var pdfPath = Directory.EnumerateFiles(outputDirectory, "*.pdf", SearchOption.TopDirectoryOnly)
                .SingleOrDefault();
            if (pdfPath is null)
                return Failure("renderer-output-missing", "LibreOffice completed without producing a PDF.");

            var pagePrefix = Path.Combine(outputDirectory, "page");
            var rasterArguments = _renderer.PdfToPpmPrefixArguments.Concat(new[]
            {
                "-png", "-r", options.Dpi.ToString(System.Globalization.CultureInfo.InvariantCulture),
                pdfPath, pagePrefix
            });
            var rasterized = await RunBoundedAsync(
                _renderer.PdfToPpmExecutable, rasterArguments, root, outputDirectory,
                options, clock, countPages: true, cancellationToken).ConfigureAwait(false);
            if (rasterized is not null) return rasterized;

            var pageFiles = PageFiles(outputDirectory).ToArray();
            if (pageFiles.Length == 0)
                return Failure("renderer-output-missing", "The rasterizer completed without producing page images.");
            if (pageFiles.Length > options.MaximumPages)
                return Failure("page-limit-exceeded", "Rendering exceeded the configured page-count limit.");

            var pages = new List<RenderedPage>(pageFiles.Length);
            long outputBytes = 0;
            for (var index = 0; index < pageFiles.Length; index++)
            {
                var bytes = await File.ReadAllBytesAsync(pageFiles[index], cancellationToken).ConfigureAwait(false);
                outputBytes += bytes.Length;
                if (outputBytes > options.MaximumOutputBytes)
                    return Failure("output-limit-exceeded", "Rendering exceeded the configured output-size limit.");
                pages.Add(new RenderedPage { PageNumber = index + 1, Content = bytes });
            }

            return new RenderResult { Succeeded = true, Pages = pages };
        }
        finally
        {
            try { Directory.Delete(root, recursive: true); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }

    private static string? Validate(RenderOptions options)
    {
        var fileName = Path.GetFileName(options.FileName);
        if (string.IsNullOrWhiteSpace(fileName) || !Extensions.Contains(Path.GetExtension(fileName)))
            return "FileName must end in .docx, .pptx, or .xlsx.";
        if (options.Dpi is < 36 or > 600) return "Dpi must be between 36 and 600.";
        if (options.Timeout <= TimeSpan.Zero) return "Timeout must be positive.";
        if (options.MaximumWorkingSetBytes <= 0 || options.MaximumPages <= 0 ||
            options.MaximumOutputBytes <= 0 || options.MaximumInputBytes <= 0)
            return "All renderer limits must be positive.";
        return null;
    }

    private static async Task<bool> CopyBoundedAsync(
        Stream source,
        string destination,
        long maximumBytes,
        CancellationToken cancellationToken)
    {
        await using var output = new FileStream(
            destination, FileMode.CreateNew, FileAccess.Write, FileShare.None, 81920, useAsync: true);
        var buffer = new byte[81920];
        long total = 0;
        while (true)
        {
            var read = await source.ReadAsync(buffer.AsMemory(), cancellationToken).ConfigureAwait(false);
            if (read == 0) return true;
            total += read;
            if (total > maximumBytes) return false;
            await output.WriteAsync(buffer.AsMemory(0, read), cancellationToken).ConfigureAwait(false);
        }
    }

    private static async Task<RenderResult?> RunBoundedAsync(
        string executable,
        IEnumerable<string> arguments,
        string workingDirectory,
        string outputDirectory,
        RenderOptions options,
        Stopwatch clock,
        bool countPages,
        CancellationToken cancellationToken)
    {
        using var process = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = executable,
                WorkingDirectory = workingDirectory,
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            }
        };
        foreach (var argument in arguments) process.StartInfo.ArgumentList.Add(argument);

        try
        {
            if (!process.Start())
                return Failure("renderer-unavailable", "The renderer process could not be started.");
        }
        catch (Exception ex) when (ex is System.ComponentModel.Win32Exception or FileNotFoundException)
        {
            return Failure("renderer-unavailable", "A configured renderer executable was not found.");
        }

        using var drainDeadline = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        var remaining = options.Timeout - clock.Elapsed;
        if (remaining <= TimeSpan.Zero) drainDeadline.Cancel(); else drainDeadline.CancelAfter(remaining);
        var stdout = DrainAsync(process.StandardOutput, drainDeadline.Token);
        var stderr = DrainAsync(process.StandardError, drainDeadline.Token);
        try
        {
            while (!process.HasExited)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (clock.Elapsed > options.Timeout)
                    return await StopAndDrainAsync(
                        process, stdout, stderr, drainDeadline, "render-timeout",
                        "Rendering exceeded the configured time limit.").ConfigureAwait(false);
                process.Refresh();
                if (process.HasExited) break;
                long workingSet;
                try { workingSet = process.WorkingSet64; }
                catch (InvalidOperationException) { break; }
                if (workingSet > options.MaximumWorkingSetBytes)
                    return await StopAndDrainAsync(
                        process, stdout, stderr, drainDeadline, "memory-limit-exceeded",
                        "A renderer process exceeded the configured memory limit.").ConfigureAwait(false);
                if (OutputBytes(outputDirectory) > options.MaximumOutputBytes)
                    return await StopAndDrainAsync(
                        process, stdout, stderr, drainDeadline, "output-limit-exceeded",
                        "Rendering exceeded the configured output-size limit.").ConfigureAwait(false);
                if (countPages && PageFiles(outputDirectory).Take(options.MaximumPages + 1).Count() > options.MaximumPages)
                    return await StopAndDrainAsync(
                        process, stdout, stderr, drainDeadline, "page-limit-exceeded",
                        "Rendering exceeded the configured page-count limit.").ConfigureAwait(false);
                await Task.Delay(50, cancellationToken).ConfigureAwait(false);
            }
            await process.WaitForExitAsync(cancellationToken).ConfigureAwait(false);
            try { await Task.WhenAll(stdout, stderr).ConfigureAwait(false); }
            catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
            {
                return Failure("render-timeout", "Rendering exceeded the configured time limit.");
            }
        }
        catch (OperationCanceledException)
        {
            Kill(process);
            drainDeadline.Cancel();
            await ObserveDrainAsync(stdout, stderr).ConfigureAwait(false);
            throw;
        }

        if (process.ExitCode != 0)
            return Failure("renderer-failed", "A renderer process exited unsuccessfully.");
        if (clock.Elapsed > options.Timeout)
            return Failure("render-timeout", "Rendering exceeded the configured time limit.");
        if (OutputBytes(outputDirectory) > options.MaximumOutputBytes)
            return Failure("output-limit-exceeded", "Rendering exceeded the configured output-size limit.");
        if (countPages && PageFiles(outputDirectory).Take(options.MaximumPages + 1).Count() > options.MaximumPages)
            return Failure("page-limit-exceeded", "Rendering exceeded the configured page-count limit.");
        return null;
    }

    private static async Task<RenderResult> StopAndDrainAsync(
        Process process,
        Task stdout,
        Task stderr,
        CancellationTokenSource drainDeadline,
        string code,
        string message)
    {
        Kill(process);
        drainDeadline.Cancel();
        try { await process.WaitForExitAsync().ConfigureAwait(false); }
        catch (InvalidOperationException) { }
        await ObserveDrainAsync(stdout, stderr).ConfigureAwait(false);
        return Failure(code, message);
    }

    private static async Task DrainAsync(StreamReader reader, CancellationToken cancellationToken)
    {
        var buffer = new char[4096];
        while (await reader.ReadAsync(buffer.AsMemory(), cancellationToken).ConfigureAwait(false) > 0)
        {
            // Drain without retaining process output; it can contain sensitive paths.
        }
    }

    private static async Task ObserveDrainAsync(Task stdout, Task stderr)
    {
        try { await Task.WhenAll(stdout, stderr).ConfigureAwait(false); }
        catch (OperationCanceledException) { }
        catch (IOException) { }
        catch (ObjectDisposedException) { }
    }

    private static void Kill(Process process)
    {
        try { if (!process.HasExited) process.Kill(entireProcessTree: true); }
        catch (InvalidOperationException) { }
        catch (System.ComponentModel.Win32Exception) { }
    }

    private static long OutputBytes(string directory) => Directory.EnumerateFiles(directory)
        .Sum(path => new FileInfo(path).Length);

    private static IEnumerable<string> PageFiles(string directory) =>
        Directory.EnumerateFiles(directory, "page-*.png", SearchOption.TopDirectoryOnly)
            .OrderBy(path => PageNumber(path));

    private static int PageNumber(string path)
    {
        var name = Path.GetFileNameWithoutExtension(path);
        return int.TryParse(name.Substring(name.LastIndexOf('-') + 1), out var number)
            ? number
            : int.MaxValue;
    }

    private static RenderResult Failure(string code, string message) => new()
    {
        Succeeded = false,
        FailureCode = code,
        Message = message
    };
}
