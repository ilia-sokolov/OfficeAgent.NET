using System.Diagnostics;
using System.Text.Json;
using OfficeAgent.Abstractions;

namespace OfficeAgent.Deploy.Renderer;

/// <summary>Host-enforced limits for one sandboxed render container.</summary>
public sealed class SandboxLimits
{
    /// <summary>The Docker command. Defaults to <c>docker</c> on <c>PATH</c>.</summary>
    public string DockerExecutable { get; init; } = "docker";

    /// <summary>The reference image, ideally pinned by digest.</summary>
    public string Image { get; init; } = "officeagent-renderer:reference";

    /// <summary>Hard memory limit for the container, swap included.</summary>
    public long MemoryBytes { get; init; } = 1024L * 1024 * 1024;

    /// <summary>CPU quota, in cores.</summary>
    public double Cpus { get; init; } = 1.0;

    /// <summary>Maximum processes and threads in the container.</summary>
    public int Pids { get; init; } = 256;

    /// <summary>Size of the only writable filesystem, the in-memory <c>/tmp</c>.</summary>
    public long ScratchBytes { get; init; } = 256L * 1024 * 1024;

    /// <summary>Allowance for container start-up on top of the render timeout.</summary>
    public TimeSpan StartupAllowance { get; init; } = TimeSpan.FromSeconds(20);

    /// <summary>Extra environment for the worker. The deployment recipe sets none.</summary>
    public IReadOnlyDictionary<string, string> Environment { get; init; } = new Dictionary<string, string>();
}

/// <summary>
/// Renders by running the reference worker in a fresh, locked-down container per document.
/// </summary>
/// <remarks>
/// The container has no network, a read-only root filesystem, one size-capped in-memory scratch
/// directory, no Linux capabilities, no privilege escalation, a non-root user, and hard memory,
/// CPU and process limits. The document goes in on stdin and the result comes out on stdout, so
/// nothing on the host is mounted. The container is removed when the render ends, which takes
/// every process LibreOffice started with it. Limits the host enforces are reported with the
/// same stable failure codes as <see cref="IDocumentRenderer"/> implementations in the library.
/// </remarks>
public sealed class SandboxedDocumentRenderer : IDocumentRenderer
{
    private static readonly JsonSerializerOptions Json = new(JsonSerializerDefaults.Web);
    private static readonly HashSet<string> Extensions = new(StringComparer.OrdinalIgnoreCase) { ".docx", ".pptx", ".xlsx" };
    private readonly SandboxLimits _limits;

    /// <summary>Creates a renderer that runs the reference image under the given limits.</summary>
    public SandboxedDocumentRenderer(SandboxLimits? limits = null) => _limits = limits ?? new SandboxLimits();

    /// <summary>The <c>docker run</c> arguments for one render; exposed so tests can assert them.</summary>
    public IReadOnlyList<string> RunArguments(string containerName)
    {
        var arguments = new List<string>
        {
            "run", "--rm", "-i", "--name", containerName,
            "--network", "none",
            "--read-only",
            "--tmpfs", $"/tmp:rw,nosuid,nodev,size={_limits.ScratchBytes}",
            "--user", "10001:10001",
            "--cap-drop", "ALL",
            "--security-opt", "no-new-privileges",
            "--pids-limit", _limits.Pids.ToString(System.Globalization.CultureInfo.InvariantCulture),
            "--memory", _limits.MemoryBytes.ToString(System.Globalization.CultureInfo.InvariantCulture),
            "--memory-swap", _limits.MemoryBytes.ToString(System.Globalization.CultureInfo.InvariantCulture),
            "--cpus", _limits.Cpus.ToString(System.Globalization.CultureInfo.InvariantCulture),
        };
        foreach (var (name, value) in _limits.Environment)
        {
            arguments.Add("--env");
            arguments.Add($"{name}={value}");
        }
        arguments.Add(_limits.Image);
        return arguments;
    }

    /// <inheritdoc />
    public async Task<RenderResult> RenderAsync(Stream document, RenderOptions options, CancellationToken cancellationToken = default)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (options is null) throw new ArgumentNullException(nameof(options));
        var fileName = Path.GetFileName(options.FileName);
        if (string.IsNullOrWhiteSpace(fileName) || !Extensions.Contains(Path.GetExtension(fileName)))
            return RenderResult.Failure(RenderFailureCodes.InvalidRenderOptions, "FileName must end in .docx, .pptx, or .xlsx.");
        if (options.Timeout <= TimeSpan.Zero || options.MaximumInputBytes <= 0 || options.MaximumOutputBytes <= 0 ||
            options.MaximumPages <= 0 || options.MaximumWorkingSetBytes <= 0)
            return RenderResult.Failure(RenderFailureCodes.InvalidRenderOptions, "All renderer limits must be positive.");

        // The input bound is enforced before any container starts.
        var bytes = await ReadBoundedAsync(document, options.MaximumInputBytes, cancellationToken).ConfigureAwait(false);
        if (bytes is null)
            return RenderResult.Failure(RenderFailureCodes.InputLimitExceeded, "The document exceeded the configured input-size limit.");

        var name = "officeagent-render-" + Guid.NewGuid().ToString("N");
        var start = new ProcessStartInfo(_limits.DockerExecutable)
        {
            UseShellExecute = false,
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        foreach (var argument in RunArguments(name)) start.ArgumentList.Add(argument);

        Process process;
        try
        {
            process = Process.Start(start) ?? throw new InvalidOperationException();
        }
        catch (Exception ex) when (ex is System.ComponentModel.Win32Exception or InvalidOperationException)
        {
            return RenderResult.Failure(RenderFailureCodes.RendererUnavailable, "The container runtime could not be started.");
        }

        using (process)
        using (var deadline = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken))
        {
            deadline.CancelAfter(options.Timeout + _limits.StartupAllowance);
            // Base64 pages plus a small JSON envelope; anything larger is refused unread.
            var maximumResponse = options.MaximumOutputBytes * 4 / 3 + 1024 * 1024;
            var stderr = DrainAsync(process.StandardError.BaseStream);
            var stdout = ReadBoundedAsync(process.StandardOutput.BaseStream, maximumResponse, CancellationToken.None);
            var killed = false;
            try
            {
                var header = JsonSerializer.SerializeToUtf8Bytes(new
                {
                    fileName,
                    dpi = options.Dpi,
                    timeoutMilliseconds = (int)Math.Min(int.MaxValue, options.Timeout.TotalMilliseconds),
                    maximumInputBytes = options.MaximumInputBytes,
                    maximumWorkingSetBytes = options.MaximumWorkingSetBytes,
                    maximumPages = options.MaximumPages,
                    maximumOutputBytes = options.MaximumOutputBytes
                }, Json);
                try
                {
                    // One buffer for the header line: an awaited write followed by a separate
                    // single-byte write once let the newline arrive after the document bytes.
                    var input = process.StandardInput.BaseStream;
                    await input.WriteAsync(header.Append((byte)'\n').ToArray(), deadline.Token).ConfigureAwait(false);
                    await input.WriteAsync(bytes, deadline.Token).ConfigureAwait(false);
                    await input.FlushAsync(deadline.Token).ConfigureAwait(false);
                    process.StandardInput.Close();
                }
                catch (IOException)
                {
                    // The container exited before reading everything; its exit code says why.
                }
                await process.WaitForExitAsync(deadline.Token).ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                killed = true;
                await KillAsync(name).ConfigureAwait(false);
                try { process.Kill(entireProcessTree: true); } catch (InvalidOperationException) { }
                await process.WaitForExitAsync(CancellationToken.None).ConfigureAwait(false);
                if (cancellationToken.IsCancellationRequested) throw;
            }

            var response = await stdout.ConfigureAwait(false);
            await stderr.ConfigureAwait(false);
            if (killed)
                return RenderResult.Failure(RenderFailureCodes.RenderTimeout, "Rendering exceeded the configured time limit.");
            // 137 without a kill from this host is the kernel's OOM killer inside the container.
            if (process.ExitCode == 137)
                return RenderResult.Failure(RenderFailureCodes.MemoryLimitExceeded, "The render container exceeded its memory limit.");
            // 125 to 127 are Docker's own failures: no daemon, no image, or no entrypoint.
            if (process.ExitCode is 125 or 126 or 127)
                return RenderResult.Failure(RenderFailureCodes.RendererUnavailable, "The render container could not be started.");
            if (process.ExitCode != 0)
                return RenderResult.Failure(RenderFailureCodes.RendererFailed, $"The render container exited with code {process.ExitCode}.");
            if (response is null)
                return RenderResult.Failure(RenderFailureCodes.OutputLimitExceeded, "Rendering exceeded the configured output-size limit.");
            return Map(response, options);
        }
    }

    private static RenderResult Map(byte[] response, RenderOptions options)
    {
        WorkerResponse? parsed;
        try { parsed = JsonSerializer.Deserialize<WorkerResponse>(response, Json); }
        catch (JsonException) { parsed = null; }
        if (parsed is null)
            return RenderResult.Failure(RenderFailureCodes.RendererFailed, "The render container returned an unreadable result.");
        if (!parsed.Succeeded)
            // Only the stable code and the library's own message cross the boundary.
            return RenderResult.Failure(parsed.FailureCode ?? RenderFailureCodes.RendererFailed, parsed.Message ?? "");
        var pages = parsed.Pages ?? Array.Empty<byte[]>();
        if (pages.Length == 0)
            return RenderResult.Failure(RenderFailureCodes.RendererOutputMissing, "The render container returned no pages.");
        if (pages.Length > options.MaximumPages)
            return RenderResult.Failure(RenderFailureCodes.PageLimitExceeded, "Rendering exceeded the configured page-count limit.");
        if (pages.Sum(page => (long)page.Length) > options.MaximumOutputBytes)
            return RenderResult.Failure(RenderFailureCodes.OutputLimitExceeded, "Rendering exceeded the configured output-size limit.");
        return RenderResult.Success(pages.Select((content, index) => new RenderedPage { PageNumber = index + 1, Content = content }).ToArray());
    }

    private async Task KillAsync(string containerName)
    {
        var kill = new ProcessStartInfo(_limits.DockerExecutable) { UseShellExecute = false, RedirectStandardOutput = true, RedirectStandardError = true };
        kill.ArgumentList.Add("kill");
        kill.ArgumentList.Add(containerName);
        try
        {
            using var process = Process.Start(kill);
            if (process is not null) await process.WaitForExitAsync().ConfigureAwait(false);
        }
        catch (System.ComponentModel.Win32Exception) { }
    }

    private static async Task<byte[]?> ReadBoundedAsync(Stream source, long maximumBytes, CancellationToken cancellationToken)
    {
        using var buffer = new MemoryStream();
        var chunk = new byte[81920];
        int read;
        while ((read = await source.ReadAsync(chunk, cancellationToken).ConfigureAwait(false)) > 0)
        {
            buffer.Write(chunk, 0, read);
            if (buffer.Length > maximumBytes)
            {
                // Keep draining so the writer is not blocked, but keep nothing.
                while (await source.ReadAsync(chunk, cancellationToken).ConfigureAwait(false) > 0) { }
                return null;
            }
        }
        return buffer.ToArray();
    }

    private static async Task DrainAsync(Stream source)
    {
        // Container stderr can carry document text or paths; it is read and discarded, never logged.
        var chunk = new byte[4096];
        try { while (await source.ReadAsync(chunk).ConfigureAwait(false) > 0) { } }
        catch (IOException) { }
    }

    private sealed record WorkerResponse(bool Succeeded, string? FailureCode, string? Message, byte[][]? Pages);
}
