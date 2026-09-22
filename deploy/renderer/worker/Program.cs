using System.Diagnostics;
using System.Text.Json;
using OfficeAgent.Abstractions;
using OfficeAgent.Rendering;

// The render worker that runs inside the sandbox container.
//
// Protocol, chosen so the container needs no network and no mounts. The host writes one line of
// JSON (a RenderRequest), then the document bytes, then closes stdin. The worker writes one JSON
// object (a RenderResponse) to stdout and exits 0. Renderer stdout and stderr are never
// forwarded: they can carry document text and paths, and the renderer already drains them
// without keeping them.

(RenderRequest Header, byte[]? Document) request;
try
{
    request = await ReadRequestAsync(Console.OpenStandardInput());
}
catch (Exception ex) when (ex is JsonException or InvalidDataException)
{
    // A malformed request is refused with a distinct exit code and no stack trace. The host
    // reports it as renderer-failed, and nothing from the request is echoed back.
    Console.Error.WriteLine("officeagent-render-worker: malformed request");
    return 2;
}
var response = await RenderAsync(request);
await using var stdout = Console.OpenStandardOutput();
await JsonSerializer.SerializeAsync(stdout, response, Protocol.Json);
return 0;

static async Task<(RenderRequest Header, byte[]? Document)> ReadRequestAsync(Stream input)
{
    var header = new List<byte>();
    while (true)
    {
        var next = input.ReadByte();
        if (next < 0 || next == '\n') break;
        header.Add((byte)next);
        if (header.Count > 64 * 1024) throw new InvalidDataException("Request header too long.");
    }
    var parsed = JsonSerializer.Deserialize<RenderRequest>(header.ToArray(), Protocol.Json)
                 ?? throw new InvalidDataException("Missing request header.");

    // Read at most one byte past the limit, so an oversized document is refused without being
    // buffered whole.
    using var buffer = new MemoryStream();
    var chunk = new byte[81920];
    int read;
    while ((read = await input.ReadAsync(chunk)) > 0)
    {
        buffer.Write(chunk, 0, read);
        if (buffer.Length > parsed.MaximumInputBytes) return (parsed, null);
    }
    return (parsed, buffer.ToArray());
}

static async Task<RenderResponse> RenderAsync((RenderRequest Header, byte[]? Document) request)
{
    var (header, document) = request;
    var backend = new BackendInfo(Version("soffice", "--version"), Version("pdftoppm", "-v"));
    if (document is null)
        return RenderResponse.Failed(RenderFailureCodes.InputLimitExceeded,
            "The document exceeded the configured input-size limit.", backend);

    var result = await new LibreOfficeDocumentRenderer().RenderAsync(
        new MemoryStream(document, writable: false),
        new RenderOptions
        {
            FileName = header.FileName,
            Dpi = header.Dpi,
            Timeout = TimeSpan.FromMilliseconds(header.TimeoutMilliseconds),
            MaximumInputBytes = header.MaximumInputBytes,
            MaximumWorkingSetBytes = header.MaximumWorkingSetBytes,
            MaximumPages = header.MaximumPages,
            MaximumOutputBytes = header.MaximumOutputBytes
        });

    var audit = Environment.GetEnvironmentVariable("OFFICEAGENT_SANDBOX_AUDIT") == "1"
        ? Audit(document, header.FileName)
        : null;
    return result.Succeeded
        ? new RenderResponse(true, null, null, result.Pages.Select(page => page.Content).ToArray(), backend, audit)
        : RenderResponse.Failed(result.FailureCode!, result.Message ?? "", backend, audit);
}

// Test-only evidence, enabled by OFFICEAGENT_SANDBOX_AUDIT=1: whether a macro in the document
// ran (it would leave a marker), and the text LibreOffice produced, so a test can see whether a
// linked file's content was pulled in. Never enabled in the deployment recipe.
static AuditInfo Audit(byte[] document, string fileName)
{
    var marker = File.Exists("/tmp/officeagent-macro-marker");
    var directory = Directory.CreateTempSubdirectory("officeagent-audit-");
    try
    {
        var input = Path.Combine(directory.FullName, Path.GetFileName(fileName));
        File.WriteAllBytes(input, document);
        Run("soffice", "--headless", "--nologo", "--nodefault", "--norestore", "--nolockcheck",
            "-env:UserInstallation=file://" + Path.Combine(directory.FullName, "profile"),
            "--convert-to", "pdf", "--outdir", directory.FullName, input);
        var pdf = Directory.EnumerateFiles(directory.FullName, "*.pdf").FirstOrDefault();
        var text = pdf is null ? "" : Run("pdftotext", pdf, "-");
        return new AuditInfo(marker || File.Exists("/tmp/officeagent-macro-marker"), text);
    }
    finally
    {
        directory.Delete(recursive: true);
    }
}

static string Version(string executable, string argument)
{
    var output = Run(executable, argument).Trim();
    return output.Split('\n').FirstOrDefault()?.Trim() ?? "";
}

static string Run(string executable, params string[] arguments)
{
    var start = new ProcessStartInfo(executable)
    {
        RedirectStandardOutput = true,
        RedirectStandardError = true,
        UseShellExecute = false
    };
    foreach (var argument in arguments) start.ArgumentList.Add(argument);
    try
    {
        using var process = Process.Start(start)!;
        var stdout = process.StandardOutput.ReadToEndAsync();
        var stderr = process.StandardError.ReadToEndAsync();
        if (!process.WaitForExit(60_000)) process.Kill(entireProcessTree: true);
        // pdftoppm prints its version on stderr.
        return stdout.Result.Length > 0 ? stdout.Result : stderr.Result;
    }
    catch (System.ComponentModel.Win32Exception)
    {
        return "";
    }
}

internal sealed record RenderRequest(
    string FileName,
    int Dpi,
    int TimeoutMilliseconds,
    long MaximumInputBytes,
    long MaximumWorkingSetBytes,
    int MaximumPages,
    long MaximumOutputBytes);

internal sealed record BackendInfo(string LibreOffice, string Poppler);

internal sealed record AuditInfo(bool MacroRan, string PdfText);

internal sealed record RenderResponse(
    bool Succeeded,
    string? FailureCode,
    string? Message,
    byte[][]? Pages,
    BackendInfo Backend,
    AuditInfo? Audit)
{
    public static RenderResponse Failed(string code, string message, BackendInfo backend, AuditInfo? audit = null) =>
        new(false, code, message, null, backend, audit);
}

internal static class Protocol
{
    public static readonly JsonSerializerOptions Json = new(JsonSerializerDefaults.Web);
}
