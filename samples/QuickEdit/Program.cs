using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Word;

// Minimal sample: edit a .docx through OfficeAgent.NET's opaque-id storage flow.
//
// The provider is a registry of references - it stores the path you register, not
// the bytes themselves. The host copies the input into the connection's root,
// registers that path to receive an opaque document id, and drives inspect → find →
// preview → commit by that id. The result is read back out of storage and written
// to the caller's chosen output path.
//
// Run the bundled example:
//   dotnet run --project samples/QuickEdit -- samples/documents/services-agreement.docx quickedit-output.docx
// Or choose the exact text replacement:
//   dotnet run --project samples/QuickEdit -- <input.docx> <output.docx> <find> <replacement>

if (args.Length is not (2 or 4))
{
    Console.Error.WriteLine("Usage: QuickEdit <input.docx> <output.docx> [<find> <replacement>]");
    return 1;
}

var input = args[0];
var output = args[1];
var findText = args.Length == 4 ? args[2] : "within thirty days of receipt";
var replacementText = args.Length == 4 ? args[3] : "within forty-five days of receipt";

// A dedicated storage root for this run. Registered paths must live under it; the
// provider keeps only an id → path index inside .officeagent/, never the bytes.
var storageRoot = Path.Combine(Path.GetTempPath(), $"quickedit-{Guid.NewGuid():N}");
Directory.CreateDirectory(storageRoot);

using var host = Host.CreateDefaultBuilder()
    .ConfigureLogging(b => b.AddSimpleConsole(o => o.SingleLine = true))
    .ConfigureServices(s => s
        .AddWordFormat()
        .AddFileSystemDocumentProvider("local", storageRoot)
        .AddOfficeAgent())
    .Build();

var client = host.Services.GetRequiredService<OfficeAgentClient>();
var log = host.Services.GetRequiredService<ILogger<Program>>();

try
{
    // ── 0. Stage the input under the connection root and register the path;
    //       everything after is addressed by the returned opaque id. ─────────────
    var stagedDirectory = Path.Combine(storageRoot, Guid.NewGuid().ToString("N"));
    Directory.CreateDirectory(stagedDirectory);
    var stagedInput = Path.Combine(stagedDirectory, Path.GetFileName(input));
    File.Copy(input, stagedInput, overwrite: true);
    var document = await client.RegisterAsync("local", stagedInput);
    log.LogInformation("Registered document with storage. path={Path} documentId={DocumentId}", stagedInput, document.ItemId);

    // ── 1. Inspect ──────────────────────────────────────────────────────────────
    var inspect = await client.InspectAsync("local", document.ItemId);
    log.LogInformation("Loaded {Paragraphs} paragraphs, snapshot={Snapshot}",
        inspect.Paragraphs.Count, inspect.Snapshot.ETag[..8]);

    foreach (var p in inspect.Paragraphs.Take(3))
        Console.WriteLine($"  {p.ParaId}: {Trim(p.Text)}");

    // ── 2. Find an anchor, preview a tracked edit, then commit ───────────────────
    var hits = await client.FindAsync("local", document.ItemId, new FindQuery(findText));
    if (hits.Count == 0)
    {
        log.LogError("Could not find the exact text {FindText}; no output was written.", findText);
        return 4;
    }

    var plan = new DocumentPlan
    {
        Snapshot = inspect.Snapshot, // opt in to drift detection
        Revision = new RevisionMetadata { Author = "QuickEdit" },
        Operations = new PlanOperation[]
        {
            new ChangeTextOp { Target = hits[0].Anchor, With = replacementText, Mode = ChangeMode.Tracked }
        }
    };

    var preview = await client.PreviewAsync("local", document.ItemId, plan);
    if (!preview.IsValid)
    {
        foreach (var e in preview.Errors)
            log.LogError("Preview failed [{Code}]: {Message}", e.Code, e.Message);
        return 3;
    }

    var commit = await client.CommitAsync("local", document.ItemId, plan);
    if (!commit.Committed)
    {
        foreach (var e in commit.Report.Errors)
            log.LogError("{Code}: {Message}", e.Code, e.Message);
        return 2;
    }

    // ── 3. Read the saved version back out of storage and write it to disk ────────
    using var saved = await client.OpenReadAsync(commit.Document!);
    using var outFile = File.Create(output);
    await saved.Stream.CopyToAsync(outFile);
    log.LogInformation("Wrote {Output} ({Changes} change(s))", output, commit.Report.Changes.Count);
}
finally
{
    if (Directory.Exists(storageRoot)) Directory.Delete(storageRoot, recursive: true);
}

return 0;

static string Trim(string s) => s.Length <= 80 ? s : s[..77] + "...";
