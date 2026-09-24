using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Word;

// Liability Cap Modification Test
//
// This sample demonstrates editing a value (like a liability cap) in a document
// that may be within a commented region. The challenge is to:
// 1. Locate the liability cap value in the document
// 2. Modify the value
// 3. Preserve any comments intact
// 4. Track the change natively
//
// Run with:
//   dotnet run --project samples/LiabilityCap -- <input.docx> <output.docx> <old_cap> <new_cap>
// Example:
//   dotnet run --project samples/LiabilityCap -- services-agreement.docx output.docx "thirty days" "forty-five days"

if (args.Length != 4)
{
    Console.Error.WriteLine("Usage: LiabilityCap <input.docx> <output.docx> <old_cap> <new_cap>");
    Console.Error.WriteLine("Example: LiabilityCap contract.docx output.docx \"$500,000\" \"$1,000,000\"");
    return 1;
}

var input = args[0];
var output = args[1];
var oldCap = args[2];
var newCap = args[3];

// Verify input exists
if (!File.Exists(input))
{
    Console.Error.WriteLine($"Input file not found: {input}");
    return 1;
}

// A dedicated storage root for this run
var storageRoot = Path.Combine(Path.GetTempPath(), $"liabilitycap-{Guid.NewGuid():N}");
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
    // ── 0. Stage the input and register the document ─────────────────────
    var stagedDirectory = Path.Combine(storageRoot, Guid.NewGuid().ToString("N"));
    Directory.CreateDirectory(stagedDirectory);
    var stagedInput = Path.Combine(stagedDirectory, Path.GetFileName(input));
    File.Copy(input, stagedInput, overwrite: true);
    var document = await client.RegisterAsync("local", stagedInput);
    log.LogInformation("Registered document. documentId={DocumentId}", document.ItemId);

    // ── 1. Inspect the document ──────────────────────────────────────────
    var inspect = await client.InspectAsync("local", document.ItemId);
    log.LogInformation("Loaded {Paragraphs} paragraphs, {Anchors} anchors",
        inspect.Paragraphs.Count, inspect.Anchors.Count);

    // ── 2. Find the old cap value ────────────────────────────────────────
    var hits = await client.FindAsync("local", document.ItemId, new FindQuery(oldCap));
    if (hits.Count == 0)
    {
        log.LogError("Could not find the value: {OldValue}", oldCap);
        return 4;
    }

    log.LogInformation("Found {Count} occurrence(s) of '{OldValue}'", hits.Count, oldCap);

    var hit = hits[0];
    log.LogInformation("  Using first occurrence for replacement");

    // ── 3. Create a plan to replace the value with tracked change ────
    var plan = new DocumentPlan
    {
        Snapshot = inspect.Snapshot, // Enable drift detection
        Revision = new RevisionMetadata { Author = "LiabilityCap Editor" },
        Operations = new PlanOperation[]
        {
            new ChangeTextOp
            {
                Target = hit.Anchor,
                With = newCap,
                Mode = ChangeMode.Tracked
            }
        }
    };

    // ── 4. Preview the changes ──────────────────────────────────────────
    var preview = await client.PreviewAsync("local", document.ItemId, plan);
    if (!preview.IsValid)
    {
        foreach (var e in preview.Errors)
            log.LogError("Preview failed [{Code}]: {Message}", e.Code, e.Message);
        return 3;
    }

    log.LogInformation("Preview valid. Ready to commit.");

    // ── 5. Commit the changes ──────────────────────────────────────────
    var commit = await client.CommitAsync("local", document.ItemId, plan);
    if (!commit.Committed)
    {
        foreach (var e in commit.Report.Errors)
            log.LogError("{Code}: {Message}", e.Code, e.Message);
        return 2;
    }

    log.LogInformation("Changes committed. {ChangeCount} change(s)", commit.Report.Changes.Count);

    // ── 6. Verify the new value exists ──────────────────────────────────
    var newHits = await client.FindAsync("local", document.ItemId, new FindQuery(newCap));
    if (newHits.Count == 0)
    {
        log.LogError("New value not found in document: {NewValue}", newCap);
        return 5;
    }

    log.LogInformation("New value confirmed in document: {NewValue}", newCap);

    // ── 7. Save the result ───────────────────────────────────────────────
    using var saved = await client.OpenReadAsync(commit.Document!);
    using var outFile = File.Create(output);
    await saved.Stream.CopyToAsync(outFile);
    log.LogInformation("Wrote {Output}", output);

    return 0;
}
catch (Exception ex)
{
    log.LogError(ex, "Operation failed");
    return 2;
}
finally
{
    if (Directory.Exists(storageRoot))
        Directory.Delete(storageRoot, recursive: true);
}
