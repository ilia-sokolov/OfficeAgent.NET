using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.PowerPoint;

// Edits a deck for review. PowerPoint files have no tracked-change markup, so OfficeAgent
// refuses a tracked deck edit instead of faking one. This sample shows that refusal, then the
// workable route: apply the edits directly, record each one in a slide comment, and keep the
// original so a reviewer can run PowerPoint's Review > Compare to accept or reject them.
//
//   dotnet run --project samples/PowerPointReviewEdits
//   dotnet run --project samples/PowerPointReviewEdits -- <input.pptx> <output.pptx>

var input = args.Length > 0 ? args[0] : Path.Combine("samples", "documents", "quarterly-review.pptx");
var output = args.Length > 1 ? args[1] : "quarterly-review-edited.pptx";

var edits = new (string Before, string After)[]
{
    ("Payment within thirty days", "Payment within forty-five days"),
    ("Fixed fee: 58,000", "Fixed fee: 61,500"),
};

using var source = new MemoryStream(await File.ReadAllBytesAsync(input), writable: false);
var deck = new StreamHandle(source, Path.GetFileName(input));
var client = new OfficeAgentClient(new PowerPointModule());

var inspection = await client.InspectAsync(deck);

// Resolve every edit to a verified anchor before planning anything.
var anchors = new List<TextSpanAnchor>();
foreach (var (before, _) in edits)
{
    var hits = await client.FindAsync(deck, new FindQuery(before));
    if (hits.Count != 1 || hits[0].Anchor is not TextSpanAnchor span)
    {
        Console.Error.WriteLine($"Expected exactly one match for '{before}', found {hits.Count}.");
        return 1;
    }
    anchors.Add(span);
}

// 1. Ask for tracked changes. The deck refuses, before anything is written.
var tracked = new DocumentPlan
{
    Snapshot = inspection.Snapshot,
    Operations = edits.Select((e, i) => (PlanOperation)new ChangeTextOp
    {
        Target = anchors[i], With = e.After, Mode = ChangeMode.Tracked
    }).ToArray()
};

var refused = await client.PreviewAsync(deck, tracked);
Console.WriteLine($"Tracked edit valid: {refused.IsValid}");
foreach (var error in refused.Errors)
    Console.WriteLine($"  {error.Code}: {error.Message}");

// 2. Apply the same edits directly, and record each one in a comment on its slide so the
//    reviewer sees what changed without the original open beside it.
var commentsBySlide = edits
    .Select((e, i) => (Slide: SlideOf(anchors[i]), Line: $"'{e.Before}' changed to '{e.After}'"))
    .GroupBy(x => x.Slide)
    .Select(g => (PlanOperation)new CommentOp
    {
        Target = new NodeAnchor { Kind = "slide", Path = $"slide#{g.Key}" },
        Text = "Edited for review: " + string.Join("; ", g.Select(x => x.Line)) +
               $". Compare with {Path.GetFileName(input)} to accept or reject.",
        Author = "Deal Desk",
        Initials = "DD"
    });

var direct = new DocumentPlan
{
    Snapshot = inspection.Snapshot,
    Operations = edits.Select((e, i) => (PlanOperation)new ChangeTextOp
    {
        Target = anchors[i], With = e.After, Mode = ChangeMode.Direct
    }).Concat(commentsBySlide).ToArray()
};

var preview = await client.PreviewAsync(deck, direct);
if (!preview.IsValid)
{
    foreach (var error in preview.Errors)
        Console.Error.WriteLine($"{error.Code}: {error.Message}");
    return 2;
}

using var result = await client.CommitAsync(deck, direct);
if (!result.Committed)
{
    foreach (var error in result.Report.Errors)
        Console.Error.WriteLine($"{error.Code}: {error.Message}");
    return 3;
}

await result.SaveAsync(output);
Console.WriteLine();
Console.WriteLine($"Direct edits applied: {result.Report.Changes.Count} change(s), written to {output}");

// 3. Read the saved deck back: the new text and the review comment are both there.
using var saved = new MemoryStream(await File.ReadAllBytesAsync(output), writable: false);
var check = await client.InspectAsync(new StreamHandle(saved, Path.GetFileName(output)));
foreach (var (_, after) in edits)
    Console.WriteLine($"  text present: {check.Paragraphs.Any(p => p.Text == after)}  '{after}'");
foreach (var node in check.Nodes.Where(n => n.Kind == "comment"))
    Console.WriteLine($"  comment: {node.Summary}");

Console.WriteLine();
Console.WriteLine($"To accept or reject each edit, open {Path.GetFileName(input)} in PowerPoint,");
Console.WriteLine($"choose Review > Compare, and select {Path.GetFileName(output)}.");
return 0;

// A deck paragraph id has the form slide{slideId}/shape{shapeId}/p{n}.
static string SlideOf(TextSpanAnchor anchor) =>
    anchor.ParaId["slide".Length..anchor.ParaId.IndexOf('/')];
