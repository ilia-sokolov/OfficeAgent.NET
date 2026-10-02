# PowerPointReviewEdits

Edits a deck for review, and shows what OfficeAgent does when asked for something the file
format cannot hold.

Word records tracked changes in the document (`w:ins`, `w:del`), so a reviewer can accept or
reject each one. PowerPoint files have no equivalent markup. Asked for a tracked deck edit,
OfficeAgent refuses before writing anything, rather than producing highlighted or
struck-through text that only looks like a redline. This sample shows that refusal and then
the workable route:

1. **The refusal.** The edits are previewed with `Mode = ChangeMode.Tracked`. The preview is
   invalid, with `invalid-operation` and the engine's own suggestion: re-issue with `Direct`,
   or add a comment.
2. **Direct edits plus a record.** The same edits are applied with `ChangeMode.Direct`, and a
   comment on each changed slide records what changed from what.
3. **Accept or reject in PowerPoint.** The original file is left untouched, so a reviewer can
   open it in desktop PowerPoint, choose **Review > Compare**, and select the edited deck to
   step through the changes.

It uses the direct API (`new OfficeAgentClient(new PowerPointModule())` over a
`StreamHandle`), with no storage provider or dependency injection.

## Run it

From the repository root:

```powershell
dotnet run --project samples/PowerPointReviewEdits
```

This reads [`samples/documents/quarterly-review.pptx`](../documents/quarterly-review.pptx)
and writes `quarterly-review-edited.pptx` to the current directory. To use other files:

```powershell
dotnet run --project samples/PowerPointReviewEdits -- <input.pptx> <output.pptx>
```

Expected output:

```text
Tracked edit valid: False
  invalid-operation: PowerPoint has no tracked-changes representation, so mode 'Tracked' cannot be honoured. Re-issue this operation with mode 'Direct', or add a comment to record the intent.
  invalid-operation: PowerPoint has no tracked-changes representation, so mode 'Tracked' cannot be honoured. Re-issue this operation with mode 'Direct', or add a comment to record the intent.

Direct edits applied: 3 change(s), written to quarterly-review-edited.pptx
  text present: True  'Payment within forty-five days'
  text present: True  'Fixed fee: 61,500'
  comment: slide 2: Deal Desk - "Edited for review: 'Payment within thirt."
```

Slide 2 of the output reads "Payment within forty-five days" and "Fixed fee: 61,500", and
carries one comment by Deal Desk that lists both changes.

## What to take from it

- Set `Mode = ChangeMode.Direct` on deck edits made through the .NET API. A `ChangeTextOp`
  defaults to `Tracked`, which a deck refuses. See
  [change mode](../../docs/document-plans.md#change-mode).
- A deck comment targets the slide node, `slide#{slideId}`. The slide id is the first segment
  of a deck paragraph id (`slide{slideId}/shape{shapeId}/p{n}`). See
  [PowerPoint addressing](../../docs/powerpoint.md#addressing).
- If reviewers must accept or reject each change inside the file itself, the deliverable has
  to be a Word document. No library can write PowerPoint revisions, because the format does
  not define them.
