# Liability Cap Modification Sample

## Overview

This sample demonstrates modifying a value (such as a liability cap) within a Word document using OfficeAgent.NET, with a focus on preserving document structure including comments and comment ranges.

## Challenge

The advanced scenario involves:
1. Locating a specific value in the document (e.g., a liability cap amount in a table cell)
2. Modifying that value
3. Preserving any existing comments or comment ranges that overlap with the modified text
4. Recording the change as a tracked revision in Word

## Why This Is Hard

Comment ranges and text edits can conflict in low-level DOM APIs. A naive approach using OpenXML directly can:
- Corrupt comment range markers
- Lose change tracking information
- Damage document structure

OfficeAgent.NET handles this by:
- Maintaining an abstract representation of document structure
- Validating edits against a document snapshot before applying
- Applying operations atomically with all-or-nothing semantics
- Preserving unrelated document content untouched

## Solution: OfficeAgent.NET

OfficeAgent.NET is a .NET document-automation library built on the Open XML SDK that specializes in preserving document structure while making targeted edits. It uses:

1. **Inspect** — Build a structured map of the document
2. **Find** — Search for text with content verification
3. **Preview** — Validate edits before writing
4. **Commit** — Apply the complete plan atomically

## How to Run

### Prerequisites

- .NET 8 SDK or later
- A Word document (`.docx`) containing a value to modify

### Modify a Document

```bash
dotnet run --project samples/LiabilityCap -- <input.docx> <output.docx> <old_value> <new_value>
```

#### Example

Using the bundled sample contract:

```bash
dotnet run --project samples/LiabilityCap -- \
  samples/documents/services-agreement.docx \
  liability-output.docx \
  "thirty days" \
  "forty-five days"
```

### Output

The program will:
1. Register the input document
2. Inspect its structure
3. Find all occurrences of the old value
4. Preview the proposed change
5. Commit the change as a tracked revision
6. Save the output file

#### Example Output

```
info: Program[0] Registered document. documentId=<id>
info: Program[0] Loaded 25 paragraphs, 37 anchors
info: Program[0] Found 1 occurrence(s) of 'thirty days'
info: Program[0]   Using first occurrence for replacement
info: Program[0] Preview valid. Ready to commit.
info: Program[0] Changes committed. 1 change(s)
info: Program[0] New value confirmed in document: forty-five days
info: Program[0] Wrote liability-output.docx
```

## Verification

Open the output document in Microsoft Word:
- The modified value appears as a **tracked change** you can accept or reject
- All other content remains unchanged
- Any existing comments remain intact and visible
- Document structure (tables, lists, formatting) is preserved

## Key Implementation Details

### Program Structure

```csharp
// 1. Register document
var document = await client.RegisterAsync("local", stagedInput);

// 2. Inspect
var inspect = await client.InspectAsync("local", document.ItemId);

// 3. Find
var hits = await client.FindAsync("local", document.ItemId, 
  new FindQuery(oldValue));

// 4. Create plan
var plan = new DocumentPlan
{
    Snapshot = inspect.Snapshot,
    Revision = new RevisionMetadata { Author = "LiabilityCap Editor" },
    Operations = new PlanOperation[] {
        new ChangeTextOp {
            Target = hits[0].Anchor,
            With = newValue,
            Mode = ChangeMode.Tracked  // ← Preserves as redline
        }
    }
};

// 5. Preview
var preview = await client.PreviewAsync("local", document.ItemId, plan);

// 6. Commit
var commit = await client.CommitAsync("local", document.ItemId, plan);

// 7. Save
using var saved = await client.OpenReadAsync(commit.Document!);
await saved.Stream.CopyToAsync(File.Create(output));
```

### Why `ChangeMode.Tracked` Matters

- `Tracked` — Records the change as a Word revision (redline) that reviewers can accept/reject
- `Direct` — Applies the change without revision markup

For contract review workflows, `Tracked` is essential for audit trails and approval workflows.

## Files

- **LiabilityCap.csproj** — Project configuration
- **Program.cs** — Main application logic
- **VerifyModification.cs** — Utility to inspect document structure

## Troubleshooting

### "Could not find the value"

The old value must match **exactly** (case-sensitive by default). Check:
- Exact spelling and punctuation
- Whitespace (e.g., "thirty days" vs "thirty  days" with extra spaces)
- Hidden characters in the document

### "Preview failed"

This indicates a structural issue with the document or the requested change. Check the error code:
- `drift` — Document changed between inspect and preview
- `invalid-anchor` — The target anchor is no longer valid
- `contract-mismatch` — Document format mismatch

### Document opened in Word is read-only

Word may lock the file for editing if it was recently modified. Close and reopen it in Word, or try opening it directly from the file location.

## Next Steps

- Modify the `ChangeMode` to `Direct` for direct (non-tracked) edits
- Change `Revision.Author` to reflect your workflow's actor
- Loop over multiple occurrences by iterating `hits`
- Use `FindQuery(text, new MatchOptions { Regex = true })` for pattern matching

## References

- [OfficeAgent.NET Documentation](https://github.com/dotaction/OfficeAgent.NET#readme)
- [Document Plans](../../docs/document-plans.md)
- [Getting Started](../../docs/getting-started.md)
- [C# API Reference](../../docs/csharp-api.md)
