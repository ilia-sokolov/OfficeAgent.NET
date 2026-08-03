using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeAgent.PowerPoint;

/// <summary>
/// Inserts a new table onto a slide.
/// </summary>
/// <remarks>
/// Word anchors a new table to a paragraph, because a Word table is part of the document
/// flow. A slide has no flow - shapes are positioned absolutely - so the target here is
/// the slide itself, addressed as <c>{ "kind": "slide", "path": "slide#256" }</c>. The
/// frame is placed below the lowest existing shape so a new table does not land on top of
/// the slide's content.
/// </remarks>
internal sealed class SlideInsertTableHandler : IOperationHandler
{
    /// <inheritdoc />
    public bool CanHandle(PlanOperation operation) =>
        operation is InsertTableOp { Target: NodeAnchor { Kind: "slide" } };

    /// <inheritdoc />
    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (InsertTableOp)operation;
        var anchor = (NodeAnchor)op.Target!;

        var slide = ResolveSlide(context, anchor);
        if (slide is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No slide with path '{anchor.Path}'. Slide paths come from inspect_document.nodes.",
                anchor));

        var rowCount = op.Table.Rows.Count + (op.Table.Headers.Count > 0 ? 1 : 0);
        if (rowCount == 0)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "insertTable needs at least one header or row.", anchor));

        var columnCount = Math.Max(
            op.Table.Headers.Count,
            op.Table.Rows.Count > 0 ? op.Table.Rows.Max(r => r.Count) : 0);

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "insertTable",
            Before = string.Empty,
            After = $"[table {rowCount}×{columnCount}]",
            Context = $"slide {slide.Number}",
            BlastRadius = 1
        });
    }

    /// <inheritdoc />
    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (InsertTableOp)operation;
        var anchor = (NodeAnchor)op.Target!;

        var slide = ResolveSlide(context, anchor)
            ?? throw new InvalidOperationException($"Slide '{anchor.Path}' vanished before apply.");

        var shapeId = PowerPointModel.NextShapeId(slide.Part);
        var frame = SlideTableBuilder.BuildFrame(op.Table, shapeId, LowestEdge(slide));

        var tree = slide.Part.Slide.CommonSlideData?.ShapeTree
            ?? throw new InvalidOperationException($"Slide '{anchor.Path}' has no shape tree.");
        tree.Append(frame);
    }

    private static SlideRef? ResolveSlide(ApplyContext context, NodeAnchor anchor) =>
        SlideNodeProvider.TryParseSlideId(anchor.Path, out var slideId)
            ? PowerPointModel.Slide(context.Package, slideId)
            : null;

    /// <summary>
    /// The Y coordinate just below everything already on the slide, so an inserted table
    /// does not overlap existing shapes. Falls back to a sensible top margin on an empty
    /// slide.
    /// </summary>
    private static long LowestEdge(SlideRef slide)
    {
        long lowest = 0;
        foreach (var transform in slide.Part.Slide.Descendants<A.Transform2D>())
        {
            var y = transform.Offset?.Y?.Value ?? 0L;
            var height = transform.Extents?.Cy?.Value ?? 0L;
            if (y + height > lowest) lowest = y + height;
        }
        foreach (var transform in slide.Part.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Transform>())
        {
            var y = transform.Offset?.Y?.Value ?? 0L;
            var height = transform.Extents?.Cy?.Value ?? 0L;
            if (y + height > lowest) lowest = y + height;
        }

        // A quarter-inch gap under the lowest shape, or an inch down on a bare slide.
        return lowest > 0 ? lowest + 228600L : 914400L;
    }
}

/// <summary>Removes an entire table, frame and all, from the slide that holds it.</summary>
internal sealed class SlideRemoveTableHandler : IOperationHandler
{
    /// <inheritdoc />
    public bool CanHandle(PlanOperation operation) =>
        operation is RemoveTableOp { Target: NodeAnchor { Kind: "table" } };

    /// <inheritdoc />
    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var anchor = (NodeAnchor)operation.Target!;
        var table = SlideTableNodeProvider.Locate(anchor.Path, new PowerPointObjectMap(context.Package));
        if (table is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No table with path '{anchor.Path}'. Table paths come from inspect_document.nodes.",
                anchor));

        var rows = SlideTableBuilder.Rows(table.Table).Count;
        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "removeTable",
            Before = $"[table {rows}×{SlideTableBuilder.ColumnCount(table.Table)}]",
            After = string.Empty,
            Context = $"slide {table.Slide.Number}",
            // Removing a table takes every cell with it, which the report should say.
            BlastRadius = rows
        });
    }

    /// <inheritdoc />
    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var anchor = (NodeAnchor)operation.Target!;
        var table = SlideTableNodeProvider.Locate(anchor.Path, new PowerPointObjectMap(context.Package))
            ?? throw new InvalidOperationException($"Table '{anchor.Path}' vanished before apply.");

        // The frame is the shape; removing only the a:tbl would leave an empty frame.
        table.Frame.Remove();
    }
}

/// <summary>Appends or inserts rows in an existing slide table.</summary>
internal sealed class SlideInsertTableRowsHandler : IOperationHandler
{
    /// <inheritdoc />
    public bool CanHandle(PlanOperation operation) =>
        operation is InsertTableRowsOp { Target: NodeAnchor { Kind: "table" } };

    /// <inheritdoc />
    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (InsertTableRowsOp)operation;
        var anchor = (NodeAnchor)op.Target!;

        var table = SlideTableNodeProvider.Locate(anchor.Path, new PowerPointObjectMap(context.Package));
        if (table is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No table with path '{anchor.Path}'.", anchor));
        if (op.Rows.Count == 0)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "insertTableRows needs at least one row.", anchor));

        var existing = SlideTableBuilder.Rows(table.Table);
        if (op.Position is TablePosition.Before or TablePosition.After &&
            SlideTableBuilder.Normalize(op.RowIndex, existing.Count) < 0)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"Row index {op.RowIndex} is outside the table's {existing.Count} row(s).", anchor));

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "insertTableRows",
            Before = string.Empty,
            After = $"[+{op.Rows.Count} row(s)]",
            Context = $"slide {table.Slide.Number}, {op.Position.ToString().ToLowerInvariant()}",
            BlastRadius = op.Rows.Count
        });
    }

    /// <inheritdoc />
    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (InsertTableRowsOp)operation;
        var anchor = (NodeAnchor)op.Target!;

        var table = SlideTableNodeProvider.Locate(anchor.Path, new PowerPointObjectMap(context.Package))
            ?? throw new InvalidOperationException($"Table '{anchor.Path}' vanished before apply.");

        var columnCount = SlideTableBuilder.ColumnCount(table.Table);
        var existing = SlideTableBuilder.Rows(table.Table);
        var built = op.Rows.Select(r => SlideTableBuilder.BuildRow(r, columnCount)).ToList();

        switch (op.Position)
        {
            case TablePosition.Start:
                var first = existing.FirstOrDefault();
                foreach (var row in built)
                    if (first is null) table.Table.Append(row);
                    else first.InsertBeforeSelf(row);
                break;

            case TablePosition.Before:
            case TablePosition.After:
                var index = SlideTableBuilder.Normalize(op.RowIndex, existing.Count);
                var pivot = existing[index];
                if (op.Position == TablePosition.Before)
                    foreach (var row in built) pivot.InsertBeforeSelf(row);
                else
                    // Reversed so the supplied order survives repeated insert-after.
                    foreach (var row in Enumerable.Reverse(built)) pivot.InsertAfterSelf(row);
                break;

            default:
                foreach (var row in built) table.Table.Append(row);
                break;
        }
    }
}

/// <summary>Removes rows from an existing slide table.</summary>
internal sealed class SlideRemoveTableRowsHandler : IOperationHandler
{
    /// <inheritdoc />
    public bool CanHandle(PlanOperation operation) =>
        operation is RemoveTableRowsOp { Target: NodeAnchor { Kind: "table" } };

    /// <inheritdoc />
    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (RemoveTableRowsOp)operation;
        var anchor = (NodeAnchor)op.Target!;

        var table = SlideTableNodeProvider.Locate(anchor.Path, new PowerPointObjectMap(context.Package));
        if (table is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No table with path '{anchor.Path}'.", anchor));

        var doomed = Doomed(op, table.Table, out var invalid);
        if (invalid is { } bad)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"Row index {bad} is outside the table's {SlideTableBuilder.Rows(table.Table).Count} row(s).",
                anchor));

        if (doomed.Count == SlideTableBuilder.Rows(table.Table).Count && doomed.Count > 0)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "Removing every row would leave a table PowerPoint cannot render. Use removeTable instead.",
                anchor));

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "removeTableRows",
            Before = $"[{doomed.Count} row(s)]",
            After = string.Empty,
            Context = $"slide {table.Slide.Number}",
            BlastRadius = doomed.Count
        });
    }

    /// <inheritdoc />
    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (RemoveTableRowsOp)operation;
        var anchor = (NodeAnchor)op.Target!;

        var table = SlideTableNodeProvider.Locate(anchor.Path, new PowerPointObjectMap(context.Package))
            ?? throw new InvalidOperationException($"Table '{anchor.Path}' vanished before apply.");

        foreach (var row in Doomed(op, table.Table, out _))
            row.Remove();
    }

    /// <summary>
    /// The rows the operation would remove. An explicit index list wins; with no indices
    /// and <c>onlyIfEmpty</c>, every blank row goes.
    /// </summary>
    private static List<A.TableRow> Doomed(RemoveTableRowsOp op, A.Table table, out int? invalidIndex)
    {
        invalidIndex = null;
        var rows = SlideTableBuilder.Rows(table);

        if (op.RowIndices.Count == 0)
            return op.OnlyIfEmpty ? rows.Where(IsBlank).ToList() : new List<A.TableRow>();

        var doomed = new List<A.TableRow>();
        foreach (var requested in op.RowIndices)
        {
            var index = SlideTableBuilder.Normalize(requested, rows.Count);
            if (index < 0)
            {
                invalidIndex = requested;
                return doomed;
            }
            var row = rows[index];
            if (op.OnlyIfEmpty && !IsBlank(row)) continue;
            if (!doomed.Contains(row)) doomed.Add(row);
        }
        return doomed;
    }

    private static bool IsBlank(A.TableRow row) =>
        row.Elements<A.TableCell>().All(c => string.IsNullOrWhiteSpace(SlideTableBuilder.TextOf(c)));
}

/// <summary>Inserts columns into an existing slide table, keeping the grid in step.</summary>
internal sealed class SlideInsertTableColumnsHandler : IOperationHandler
{
    /// <inheritdoc />
    public bool CanHandle(PlanOperation operation) =>
        operation is InsertTableColumnsOp { Target: NodeAnchor { Kind: "table" } };

    /// <inheritdoc />
    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (InsertTableColumnsOp)operation;
        var anchor = (NodeAnchor)op.Target!;

        var table = SlideTableNodeProvider.Locate(anchor.Path, new PowerPointObjectMap(context.Package));
        if (table is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No table with path '{anchor.Path}'.", anchor));
        if (op.Columns.Count == 0)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "insertTableColumns needs at least one column.", anchor));
        if (table.Table.TableGrid is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"The table at '{anchor.Path}' has no column grid, so columns cannot be placed. " +
                "Recreate it with insertTable.", anchor));

        var columnCount = SlideTableBuilder.ColumnCount(table.Table);
        if (op.Position is TablePosition.Before or TablePosition.After &&
            SlideTableBuilder.Normalize(op.ColumnIndex, columnCount) < 0)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"Column index {op.ColumnIndex} is outside the table's {columnCount} column(s).", anchor));

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "insertTableColumns",
            Before = string.Empty,
            After = $"[+{op.Columns.Count} column(s)]",
            Context = $"slide {table.Slide.Number}, {op.Position.ToString().ToLowerInvariant()}",
            BlastRadius = op.Columns.Count
        });
    }

    /// <inheritdoc />
    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (InsertTableColumnsOp)operation;
        var anchor = (NodeAnchor)op.Target!;

        var table = SlideTableNodeProvider.Locate(anchor.Path, new PowerPointObjectMap(context.Package))
            ?? throw new InvalidOperationException($"Table '{anchor.Path}' vanished before apply.");

        var grid = table.Table.TableGrid
            ?? throw new InvalidOperationException($"Table '{anchor.Path}' has no column grid.");
        var columns = grid.Elements<A.GridColumn>().ToList();
        var rows = SlideTableBuilder.Rows(table.Table);

        var at = op.Position switch
        {
            TablePosition.Start => 0,
            TablePosition.Before => SlideTableBuilder.Normalize(op.ColumnIndex, columns.Count),
            TablePosition.After => SlideTableBuilder.Normalize(op.ColumnIndex, columns.Count) + 1,
            _ => columns.Count
        };

        // New columns share the width of the one they sit beside, so the table keeps its
        // overall width rather than growing off the slide. Measured against the original
        // neighbour, before any insertion shifts what sits there.
        var width = columns.Count > 0
            ? columns[Math.Min(Math.Max(at - 1, 0), columns.Count - 1)].Width?.Value ?? 1000000L
            : 1000000L;

        for (var i = 0; i < op.Columns.Count; i++)
        {
            var insertAt = at + i;

            // Both the grid and the cells must be re-read each time round: an insertion
            // shifts every later position, and using a list captured before the loop for
            // one but not the other silently attaches column widths to the wrong data.
            var current = grid.Elements<A.GridColumn>().ToList();
            var column = new A.GridColumn { Width = width };
            if (insertAt >= current.Count) grid.Append(column);
            else current[insertAt].InsertBeforeSelf(column);

            var data = op.Columns[i];
            for (var r = 0; r < rows.Count; r++)
            {
                var cell = SlideTableBuilder.BuildCell(r < data.Count ? data[r] : string.Empty);
                var cells = rows[r].Elements<A.TableCell>().ToList();
                if (insertAt >= cells.Count) rows[r].Append(cell);
                else cells[insertAt].InsertBeforeSelf(cell);
            }
        }
    }
}

/// <summary>Removes columns from an existing slide table, keeping the grid in step.</summary>
internal sealed class SlideRemoveTableColumnsHandler : IOperationHandler
{
    /// <inheritdoc />
    public bool CanHandle(PlanOperation operation) =>
        operation is RemoveTableColumnsOp { Target: NodeAnchor { Kind: "table" } };

    /// <inheritdoc />
    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (RemoveTableColumnsOp)operation;
        var anchor = (NodeAnchor)op.Target!;

        var table = SlideTableNodeProvider.Locate(anchor.Path, new PowerPointObjectMap(context.Package));
        if (table is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No table with path '{anchor.Path}'.", anchor));

        if (table.Table.TableGrid is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"The table at '{anchor.Path}' has no column grid, so there are no columns to remove.",
                anchor));

        var columnCount = SlideTableBuilder.ColumnCount(table.Table);
        var doomed = Doomed(op, columnCount, out var invalid);
        if (invalid is { } bad)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"Column index {bad} is outside the table's {columnCount} column(s).", anchor));

        if (doomed.Count == columnCount && doomed.Count > 0)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "Removing every column would leave a table PowerPoint cannot render. Use removeTable instead.",
                anchor));

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "removeTableColumns",
            Before = $"[{doomed.Count} column(s)]",
            After = string.Empty,
            Context = $"slide {table.Slide.Number}",
            BlastRadius = doomed.Count
        });
    }

    /// <inheritdoc />
    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (RemoveTableColumnsOp)operation;
        var anchor = (NodeAnchor)op.Target!;

        var table = SlideTableNodeProvider.Locate(anchor.Path, new PowerPointObjectMap(context.Package))
            ?? throw new InvalidOperationException($"Table '{anchor.Path}' vanished before apply.");

        var grid = table.Table.TableGrid
            ?? throw new InvalidOperationException($"Table '{anchor.Path}' has no column grid.");
        var columns = grid.Elements<A.GridColumn>().ToList();
        var doomed = Doomed(op, columns.Count, out _);

        // Descending, so removing one index does not shift the ones still to go.
        foreach (var index in doomed.OrderByDescending(i => i))
        {
            if (index < columns.Count) columns[index].Remove();
            foreach (var row in SlideTableBuilder.Rows(table.Table))
            {
                var cells = row.Elements<A.TableCell>().ToList();
                if (index < cells.Count) cells[index].Remove();
            }
        }
    }

    private static List<int> Doomed(RemoveTableColumnsOp op, int columnCount, out int? invalidIndex)
    {
        invalidIndex = null;
        var doomed = new List<int>();
        foreach (var requested in op.ColumnIndices)
        {
            var index = SlideTableBuilder.Normalize(requested, columnCount);
            if (index < 0)
            {
                invalidIndex = requested;
                return doomed;
            }
            if (!doomed.Contains(index)) doomed.Add(index);
        }
        return doomed;
    }
}
