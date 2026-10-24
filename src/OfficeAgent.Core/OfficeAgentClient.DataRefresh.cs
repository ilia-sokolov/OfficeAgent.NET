using System.Globalization;
using OfficeAgent.Abstractions;
using DocFormat = OfficeAgent.Abstractions.DocumentFormat;

namespace OfficeAgent.Core;

public sealed partial class OfficeAgentClient
{
    /// <summary>The most cells one binding reads from the workbook.</summary>
    public const int MaximumRefreshCells = 10000;

    /// <summary>
    /// Works out how to update a Word document or PowerPoint deck from a workbook: finds every
    /// <see cref="DataBindingMarker"/> in the document, reads the range each one names, and
    /// returns the plan that writes them. Nothing is applied.
    /// </summary>
    /// <remarks>
    /// <list type="bullet">
    /// <item>A Word content control or PowerPoint text shape bound to one cell takes that
    /// cell's displayed value, number format applied.</item>
    /// <item>A bound table is filled from its range cell by cell, growing or shrinking by rows.</item>
    /// <item>A bound chart takes its range with series in columns: the first row holds the
    /// series names, the first column the categories.</item>
    /// </list>
    /// A formula's stored result is used; a workbook saved by Excel always has one. If any
    /// binding cannot be read, the returned plan is empty and <see cref="DataRefreshPlan.IsComplete"/>
    /// is false, so a report is never left half refreshed.
    /// </remarks>
    /// <param name="document">The Word document or PowerPoint deck to refresh.</param>
    /// <param name="workbook">The Excel workbook that supplies the data.</param>
    public DataRefreshPlan PlanDataRefresh(byte[] document, byte[] workbook)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (workbook is null) throw new ArgumentNullException(nameof(workbook));

        var target = Inspect(document);
        if (target.Format is not (DocFormat.Word or DocFormat.PowerPoint))
            throw new ArgumentException(
                $"A data refresh updates a Word document or a PowerPoint deck, and this is {target.Format}.", nameof(document));

        var book = Inspect(workbook, new InspectOptions { Fidelity = Fidelity.Structure });
        if (book.Format != DocFormat.Excel)
            throw new ArgumentException(
                $"A data refresh reads an Excel workbook, and this is {book.Format}.", nameof(workbook));

        var bindings = new List<DataRefreshBinding>();
        var operations = new List<PlanOperation>();
        var reader = new RangeReader(this, workbook, book);

        foreach (var found in Discover(target))
        {
            if (found.Duplicate)
            {
                bindings.Add(Failed(found, null, DataRefreshErrorCodes.DuplicateBinding,
                    $"'{found.Marker}' marks more than one place. Office needs these names to be unique: add '#2' to " +
                    "the second, which the refresh ignores, such as 'xl:Summary!B2#2'."));
                continue;
            }

            var range = reader.Read(found.Reference, out var resolved, out var code, out var message);
            if (range is null)
            {
                bindings.Add(Failed(found, resolved, code!, message!));
                continue;
            }

            var (operation, summary, error, errorMessage) = found.Kind switch
            {
                DataBindingKind.Value => ValueOperation(found, range, resolved!),
                DataBindingKind.Table => TableOperation(found, range, resolved!),
                _ => ChartOperation(found, range, resolved!)
            };

            if (operation is null)
            {
                bindings.Add(Failed(found, resolved, error!, errorMessage!));
                continue;
            }

            operations.Add(operation);
            bindings.Add(new DataRefreshBinding
            {
                Kind = found.Kind,
                Marker = found.Marker,
                Reference = found.Reference,
                Target = found.Target,
                Range = resolved,
                Summary = summary
            });
        }

        var complete = bindings.Count > 0 && bindings.All(binding => binding.ErrorCode is null);
        return new DataRefreshPlan
        {
            // Bound to the state it was read from, so a document edited between planning and
            // applying is refused as stale rather than overwritten with a plan for its old self.
            Plan = new DocumentPlan
            {
                Snapshot = target.Snapshot,
                Operations = complete ? operations : Array.Empty<PlanOperation>()
            },
            Bindings = bindings
        };
    }

    /// <summary>Plans a refresh of a stored document from a stored workbook.</summary>
    public async Task<DataRefreshPlan> PlanDataRefreshAsync(
        DocumentReference document,
        DocumentReference workbook,
        CancellationToken cancellationToken = default)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (workbook is null) throw new ArgumentNullException(nameof(workbook));

        var documentBytes = await ReadBytesAsync(document, cancellationToken).ConfigureAwait(false);
        var workbookBytes = await ReadBytesAsync(workbook, cancellationToken).ConfigureAwait(false);
        return PlanDataRefresh(documentBytes, workbookBytes);
    }

    private async Task<byte[]> ReadBytesAsync(DocumentReference reference, CancellationToken cancellationToken)
    {
        using var content = await OpenReadAsync(reference, cancellationToken).ConfigureAwait(false);
        using var buffer = new MemoryStream();
        await content.Stream.CopyToAsync(buffer, 81920, cancellationToken).ConfigureAwait(false);
        return buffer.ToArray();
    }

    /// <summary>One binding marker found in the document.</summary>
    private sealed record FoundBinding(DataBindingKind Kind, string Marker, string Reference, Anchor Target, bool Duplicate);

    /// <summary>
    /// The bindings a document carries, in inspection order: content controls and text shapes
    /// for values, and named tables and charts.
    /// </summary>
    private static IEnumerable<FoundBinding> Discover(InspectResult target)
    {
        var slots = target.StructuralAnchors
            .Where(anchor => anchor.Kind is "contentControl" or "shapeName")
            .ToList();
        var counts = slots.GroupBy(anchor => anchor.Tag, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);
        var reported = new HashSet<string>(StringComparer.Ordinal);

        foreach (var slot in slots)
        {
            if (!DataBindingMarker.TryParse(slot.Tag, out var reference)) continue;
            var duplicate = counts[slot.Tag] > 1;
            if (duplicate && !reported.Add(slot.Tag)) continue;
            yield return new FoundBinding(DataBindingKind.Value, slot.Tag, reference, slot, duplicate);
        }

        foreach (var node in target.Nodes)
        {
            if (node.Kind is not ("table" or "chart") || node.Anchor is null ||
                !DataBindingMarker.TryParse(node.Name, out var reference)) continue;
            yield return new FoundBinding(
                node.Kind == "table" ? DataBindingKind.Table : DataBindingKind.Chart,
                node.Name!, reference, node.Anchor, Duplicate: false);
        }
    }

    private static (PlanOperation? Operation, string? Summary, string? Code, string? Message) ValueOperation(
        FoundBinding found, CellInfo?[][] range, string resolved)
    {
        if (range.Length != 1 || range[0].Length != 1)
            return (null, null, DataRefreshErrorCodes.NotOneCell,
                $"'{found.Marker}' shows one value, but {resolved} is {range.Length}×{range[0].Length} cells. Bind it to one cell.");

        var cell = range[0][0];
        if (Text(cell) is not { } value)
            return (null, null, DataRefreshErrorCodes.NotCalculated, NotCalculated(resolved));

        return (new FillOp
        {
            Target = found.Target,
            Value = value,
            Mode = ChangeMode.Direct
        }, value, null, null);
    }

    private static (PlanOperation? Operation, string? Summary, string? Code, string? Message) TableOperation(
        FoundBinding found, CellInfo?[][] range, string resolved)
    {
        var rows = new List<IReadOnlyList<string>>();
        foreach (var row in range)
        {
            var values = new List<string>();
            foreach (var cell in row)
            {
                if (Text(cell) is not { } value)
                    return (null, null, DataRefreshErrorCodes.NotCalculated, NotCalculated(cell!.Anchor.Address));
                values.Add(value);
            }

            rows.Add(values);
        }

        return (new FillTableOp { Target = found.Target, Rows = rows },
            $"{rows.Count} row(s) × {range[0].Length} column(s)", null, null);
    }

    private static (PlanOperation? Operation, string? Summary, string? Code, string? Message) ChartOperation(
        FoundBinding found, CellInfo?[][] range, string resolved)
    {
        if (range.Length < 2 || range[0].Length < 2)
            return (null, null, DataRefreshErrorCodes.NotChartData,
                $"{resolved} is too small for a chart. Bind a range whose first row names the series and whose first " +
                "column names the categories, with at least one value.");

        var categories = new List<string>();
        var series = new List<ChartSeries>();
        for (var column = 1; column < range[0].Length; column++)
        {
            var values = new List<double?>();
            for (var row = 1; row < range.Length; row++)
            {
                var cell = range[row][column];
                if (cell is not null && cell.Formula is not null && cell.FormattedValue is null)
                    return (null, null, DataRefreshErrorCodes.NotCalculated, NotCalculated(cell.Anchor.Address));
                if (cell is null || string.IsNullOrEmpty(cell.DisplayValue))
                {
                    values.Add(null);
                    continue;
                }

                // A number is stored as itself; a shared string stores an index, so its raw and
                // displayed values differ.
                if (cell.RawValue != cell.DisplayValue ||
                    !double.TryParse(cell.RawValue, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
                    return (null, null, DataRefreshErrorCodes.NotChartData,
                        $"Cell {cell.Anchor.Address} holds '{cell.DisplayValue}' where the chart needs a number.");
                values.Add(number);
            }

            series.Add(new ChartSeries { Name = Text(range[0][column]) ?? string.Empty, Values = values });
        }

        for (var row = 1; row < range.Length; row++)
            categories.Add(Text(range[row][0]) ?? string.Empty);

        return (new SetChartDataOp { Target = found.Target, Categories = categories, Series = series },
            $"{categories.Count} categories, {series.Count} series", null, null);
    }

    /// <summary>A cell as displayed, empty for a blank cell, or null for a formula with no stored result.</summary>
    private static string? Text(CellInfo? cell) =>
        cell is null ? string.Empty
        : cell.Formula is not null && cell.FormattedValue is null ? null
        : cell.FormattedValue ?? cell.DisplayValue ?? string.Empty;

    private static string NotCalculated(string where) =>
        $"{where} holds a formula with no calculated result stored in the workbook. Open and save the workbook " +
        "in Excel, which stores every result, then refresh again.";

    private static DataRefreshBinding Failed(FoundBinding found, string? resolved, string code, string message) => new()
    {
        Kind = found.Kind,
        Marker = found.Marker,
        Reference = found.Reference,
        Target = found.Target,
        Range = resolved,
        ErrorCode = code,
        Message = message
    };

    /// <summary>Resolves references against one workbook and reads the cells they name.</summary>
    private sealed class RangeReader
    {
        private readonly OfficeAgentClient _client;
        private readonly byte[] _workbook;
        private readonly InspectResult _book;

        public RangeReader(OfficeAgentClient client, byte[] workbook, InspectResult book)
        {
            _client = client;
            _workbook = workbook;
            _book = book;
        }

        /// <summary>The cells of a range, row by row with nulls for empty cells, or null with the reason.</summary>
        public CellInfo?[][]? Read(string reference, out string? resolved, out string? code, out string? message)
        {
            resolved = null;
            code = DataRefreshErrorCodes.RangeNotFound;
            if (!WorkbookReference.TryParse(reference, out var parsed))
            {
                message = $"'{reference}' is not a range or name. Write a range as Sheet!A1 or 'Sheet name'!A1:D6.";
                return null;
            }

            WorksheetInfo? sheet;
            string? range;
            if (parsed.Sheet is not null)
            {
                sheet = _book.Worksheets.FirstOrDefault(candidate =>
                    string.Equals(candidate.Name, parsed.Sheet, StringComparison.OrdinalIgnoreCase));
                range = parsed.Range;
                if (sheet is null)
                {
                    message = $"The workbook has no sheet named '{parsed.Sheet}'. Its sheets are " +
                              $"{string.Join(", ", _book.Worksheets.Select(w => $"'{w.Name}'"))}.";
                    return null;
                }
            }
            else if (_book.DefinedNames.FirstOrDefault(name =>
                         string.Equals(name.Name, parsed.Name, StringComparison.OrdinalIgnoreCase)) is { } defined)
            {
                sheet = _book.Worksheets.FirstOrDefault(candidate => candidate.SheetId == defined.SheetId);
                range = defined.Range;
            }
            else
            {
                var table = _book.Worksheets
                    .SelectMany(candidate => candidate.Tables.Select(t => (Sheet: candidate, Table: t)))
                    .FirstOrDefault(entry =>
                        string.Equals(entry.Table.DisplayName, parsed.Name, StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(entry.Table.Name, parsed.Name, StringComparison.OrdinalIgnoreCase));
                sheet = table.Sheet;
                range = table.Table?.Reference;
            }

            if (sheet is null || range is null ||
                !WorkbookReference.TryBounds(range, out var left, out var top, out var right, out var bottom))
            {
                message = $"The workbook has no defined name or table called '{parsed.Name}'.";
                return null;
            }

            resolved = $"{Quote(sheet.Name)}!{range}";
            var rows = (int)(bottom - top + 1);
            var columns = right - left + 1;
            if ((long)rows * columns > MaximumRefreshCells)
            {
                code = DataRefreshErrorCodes.RangeTooLarge;
                message = $"{resolved} is {rows}×{columns} cells; one binding reads at most {MaximumRefreshCells}.";
                return null;
            }

            var cells = _client.Inspect(_workbook, new InspectOptions
            {
                Fidelity = Fidelity.Content,
                SheetId = sheet.SheetId,
                Range = range,
                MaximumCells = rows * columns
            }).Cells;

            var grid = new CellInfo?[rows][];
            for (var row = 0; row < rows; row++) grid[row] = new CellInfo?[columns];
            foreach (var cell in cells)
                if (SpreadsheetPartUtility.TryParseCell(cell.Anchor.Address, out var column, out var rowNumber))
                    grid[rowNumber - top][column - left] = cell;

            code = null;
            message = null;
            return grid;
        }

        private static string Quote(string sheet) =>
            sheet.All(c => char.IsLetterOrDigit(c) || c == '_') ? sheet : "'" + sheet.Replace("'", "''") + "'";
    }
}
