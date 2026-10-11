using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeAgent.Abstractions;

namespace OfficeAgent.Core;

/// <summary>
/// Shared low-level spreadsheet helpers used by Excel editing and PowerPoint embedded
/// chart workbooks. This is infrastructure rather than a formula calculation engine.
/// </summary>
internal static class SpreadsheetPartUtility
{
    private static readonly Regex CellReference = new(
        @"^\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)$",
        RegexOptions.CultureInvariant,
        TimeSpan.FromMilliseconds(100));

    /// <summary>Creates the minimal editable workbook embedded behind a native chart.</summary>
    public static byte[] CreateChartWorkbook(
        IReadOnlyList<string> categories,
        IReadOnlyList<ChartSeries> series)
    {
        using var stream = new MemoryStream();
        using (var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            var data = new SheetData();
            worksheetPart.Worksheet = new Worksheet(data);

            var header = new Row { RowIndex = 1U };
            header.Append(InlineCell("A1", "Category"));
            for (var i = 0; i < series.Count; i++)
                header.Append(InlineCell(ColumnName(i + 2) + "1", series[i].Name));
            data.Append(header);

            for (var rowIndex = 0; rowIndex < categories.Count; rowIndex++)
            {
                var rowNumber = (uint)(rowIndex + 2);
                var row = new Row { RowIndex = rowNumber };
                row.Append(InlineCell("A" + rowNumber, categories[rowIndex]));
                for (var seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
                {
                    var value = series[seriesIndex].Values[rowIndex];
                    var cell = new Cell { CellReference = ColumnName(seriesIndex + 2) + rowNumber };
                    if (value is not null)
                        cell.CellValue = new CellValue(value.Value.ToString("R", CultureInfo.InvariantCulture));
                    row.Append(cell);
                }
                data.Append(row);
            }

            var sheets = workbookPart.Workbook.AppendChild(new Sheets());
            sheets.Append(new Sheet
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1U,
                Name = "ChartData"
            });
            workbookPart.Workbook.CalculationProperties = new CalculationProperties
            {
                CalculationMode = CalculateModeValues.Auto,
                FullCalculationOnLoad = true,
                ForceFullCalculation = true
            };
            worksheetPart.Worksheet.Save();
            workbookPart.Workbook.Save();
        }
        return stream.ToArray();
    }

    /// <summary>Resolves a worksheet by workbook sheet id.</summary>
    public static (Sheet Sheet, WorksheetPart Part)? ResolveSheet(SpreadsheetDocument document, uint sheetId)
    {
        var workbookPart = document.WorkbookPart;
        var sheet = workbookPart?.WorkbookRoot().Sheets?.Elements<Sheet>()
            .FirstOrDefault(candidate => candidate.SheetId?.Value == sheetId);
        if (sheet?.Id?.Value is not { Length: > 0 } relationshipId) return null;
        return workbookPart!.GetPartById(relationshipId) is WorksheetPart part
            ? (sheet, part)
            : null;
    }

    /// <summary>Returns or creates a cell while preserving existing row and cell ordering.</summary>
    public static Cell? Cell(WorksheetPart part, string address, bool create)
    {
        if (!TryParseCell(address, out var column, out var rowIndex)) return null;
        var normalized = ColumnName(column) + rowIndex.ToString(CultureInfo.InvariantCulture);
        var sheetData = part.WorksheetRoot().GetFirstChild<SheetData>();
        if (sheetData is null)
        {
            if (!create) return null;
            sheetData = part.WorksheetRoot().AppendChild(new SheetData());
        }

        var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == rowIndex);
        if (row is null)
        {
            if (!create) return null;
            row = new Row { RowIndex = rowIndex };
            var next = sheetData.Elements<Row>().FirstOrDefault(r => (r.RowIndex?.Value ?? 0) > rowIndex);
            if (next is null) sheetData.Append(row); else sheetData.InsertBefore(row, next);
        }

        var existing = row.Elements<Cell>().FirstOrDefault(c =>
            string.Equals(c.CellReference?.Value, normalized, StringComparison.OrdinalIgnoreCase));
        if (existing is not null || !create) return existing;

        var cell = new Cell { CellReference = normalized };
        var nextCell = row.Elements<Cell>().FirstOrDefault(c =>
            TryParseCell(c.CellReference?.Value ?? string.Empty, out var nextColumn, out _) && nextColumn > column);
        if (nextCell is null) row.Append(cell); else row.InsertBefore(cell, nextCell);
        return cell;
    }

    /// <summary>Resolves inline and shared strings; formulas use their cached value.</summary>
    public static string DisplayValue(SpreadsheetDocument document, Cell cell)
    {
        if (cell.DataType?.Value == CellValues.InlineString)
            return string.Concat(cell.InlineString?.Descendants<Text>().Select(t => t.Text) ?? Enumerable.Empty<string>());

        var raw = cell.CellValue?.InnerText ?? string.Empty;
        if (cell.DataType?.Value == CellValues.SharedString &&
            int.TryParse(raw, NumberStyles.None, CultureInfo.InvariantCulture, out var index))
        {
            var item = document.WorkbookPart?.SharedStringTablePart?.SharedStringTable?
                .Elements<SharedStringItem>().ElementAtOrDefault(index);
            return item is null ? raw : string.Concat(item.Descendants<Text>().Select(t => t.Text));
        }

        if (cell.DataType?.Value == CellValues.Boolean)
            return raw == "1" ? "TRUE" : raw == "0" ? "FALSE" : raw;
        return raw;
    }

    /// <summary>Parses a single A1 reference.</summary>
    public static bool TryParseCell(string address, out int column, out uint row)
    {
        column = 0;
        row = 0;
        var match = CellReference.Match(address ?? string.Empty);
        if (!match.Success || !uint.TryParse(match.Groups[2].Value, out row)) return false;
        foreach (var character in match.Groups[1].Value.ToUpperInvariant())
            column = checked(column * 26 + character - 'A' + 1);
        return column is >= 1 and <= 16384 && row <= 1048576;
    }

    /// <summary>Parses a single cell or rectangular A1 range.</summary>
    public static bool TryParseRange(string address, out int left, out uint top, out int right, out uint bottom)
    {
        left = right = 0;
        top = bottom = 0;
        var parts = (address ?? string.Empty).Split(':');
        if (parts.Length is < 1 or > 2 || !TryParseCell(parts[0], out left, out top)) return false;
        if (parts.Length == 1) { right = left; bottom = top; return true; }
        if (!TryParseCell(parts[1], out right, out bottom)) return false;
        if (left > right) (left, right) = (right, left);
        if (top > bottom) (top, bottom) = (bottom, top);
        return true;
    }

    /// <summary>Converts a one-based worksheet column number to letters.</summary>
    public static string ColumnName(int column)
    {
        if (column is < 1 or > 16384) throw new ArgumentOutOfRangeException(nameof(column));
        var result = string.Empty;
        while (column > 0)
        {
            column--;
            result = (char)('A' + column % 26) + result;
            column /= 26;
        }
        return result;
    }

    /// <summary>Marks formulas stale so Excel recalculates them when the workbook opens.</summary>
    public static void RecalculateOnOpen(SpreadsheetDocument document)
    {
        var workbook = document.WorkbookPart?.Workbook
            ?? throw new InvalidOperationException("Workbook has no workbook part.");
        workbook.CalculationProperties ??= new CalculationProperties();
        workbook.CalculationProperties.CalculationMode = CalculateModeValues.Auto;
        workbook.CalculationProperties.FullCalculationOnLoad = true;
        workbook.CalculationProperties.ForceFullCalculation = true;
        if (document.WorkbookPart?.CalculationChainPart is { } chain)
            document.WorkbookPart.DeletePart(chain);
    }

    /// <summary>
    /// The Excel error values the workbook's formula cells already carry as cached values,
    /// keyed by sheet id and address - the errors a plan did not introduce.
    /// </summary>
    public static Dictionary<(uint SheetId, string Address), string> CachedFormulaErrors(SpreadsheetDocument document)
    {
        var errors = new Dictionary<(uint, string), string>();
        foreach (var (sheetId, part) in Worksheets(document))
            foreach (var cell in part.Worksheet!.Descendants<Cell>())
                if (cell.CellFormula is not null && cell.DataType?.Value == CellValues.Error &&
                    cell.CellReference?.Value is { } address && cell.CellValue?.Text is { } value)
                    errors[(sheetId, address)] = value;
        return errors;
    }

    /// <summary>
    /// Writes calculated results into the formula cells' cached values (<c>&lt;v&gt;</c> and
    /// the type attribute) and nothing else. A cell whose cached value already matches is left
    /// untouched, and a worksheet is saved only when one of its cells changed. Cells that were
    /// not calculated keep whatever they had.
    /// </summary>
    /// <returns>The number of cells whose cached value was written.</returns>
    public static int WriteCachedValues(SpreadsheetDocument document, IEnumerable<FormulaCellResult> results)
    {
        var calculated = results
            .Where(r => r.Outcome == FormulaOutcome.Calculated && r.ValueKind is not null && r.Value is not null)
            .GroupBy(r => r.SheetId)
            .ToDictionary(g => g.Key, g => g.ToDictionary(r => r.Address, StringComparer.OrdinalIgnoreCase));
        var written = 0;
        foreach (var (sheetId, part) in Worksheets(document))
        {
            if (!calculated.TryGetValue(sheetId, out var cells)) continue;
            var changed = false;
            foreach (var cell in part.Worksheet!.Descendants<Cell>())
            {
                if (cell.CellFormula is null || cell.CellReference?.Value is not { } address ||
                    !cells.TryGetValue(address, out var result)) continue;
                var type = result.ValueKind switch
                {
                    FormulaValueKind.Text => CellValues.String,
                    FormulaValueKind.Boolean => CellValues.Boolean,
                    FormulaValueKind.Error => CellValues.Error,
                    _ => (CellValues?)null
                };
                var value = result.ValueKind == FormulaValueKind.Boolean
                    ? (string.Equals(result.Value, "TRUE", StringComparison.OrdinalIgnoreCase) ? "1" : "0")
                    : result.Value!;
                if (cell.CellValue?.Text == value && cell.DataType?.Value == type) continue;
                cell.DataType = type is null ? null : new EnumValue<CellValues>(type.Value);
                cell.CellValue = new CellValue(value);
                changed = true;
                written++;
            }
            if (changed) part.Worksheet.Save();
        }
        return written;
    }

    private static IEnumerable<(uint SheetId, WorksheetPart Part)> Worksheets(SpreadsheetDocument document)
    {
        var workbookPart = document.WorkbookPart;
        if (workbookPart?.Workbook?.Sheets is not { } sheets) yield break;
        foreach (var sheet in sheets.Elements<Sheet>())
            if (sheet.SheetId?.Value is { } id && sheet.Id?.Value is { Length: > 0 } relationship &&
                workbookPart.GetPartById(relationship) is WorksheetPart part && part.Worksheet is not null)
                yield return (id, part);
    }

    /// <summary>Writes a typed scalar value without changing the cell's style.</summary>
    public static void WriteValue(Cell cell, string value, SpreadsheetCellValueKind kind = SpreadsheetCellValueKind.Auto)
    {
        cell.CellFormula = null;
        cell.CellValue = null;
        cell.InlineString = null;
        cell.DataType = null;

        var boolean = bool.TryParse(value, out var booleanValue);
        var number = double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var numericValue) &&
            !double.IsNaN(numericValue) && !double.IsInfinity(numericValue);
        if (kind == SpreadsheetCellValueKind.Boolean && !boolean)
            throw new ArgumentException($"'{value}' is not a Boolean value.", nameof(value));
        if (kind == SpreadsheetCellValueKind.Number && !number)
            throw new ArgumentException($"'{value}' is not an invariant numeric value.", nameof(value));

        if (kind == SpreadsheetCellValueKind.Boolean || (kind == SpreadsheetCellValueKind.Auto && boolean))
        {
            cell.DataType = CellValues.Boolean;
            cell.CellValue = new CellValue(booleanValue ? "1" : "0");
        }
        else if (kind == SpreadsheetCellValueKind.Number || (kind == SpreadsheetCellValueKind.Auto && number))
        {
            cell.CellValue = new CellValue(numericValue.ToString("R", CultureInfo.InvariantCulture));
        }
        else
        {
            cell.DataType = CellValues.InlineString;
            cell.InlineString = new InlineString(new Text(value));
        }
    }

    private static Cell InlineCell(string reference, string value) => new()
    {
        CellReference = reference,
        DataType = CellValues.InlineString,
        InlineString = new InlineString(new Text(value ?? string.Empty))
    };
}
