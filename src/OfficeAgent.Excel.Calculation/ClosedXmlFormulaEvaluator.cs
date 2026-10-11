using System.Globalization;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using OfficeAgent.Abstractions;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeAgent.Excel.Calculation;

/// <summary>
/// Calculates workbook formulas with ClosedXML, for the functions it is verified to calculate
/// correctly, and reports every other formula as not calculated.
/// </summary>
/// <remarks>
/// <para>
/// A formula is calculated only when every function it calls is on the
/// <see cref="SupportedFunctions"/> list and it has no structured or external reference, is
/// not an array formula, and depends on no cell that was itself not calculated. ClosedXML
/// evaluates structured references to 0 without an error and refuses a whole workbook that
/// contains a circular reference, so neither may produce a value here.
/// </para>
/// <para>
/// The workbook is loaded into ClosedXML separately; nothing ClosedXML would write is ever
/// saved. Results go back into the caller's package as cached values only. Text is formatted
/// and parsed with the invariant culture, whatever the server's locale.
/// </para>
/// </remarks>
public sealed class ClosedXmlFormulaEvaluator : IFormulaEvaluator
{
    /// <inheritdoc />
    public string Name { get; } = "ClosedXML " +
        (typeof(XLWorkbook).Assembly.GetName().Version?.ToString(3) ?? "unknown");

    /// <inheritdoc />
    public IReadOnlyList<string> SupportedFunctions => Calculation.SupportedFunctions.All;

    /// <inheritdoc />
    public FormulaEvaluation Evaluate(byte[] workbook, CancellationToken cancellationToken = default)
    {
        if (workbook is null) throw new ArgumentNullException(nameof(workbook));

        // ClosedXML formats and parses text (TEXT, FIXED, NUMBERVALUE, VALUE) with the thread's
        // culture. A cached value must not depend on the locale of the server that wrote it.
        var culture = CultureInfo.CurrentCulture;
        CultureInfo.CurrentCulture = CultureInfo.InvariantCulture;
        try
        {
            return EvaluateInvariant(workbook, cancellationToken);
        }
        finally
        {
            CultureInfo.CurrentCulture = culture;
        }
    }

    private FormulaEvaluation EvaluateInvariant(byte[] workbook, CancellationToken cancellationToken)
    {

        var sheetIds = SheetIds(workbook);
        XLWorkbook book;
        try
        {
            book = new XLWorkbook(new MemoryStream(workbook, writable: false));
        }
        catch (Exception e) when (e is not OperationCanceledException)
        {
            return new FormulaEvaluation { NotCalculatedReason = $"The workbook could not be loaded for calculation: {e.Message}" };
        }

        using (book)
        {
            var cells = new List<Candidate>();
            foreach (var sheet in book.Worksheets)
            {
                if (!sheetIds.TryGetValue(sheet.Name, out var sheetId)) continue;
                foreach (var cell in sheet.CellsUsed(XLCellsUsedOptions.All, c => c.HasFormula))
                    cells.Add(new Candidate(sheet.Name, sheetId, cell, Classify(book, sheet, cell)));
            }
            cancellationToken.ThrowIfCancellationRequested();

            PropagateUncalculated(cells);

            try
            {
                book.RecalculateAllFormulas();
            }
            catch (InvalidOperationException e) when (e.Message.IndexOf("cycle", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                const string circular = "The workbook contains a circular reference, so no formula in it was calculated.";
                return new FormulaEvaluation
                {
                    Cells = cells.Select(c => c.NotCalculated(circular)).ToList(),
                    NotCalculatedReason = circular,
                    CircularReference = true
                };
            }
            catch (Exception e) when (e is not OperationCanceledException)
            {
                var reason = $"The workbook could not be calculated: {e.Message}";
                return new FormulaEvaluation { Cells = cells.Select(c => c.NotCalculated(reason)).ToList(), NotCalculatedReason = reason };
            }
            cancellationToken.ThrowIfCancellationRequested();

            return new FormulaEvaluation { Cells = cells.Select(Result).ToList() };
        }
    }

    /// <summary>Why the formula cannot be calculated on its own terms, or null when it can.</summary>
    private static string? Classify(XLWorkbook book, IXLWorksheet sheet, IXLCell cell)
    {
        if (cell.HasArrayFormula)
            return "It is an array formula, which this engine does not calculate.";

        var analysis = FormulaAnalysis.Of(cell.FormulaA1);
        if (analysis.HasStructuredOrExternalReference)
            return "It uses a structured table reference or a link to another workbook, which this engine does not calculate.";

        foreach (var function in analysis.Functions)
        {
            if (Calculation.SupportedFunctions.IsVolatile(function))
                return $"It uses {function}, a volatile function whose value changes on every recalculation.";
            if (!Calculation.SupportedFunctions.Contains(function))
                return $"It uses {function}, which this engine does not calculate.";
        }

        foreach (var function in analysis.Unprefixed)
        {
            if (Calculation.SupportedFunctions.IsFuture(function))
                return $"It calls {function} without the _xlfn. prefix a workbook stores it with, so Excel reads it as an unknown name (#NAME?).";
        }

        foreach (var name in analysis.Names)
            if (!IsKnownName(book, sheet, name))
                return $"It uses the name {name}, which the workbook does not define.";
        return null;
    }

    private static bool IsKnownName(XLWorkbook book, IXLWorksheet sheet, string name) =>
        sheet.DefinedNames.Contains(name) || book.DefinedNames.Contains(name);

    /// <summary>
    /// Marks every formula that refers to a cell which will not be calculated, repeating until
    /// nothing changes, so a value is never computed from an uncalculated input.
    /// </summary>
    private static void PropagateUncalculated(List<Candidate> cells)
    {
        var uncalculated = cells.Where(c => c.Reason is not null).ToList();
        while (uncalculated.Count > 0)
        {
            var newly = new List<Candidate>();
            foreach (var candidate in cells.Where(c => c.Reason is null))
            {
                var references = FormulaAnalysis.Of(candidate.Cell.FormulaA1).References;
                var blocker = uncalculated.FirstOrDefault(u => references.Any(r =>
                    string.Equals(r.Sheet ?? candidate.Sheet, u.Sheet, StringComparison.OrdinalIgnoreCase) &&
                    r.Contains(u.Column, u.Row)));
                if (blocker is null) continue;
                candidate.Reason = $"It depends on {blocker.Sheet}!{blocker.Address}, which was not calculated.";
                newly.Add(candidate);
            }
            uncalculated = newly;
        }
    }

    private static FormulaCellResult Result(Candidate candidate)
    {
        if (candidate.Reason is not null) return candidate.NotCalculated(candidate.Reason);

        var value = candidate.Cell.Value;
        string text;
        FormulaValueKind kind;
        if (value.IsNumber || value.IsDateTime || value.IsTimeSpan)
        {
            var number = value.IsNumber ? value.GetNumber()
                : value.IsDateTime ? value.GetDateTime().ToOADate()
                : value.GetTimeSpan().TotalDays;
            if (double.IsNaN(number) || double.IsInfinity(number)) { kind = FormulaValueKind.Error; text = "#NUM!"; }
            else { kind = FormulaValueKind.Number; text = number.ToString("R", CultureInfo.InvariantCulture); }
        }
        else if (value.IsBoolean) { kind = FormulaValueKind.Boolean; text = value.GetBoolean() ? "TRUE" : "FALSE"; }
        else if (value.IsText) { kind = FormulaValueKind.Text; text = value.GetText(); }
        else if (value.IsError) { kind = FormulaValueKind.Error; text = ErrorText(value.GetError()); }
        else return candidate.NotCalculated("The formula produced no value this engine can record.");

        return new FormulaCellResult
        {
            SheetId = candidate.SheetId,
            Address = candidate.Address,
            Formula = candidate.Cell.FormulaA1,
            Outcome = FormulaOutcome.Calculated,
            ValueKind = kind,
            Value = text
        };
    }

    private static string ErrorText(XLError error) => error switch
    {
        XLError.DivisionByZero => "#DIV/0!",
        XLError.IncompatibleValue => "#VALUE!",
        XLError.NameNotRecognized => "#NAME?",
        XLError.NoValueAvailable => "#N/A",
        XLError.CellReference => "#REF!",
        XLError.NumberInvalid => "#NUM!",
        XLError.NullValue => "#NULL!",
        _ => "#VALUE!"
    };

    private static Dictionary<string, uint> SheetIds(byte[] workbook)
    {
        using var document = SpreadsheetDocument.Open(new MemoryStream(workbook, writable: false), false);
        return document.WorkbookPart?.Workbook?.Sheets?.Elements<S.Sheet>()
                   .Where(s => s.Name?.Value is not null && s.SheetId?.Value is not null)
                   .ToDictionary(s => s.Name!.Value!, s => s.SheetId!.Value, StringComparer.OrdinalIgnoreCase)
               ?? new Dictionary<string, uint>(StringComparer.OrdinalIgnoreCase);
    }

    private sealed class Candidate
    {
        public Candidate(string sheet, uint sheetId, IXLCell cell, string? reason)
        {
            Sheet = sheet;
            SheetId = sheetId;
            Cell = cell;
            Reason = reason;
            Address = cell.Address.ToString()!;
            Column = cell.Address.ColumnNumber;
            Row = cell.Address.RowNumber;
        }

        public string Sheet { get; }
        public uint SheetId { get; }
        public IXLCell Cell { get; }
        public string Address { get; }
        public int Column { get; }
        public int Row { get; }
        public string? Reason { get; set; }

        public FormulaCellResult NotCalculated(string reason) => new()
        {
            SheetId = SheetId,
            Address = Address,
            Formula = Cell.FormulaA1,
            Outcome = FormulaOutcome.NotCalculated,
            Reason = reason
        };
    }
}
