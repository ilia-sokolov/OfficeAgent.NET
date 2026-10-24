using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using ExcelNumberFormat;
using OfficeAgent.Core;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeAgent.Excel;

/// <summary>
/// Writes a cell the way Excel displays it: its stored value with the cell's number format
/// applied, so <c>4200000</c> under <c>$#,##0</c> reads <c>$4,200,000</c>.
/// </summary>
/// <remarks>
/// The workbook stores no display text, only the value and a format code, so this is the only
/// way to put the number a reader sees into a report. Formatting uses the invariant culture's
/// separators: a workbook does not record the locale it was formatted in, and a report that
/// changed its decimal point with the server's culture would be worse than one that never does.
/// A format code the formatter cannot read leaves the value as stored.
/// </remarks>
internal sealed class CellFormatter
{
    /// <summary>
    /// The formats Excel assigns by id without writing their code into the workbook (ECMA-376
    /// 18.8.30). Ids 14 and 22 are locale-dependent in Excel; these are its en-US forms.
    /// </summary>
    private static readonly IReadOnlyDictionary<uint, string> BuiltIn = new Dictionary<uint, string>
    {
        [0] = "General", [1] = "0", [2] = "0.00", [3] = "#,##0", [4] = "#,##0.00",
        [5] = "\"$\"#,##0_);(\"$\"#,##0)", [6] = "\"$\"#,##0_);[Red](\"$\"#,##0)",
        [7] = "\"$\"#,##0.00_);(\"$\"#,##0.00)", [8] = "\"$\"#,##0.00_);[Red](\"$\"#,##0.00)",
        [9] = "0%", [10] = "0.00%", [11] = "0.00E+00", [12] = "# ?/?", [13] = "# ??/??",
        [14] = "m/d/yyyy", [15] = "d-mmm-yy", [16] = "d-mmm", [17] = "mmm-yy",
        [18] = "h:mm AM/PM", [19] = "h:mm:ss AM/PM", [20] = "h:mm", [21] = "h:mm:ss", [22] = "m/d/yyyy h:mm",
        // 41-44 are Excel's Accounting formats, the usual choice in a finance model.
        [41] = "_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)", [42] = "_(\"$\"* #,##0_);_(\"$\"* (#,##0);_(\"$\"* \"-\"_);_(@_)",
        [43] = "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)", [44] = "_(\"$\"* #,##0.00_);_(\"$\"* (#,##0.00);_(\"$\"* \"-\"??_);_(@_)",
        [37] = "#,##0 ;(#,##0)", [38] = "#,##0 ;[Red](#,##0)", [39] = "#,##0.00;(#,##0.00)", [40] = "#,##0.00;[Red](#,##0.00)",
        [45] = "mm:ss", [46] = "[h]:mm:ss", [47] = "mmss.0", [48] = "##0.0E+0", [49] = "@"
    };

    private readonly SpreadsheetDocument _document;
    private readonly List<S.CellFormat> _cellFormats;
    private readonly Dictionary<uint, string> _customFormats;
    private readonly Dictionary<string, NumberFormat?> _parsed = new(StringComparer.Ordinal);
    private readonly bool _date1904;

    public CellFormatter(SpreadsheetDocument document)
    {
        _document = document;
        var stylesheet = document.WorkbookPart?.WorkbookStylesPart?.Stylesheet;
        _cellFormats = stylesheet?.CellFormats?.Elements<S.CellFormat>().ToList() ?? new List<S.CellFormat>();
        _customFormats = stylesheet?.NumberingFormats?.Elements<S.NumberingFormat>()
            .Where(format => format.NumberFormatId?.Value is not null && format.FormatCode?.Value is not null)
            .GroupBy(format => format.NumberFormatId!.Value)
            .ToDictionary(group => group.Key, group => group.First().FormatCode!.Value!)
            ?? new Dictionary<uint, string>();
        _date1904 = document.WorkbookPart?.Workbook?.WorkbookProperties?.Date1904?.Value == true;
    }

    /// <summary>
    /// The cell as Excel displays it, or <see langword="null"/> for a formula whose result the
    /// workbook does not store.
    /// </summary>
    public string? Format(S.Cell cell)
    {
        var raw = cell.CellValue?.InnerText;
        if (raw is null)
            return cell.CellFormula is not null && cell.DataType?.Value != S.CellValues.InlineString
                ? null
                : SpreadsheetPartUtility.DisplayValue(_document, cell);

        var type = cell.DataType?.Value;
        if (type is not null && type != S.CellValues.Number)
            return SpreadsheetPartUtility.DisplayValue(_document, cell);

        if (!double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
            return raw;

        var format = NumberFormatFor(cell);
        if (format is null) return raw;

        try
        {
            // Alignment padding - the space "_)" reserves and the gap a "*" fill stretches -
            // lines numbers up in a column and means nothing in running text.
            return Collapse(format.Format(number, CultureInfo.InvariantCulture, _date1904));
        }
        catch (Exception exception) when (exception is ArgumentException or FormatException or OverflowException or InvalidOperationException)
        {
            return raw;
        }
    }

    /// <summary>Trims the padding and removes the fill gap between a currency sign and its number.</summary>
    private static string Collapse(string text) => FillGap.Replace(text.Trim(), "$1");

    /// <summary>Whitespace between a leading symbol such as <c>$</c> and the number, sign or bracket after it.</summary>
    private static readonly Regex FillGap = new(@"^([^\d\s(\-]+)\s+(?=[\d(\-])", RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));

    private NumberFormat? NumberFormatFor(S.Cell cell)
    {
        var styleIndex = (int)(cell.StyleIndex?.Value ?? 0U);
        var formatId = styleIndex < _cellFormats.Count ? _cellFormats[styleIndex].NumberFormatId?.Value ?? 0U : 0U;
        var code = _customFormats.TryGetValue(formatId, out var custom)
            ? custom
            : BuiltIn.TryGetValue(formatId, out var builtIn) ? builtIn : "General";

        if (!_parsed.TryGetValue(code, out var parsed))
        {
            var candidate = new NumberFormat(code);
            parsed = candidate.IsValid ? candidate : null;
            _parsed[code] = parsed;
        }

        return parsed;
    }
}
