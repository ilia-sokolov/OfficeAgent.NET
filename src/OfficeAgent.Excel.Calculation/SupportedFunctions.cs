namespace OfficeAgent.Excel.Calculation;

/// <summary>
/// The functions <see cref="ClosedXmlFormulaEvaluator"/> calculates. Each one was checked
/// against an independently computed result (see the spike report on issue #48), and each has
/// a case in the calculation tests; a function joins the list only with a test.
/// </summary>
/// <remarks>
/// Left out on purpose: functions ClosedXML 0.105 does not implement (<c>XLOOKUP</c>,
/// <c>IFS</c>, <c>SWITCH</c>, <c>AVERAGEIFS</c>, <c>SMALL</c>, the dynamic-array functions,
/// <c>LET</c>, <c>LAMBDA</c>, most financial functions), and volatile functions
/// (<c>NOW</c>, <c>TODAY</c>, <c>RAND</c>, <c>RANDBETWEEN</c>, <c>OFFSET</c>,
/// <c>INDIRECT</c>), whose cached value would be stale as soon as it was written.
/// </remarks>
internal static class SupportedFunctions
{
    public static readonly IReadOnlyList<string> All = new[]
    {
        // Math and aggregation.
        "ABS", "AVERAGE", "CEILING", "CEILING.MATH", "COMBIN", "COUNT", "COUNTA", "COUNTBLANK", "COUNTIF",
        "COUNTIFS", "EVEN", "EXP", "FACT", "FLOOR", "FLOOR.MATH", "GCD", "INT", "LCM", "LN", "LOG", "LOG10",
        "MAX", "MIN", "MOD", "MROUND", "ODD", "PI", "POWER", "PRODUCT", "QUOTIENT", "ROUND", "ROUNDDOWN",
        "ROUNDUP", "SIGN", "SQRT", "SUBTOTAL", "SUM", "SUMIF", "SUMIFS", "SUMPRODUCT", "TRUNC",
        // Statistics.
        "LARGE", "MEDIAN", "STDEV.P", "STDEV.S", "VAR.S",
        // Logic and information.
        "AND", "IF", "IFERROR", "ISBLANK", "ISERROR", "ISEVEN", "ISNA", "ISNUMBER", "ISODD", "ISTEXT", "N",
        "NA", "NOT", "OR", "TYPE",
        // Text.
        "CHAR", "CODE", "CONCAT", "CONCATENATE", "EXACT", "FIND", "FIXED", "LEFT", "LEN", "LOWER", "MID",
        "NUMBERVALUE", "PROPER", "REPLACE", "REPT", "RIGHT", "SEARCH", "SUBSTITUTE", "T", "TEXT", "TEXTJOIN",
        "TRIM", "UPPER", "VALUE",
        // Lookup and reference.
        "COLUMN", "COLUMNS", "HLOOKUP", "INDEX", "MATCH", "ROW", "ROWS", "VLOOKUP",
        // Date and time.
        "DATE", "DAY", "DAYS", "EDATE", "EOMONTH", "HOUR", "MINUTE", "MONTH", "NETWORKDAYS", "TIME",
        "WEEKDAY", "WORKDAY", "YEAR", "YEARFRAC",
        // Finance.
        "FV", "PMT"
    };

    /// <summary>Functions whose result changes on every recalculation.</summary>
    public static readonly IReadOnlyList<string> Volatile = new[]
    {
        "NOW", "TODAY", "RAND", "RANDBETWEEN", "RANDARRAY", "OFFSET", "INDIRECT", "INFO", "CELL"
    };

    /// <summary>
    /// The allowlisted functions a workbook must store with the <c>_xlfn.</c> prefix. Excel reads
    /// one stored without it as an unknown name and shows <c>#NAME?</c>.
    /// </summary>
    public static readonly IReadOnlyList<string> Future = new[]
    {
        "CEILING.MATH", "CONCAT", "DAYS", "FLOOR.MATH", "NUMBERVALUE", "STDEV.P", "STDEV.S", "TEXTJOIN", "VAR.S"
    };

    private static readonly HashSet<string> Set = new(All, StringComparer.OrdinalIgnoreCase);
    private static readonly HashSet<string> FutureSet = new(Future, StringComparer.OrdinalIgnoreCase);
    private static readonly HashSet<string> VolatileSet = new(Volatile, StringComparer.OrdinalIgnoreCase);

    public static bool Contains(string name) => Set.Contains(name);

    public static bool IsVolatile(string name) => VolatileSet.Contains(name);

    public static bool IsFuture(string name) => FutureSet.Contains(name);
}
