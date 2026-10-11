using System.Text;
using System.Text.RegularExpressions;

namespace OfficeAgent.Excel;

/// <summary>
/// Functions added to Excel after the 2007 file format, which a workbook stores with a prefix:
/// <c>_xlfn.STDEV.S(A1:A5)</c>, and <c>_xlfn._xlws.FILTER(…)</c> for FILTER and SORT. Excel
/// displays them without it, but reads an unprefixed name in the file as unknown and shows
/// <c>#NAME?</c>. The list is what Excel 16 itself writes for each function (see the native
/// compatibility record).
/// </summary>
internal static class FutureFunctions
{
    private const string Prefix = "_xlfn.";
    private const string WorksheetPrefix = "_xlfn._xlws.";

    private static readonly HashSet<string> Worksheet = new(StringComparer.OrdinalIgnoreCase) { "FILTER", "SORT" };

    private static readonly HashSet<string> Prefixed = new(StringComparer.OrdinalIgnoreCase)
    {
        "ACOT", "ACOTH", "AGGREGATE", "ARABIC", "ARRAYTOTEXT", "BASE", "BETA.DIST", "BETA.INV", "BINOM.DIST",
        "BINOM.DIST.RANGE", "BINOM.INV", "BITAND", "BITLSHIFT", "BITOR", "BITRSHIFT", "BITXOR", "BYCOL", "BYROW",
        "CEILING.MATH", "CEILING.PRECISE", "CHISQ.DIST", "CHISQ.DIST.RT", "CHISQ.INV", "CHISQ.INV.RT", "CHISQ.TEST",
        "CHOOSECOLS", "CHOOSEROWS", "COMBINA", "CONCAT", "CONFIDENCE.NORM", "CONFIDENCE.T", "COT", "COTH",
        "COVARIANCE.P", "COVARIANCE.S", "CSC", "CSCH", "DAYS", "DECIMAL", "DROP", "ENCODEURL", "ERF.PRECISE",
        "ERFC.PRECISE", "EXPAND", "EXPON.DIST", "F.DIST", "F.DIST.RT", "F.INV", "F.INV.RT", "F.TEST", "FILTERXML",
        "FLOOR.MATH", "FLOOR.PRECISE", "FORECAST.ETS", "FORECAST.ETS.CONFINT", "FORECAST.ETS.SEASONALITY",
        "FORECAST.ETS.STAT", "FORECAST.LINEAR", "FORMULATEXT", "GAMMA", "GAMMA.DIST", "GAMMA.INV", "GAMMALN.PRECISE",
        "GAUSS", "GROUPBY", "HSTACK", "HYPGEOM.DIST", "IFNA", "IFS", "IMAGE", "IMCOSH", "IMCOT", "IMCSC", "IMCSCH",
        "IMSEC", "IMSECH", "IMSINH", "IMTAN", "ISFORMULA", "ISOMITTED", "ISOWEEKNUM", "LAMBDA", "LET",
        "LOGNORM.DIST", "LOGNORM.INV", "MAKEARRAY", "MAP", "MAXIFS", "MINIFS", "MODE.MULT", "MODE.SNGL", "MUNIT",
        "NEGBINOM.DIST", "NORM.DIST", "NORM.INV", "NORM.S.DIST", "NORM.S.INV", "NUMBERVALUE", "PDURATION",
        "PERCENTILE.EXC", "PERCENTILE.INC", "PERCENTOF", "PERCENTRANK.EXC", "PERCENTRANK.INC", "PERMUTATIONA", "PHI",
        "PIVOTBY", "POISSON.DIST", "QUARTILE.EXC", "QUARTILE.INC", "RANDARRAY", "RANK.AVG", "RANK.EQ", "REDUCE",
        "REGEXEXTRACT", "REGEXREPLACE", "REGEXTEST", "RRI", "SCAN", "SEC", "SECH", "SEQUENCE", "SHEET", "SHEETS",
        "SKEW.P", "SORTBY", "STDEV.P", "STDEV.S", "STOCKHISTORY", "SWITCH", "T.DIST", "T.DIST.2T", "T.DIST.RT",
        "T.INV", "T.INV.2T", "T.TEST", "TAKE", "TEXTAFTER", "TEXTBEFORE", "TEXTJOIN", "TEXTSPLIT", "TOCOL", "TOROW",
        "TRIMRANGE", "UNICHAR", "UNICODE", "UNIQUE", "VALUETOTEXT", "VAR.P", "VAR.S", "VSTACK", "WEBSERVICE",
        "WEIBULL.DIST", "WRAPCOLS", "WRAPROWS", "XLOOKUP", "XMATCH", "XOR"
    };

    // A function name directly followed by its argument list, not preceded by a character that
    // would make it part of a longer name, a sheet name or an existing prefix.
    private static readonly Regex Call = new(
        @"(?<![A-Za-z0-9_.$'\\])(?<name>[A-Za-z][A-Za-z0-9_.]*)(?=\s*\()",
        RegexOptions.CultureInvariant);

    /// <summary>
    /// Returns the formula as a workbook stores it: every call to a future function carries its
    /// prefix. String literals, names already prefixed and every other function are unchanged.
    /// </summary>
    public static string Qualify(string formula)
    {
        var result = new StringBuilder(formula.Length + 16);
        var start = 0;
        while (start < formula.Length)
        {
            // Copy a string literal ("…", with "" as an escaped quote) through untouched.
            var quote = formula.IndexOf('"', start);
            var end = quote < 0 ? formula.Length : quote;
            result.Append(Call.Replace(formula.Substring(start, end - start), match =>
            {
                var name = match.Groups["name"].Value;
                return Worksheet.Contains(name) ? WorksheetPrefix + name
                    : Prefixed.Contains(name) ? Prefix + name
                    : name;
            }));
            if (quote < 0) break;

            var close = quote + 1;
            while (close < formula.Length)
            {
                if (formula[close] == '"' && close + 1 < formula.Length && formula[close + 1] == '"') close += 2;
                else if (formula[close] == '"') break;
                else close++;
            }
            close = Math.Min(close + 1, formula.Length);
            result.Append(formula, quote, close - quote);
            start = close;
        }
        return result.ToString();
    }
}
