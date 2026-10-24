namespace OfficeAgent.Core;

/// <summary>
/// Reads the range references Excel writes and people type: <c>Summary!B2</c>,
/// <c>'Q3 Sales'!$A$1:$D$6</c>, or a bare name that a workbook defines.
/// </summary>
internal static class WorkbookReference
{
    /// <summary>One parsed reference: a sheet and an A1 range, or a name to look up.</summary>
    internal readonly struct Parsed
    {
        internal Parsed(string? sheet, string? range, string? name)
        {
            Sheet = sheet;
            Range = range;
            Name = name;
        }

        /// <summary>Gets the sheet name, unquoted, when the reference names one.</summary>
        internal string? Sheet { get; }

        /// <summary>Gets the A1 range without <c>$</c> signs, such as <c>B2</c> or <c>A1:D6</c>.</summary>
        internal string? Range { get; }

        /// <summary>Gets the defined name or table name, when the reference is not a sheet range.</summary>
        internal string? Name { get; }
    }

    /// <summary>Parses a reference, or returns <see langword="false"/> when it is neither a sheet range nor a plausible name.</summary>
    internal static bool TryParse(string text, out Parsed parsed)
    {
        parsed = default;
        var value = (text ?? string.Empty).Trim();
        if (value.StartsWith("=", StringComparison.Ordinal)) value = value.Substring(1).Trim();
        if (value.Length == 0) return false;

        var bang = value.LastIndexOf('!');
        if (bang < 0)
        {
            // A bare cell such as "B2" names no sheet, and guessing the first one would bind a
            // report to whichever sheet happens to come first.
            if (IsRange(Strip(value))) return false;
            if (value.Any(c => char.IsWhiteSpace(c) || c is '\'' or '"' or ':' or '$')) return false;
            parsed = new Parsed(null, null, value);
            return true;
        }

        var sheet = value.Substring(0, bang);
        var range = Strip(value.Substring(bang + 1));
        if (sheet.Length >= 2 && sheet[0] == '\'' && sheet[sheet.Length - 1] == '\'')
            sheet = sheet.Substring(1, sheet.Length - 2).Replace("''", "'");
        if (sheet.Length == 0 || !IsRange(range)) return false;

        parsed = new Parsed(sheet, range.ToUpperInvariant(), null);
        return true;
    }

    /// <summary>Whether text is a cell or a two-cell range in A1 notation.</summary>
    internal static bool IsRange(string range) =>
        SpreadsheetPartUtility.TryParseCell(range, out _, out _) ||
        SpreadsheetPartUtility.TryParseRange(range, out _, out _, out _, out _);

    /// <summary>The range's corners, with a single cell as a one-cell range.</summary>
    internal static bool TryBounds(string range, out int left, out uint top, out int right, out uint bottom)
    {
        if (SpreadsheetPartUtility.TryParseRange(range, out left, out top, out right, out bottom))
            return true;
        if (SpreadsheetPartUtility.TryParseCell(range, out left, out top))
        {
            right = left;
            bottom = top;
            return true;
        }

        right = 0;
        bottom = 0;
        return false;
    }

    private static string Strip(string range) => range.Replace("$", string.Empty).Trim();
}
