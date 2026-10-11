using System.Globalization;
using System.Text.RegularExpressions;

namespace OfficeAgent.Excel.Calculation;

/// <summary>A rectangle of cells a formula refers to; <see cref="Sheet"/> is null for the formula's own sheet.</summary>
internal readonly struct RangeReference
{
    public RangeReference(string? sheet, int left, int right, int top, int bottom)
    {
        Sheet = sheet;
        Left = left;
        Right = right;
        Top = top;
        Bottom = bottom;
    }

    public string? Sheet { get; }
    public int Left { get; }
    public int Right { get; }
    public int Top { get; }
    public int Bottom { get; }

    public bool Contains(int column, int row) => column >= Left && column <= Right && row >= Top && row <= Bottom;
}

/// <summary>
/// What a formula uses: the functions it calls, the cells and ranges it refers to, the
/// defined names it mentions, and whether it contains a structured (<c>Table[Column]</c>)
/// or external (<c>[Book.xlsx]Sheet!A1</c>) reference. Read from the formula text, without
/// evaluating it.
/// </summary>
internal sealed class FormulaAnalysis
{
    private const int MaxColumn = 16384;
    private const int MaxRow = 1048576;

    private static readonly Regex StringLiteral = new("\"(?:[^\"]|\"\")*\"", RegexOptions.CultureInvariant);

    // One pass over the formula with the literals removed. Alternatives are tried in order at
    // each position: a function call, a (sheet-qualified) cell or range reference, a name.
    private static readonly Regex Token = new(
        @"(?<![A-Za-z0-9_.$'\\])(?:" +
        @"(?<func>(?:_xlfn\.)?(?:_xlws\.)?[A-Za-z][A-Za-z0-9_.]*)\s*\(" +
        @"|(?:(?<sheet>'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!)?" +
        @"(?<ref>\$?[A-Za-z]{1,3}\$?[0-9]+(?::\$?[A-Za-z]{1,3}\$?[0-9]+)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?[0-9]+:\$?[0-9]+)(?![A-Za-z0-9_(!])" +
        @"|(?<name>[A-Za-z_\\][A-Za-z0-9_.\\]*)(?![A-Za-z0-9_.(!])" +
        @")",
        RegexOptions.CultureInvariant);

    private FormulaAnalysis(List<string> functions, List<string> unprefixed, List<RangeReference> references, List<string> names, bool bracket)
    {
        Functions = functions;
        Unprefixed = unprefixed;
        References = references;
        Names = names;
        HasStructuredOrExternalReference = bracket;
    }

    public IReadOnlyList<string> Functions { get; }

    /// <summary>The functions called without an <c>_xlfn.</c> prefix, upper-cased.</summary>
    public IReadOnlyList<string> Unprefixed { get; }
    public IReadOnlyList<RangeReference> References { get; }
    public IReadOnlyList<string> Names { get; }
    public bool HasStructuredOrExternalReference { get; }

    public static FormulaAnalysis Of(string formula)
    {
        var text = StringLiteral.Replace(formula, "\"\"");
        var functions = new List<string>();
        var unprefixed = new List<string>();
        var references = new List<RangeReference>();
        var names = new List<string>();

        foreach (Match match in Token.Matches(text))
        {
            if (match.Groups["func"].Success)
            {
                var name = match.Groups["func"].Value.ToUpperInvariant();
                var prefixed = name.StartsWith("_XLFN.", StringComparison.Ordinal);
                if (prefixed) name = name.Substring(6);
                if (name.StartsWith("_XLWS.", StringComparison.Ordinal)) name = name.Substring(6);
                functions.Add(name);
                if (!prefixed) unprefixed.Add(name);
            }
            else if (match.Groups["ref"].Success)
            {
                var sheet = match.Groups["sheet"].Success ? Unquote(match.Groups["sheet"].Value) : null;
                if (TryRange(match.Groups["ref"].Value, out var left, out var right, out var top, out var bottom))
                    references.Add(new RangeReference(sheet, left, right, top, bottom));
            }
            else if (match.Groups["name"].Success)
            {
                var name = match.Groups["name"].Value;
                if (!name.Equals("TRUE", StringComparison.OrdinalIgnoreCase) && !name.Equals("FALSE", StringComparison.OrdinalIgnoreCase))
                    names.Add(name);
            }
        }

        return new FormulaAnalysis(functions, unprefixed, references, names, text.IndexOf('[') >= 0);
    }

    private static string Unquote(string sheet) =>
        sheet.Length >= 2 && sheet[0] == '\'' && sheet[sheet.Length - 1] == '\''
            ? sheet.Substring(1, sheet.Length - 2).Replace("''", "'")
            : sheet;

    private static bool TryRange(string reference, out int left, out int right, out int top, out int bottom)
    {
        left = right = top = bottom = 0;
        var parts = reference.Replace("$", string.Empty).Split(':');
        if (!TryPoint(parts[0], out var c1, out var r1)) return false;
        var c2 = c1;
        var r2 = r1;
        if (parts.Length == 2 && !TryPoint(parts[1], out c2, out r2)) return false;

        // A whole column (A:A) has no row, a whole row (1:1) no column.
        left = Math.Min(c1 == 0 ? 1 : c1, c2 == 0 ? 1 : c2);
        right = Math.Max(c1 == 0 ? MaxColumn : c1, c2 == 0 ? MaxColumn : c2);
        top = Math.Min(r1 == 0 ? 1 : r1, r2 == 0 ? 1 : r2);
        bottom = Math.Max(r1 == 0 ? MaxRow : r1, r2 == 0 ? MaxRow : r2);
        return true;
    }

    /// <summary>Parses <c>B7</c>, <c>B</c> (column only) or <c>7</c> (row only); a missing part is 0.</summary>
    private static bool TryPoint(string point, out int column, out int row)
    {
        column = 0;
        row = 0;
        var i = 0;
        while (i < point.Length && char.IsLetter(point[i]))
        {
            column = column * 26 + (char.ToUpperInvariant(point[i]) - 'A' + 1);
            i++;
        }
        if (i < point.Length && !int.TryParse(point.Substring(i), NumberStyles.None, CultureInfo.InvariantCulture, out row))
            return false;
        return (column > 0 || row > 0) && column <= MaxColumn && row <= MaxRow;
    }

    /// <summary>The 1-based column and row of an A1 address such as <c>C12</c>.</summary>
    public static (int Column, int Row) Point(string address)
    {
        TryPoint(address.Replace("$", string.Empty), out var column, out var row);
        return (column, row);
    }
}
