namespace OfficeAgent.Core;

/// <summary>The checks and preview text the Word and PowerPoint table fills share.</summary>
internal static class TableFill
{
    /// <summary>Why the rows cannot fill the table, or <see langword="null"/> when they can.</summary>
    public static string? Check(IReadOnlyList<IReadOnlyList<string>> rows, int columns, bool hasRows, bool merged)
    {
        if (!hasRows) return "The table has no rows to fill.";
        if (merged) return "The table has merged cells, so its cells cannot be filled by position. Unmerge them first.";
        if (rows.Count == 0) return "fillTable needs at least one row.";
        for (var i = 0; i < rows.Count; i++)
            if (rows[i] is null || rows[i].Count != columns)
                return $"Row {i} has {rows[i]?.Count ?? 0} value(s) but the table has {columns} column(s). " +
                       "Every row must fill every column; fillTable does not add or remove columns.";
        return null;
    }

    /// <summary>A short preview of the rows a fill writes.</summary>
    public static string Describe(IReadOnlyList<IReadOnlyList<string>> rows) =>
        $"[{rows.Count} row(s)] " + string.Join(" / ", rows.Take(2).Select(row => string.Join(" | ", row))) +
        (rows.Count > 2 ? $" / (+{rows.Count - 2} more)" : string.Empty);
}
