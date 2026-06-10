using OfficeAgent.Core;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.Word;

/// <summary>
/// Resolves table-related node paths to OpenXml elements:
/// <c>table#N</c>, <c>table#N/row#M</c>, <c>table#N/cell#R/C</c>.
/// </summary>
internal static class TableLocator
{
    public static Table? FindTable(IOpenXmlPackage package, string path)
    {
        var head = SplitFirst(path, '/');
        if (!head.StartsWith("table#", System.StringComparison.Ordinal)) return null;
        if (!int.TryParse(head.Substring("table#".Length), out var target)) return null;

        int i = 0;
        foreach (var (root, _) in WordModel.TextHosts(package))
            foreach (var table in root.Descendants<Table>())
            {
                if (i == target) return table;
                i++;
            }
        return null;
    }

    public static TableRow? FindRow(IOpenXmlPackage package, string path)
    {
        var parts = path.Split('/');
        if (parts.Length < 2) return null;
        var table = FindTable(package, parts[0]);
        if (table is null) return null;
        if (!parts[1].StartsWith("row#", System.StringComparison.Ordinal)) return null;
        if (!int.TryParse(parts[1].Substring("row#".Length), out var rowIndex)) return null;
        return table.Elements<TableRow>().ElementAtOrDefault(rowIndex);
    }

    public static TableCell? FindCell(IOpenXmlPackage package, string path)
    {
        // Accept either "table#N/cell#R/C" (preferred) or "table#N/row#R/cell#C".
        var parts = path.Split('/');
        if (parts.Length < 3) return null;
        var table = FindTable(package, parts[0]);
        if (table is null) return null;

        int rowIndex, colIndex;
        if (parts[1].StartsWith("cell#", System.StringComparison.Ordinal)
            && int.TryParse(parts[1].Substring("cell#".Length), out rowIndex)
            && int.TryParse(parts[2], out colIndex))
        {
            // ok
        }
        else if (parts.Length >= 3
                 && parts[1].StartsWith("row#", System.StringComparison.Ordinal)
                 && parts[2].StartsWith("cell#", System.StringComparison.Ordinal)
                 && int.TryParse(parts[1].Substring("row#".Length), out rowIndex)
                 && int.TryParse(parts[2].Substring("cell#".Length), out colIndex))
        {
            // ok
        }
        else
        {
            return null;
        }

        var row = table.Elements<TableRow>().ElementAtOrDefault(rowIndex);
        return row?.Elements<TableCell>().ElementAtOrDefault(colIndex);
    }

    private static string SplitFirst(string s, char sep)
    {
        var idx = s.IndexOf(sep);
        return idx < 0 ? s : s.Substring(0, idx);
    }
}
