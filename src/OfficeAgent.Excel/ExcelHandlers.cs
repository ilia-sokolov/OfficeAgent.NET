using System.Globalization;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeAgent.Excel;

internal static class ExcelTargets
{
    public static (SpreadsheetDocument Document, S.Sheet Sheet, WorksheetPart Part, S.Cell? Cell)? Cell(
        ApplyContext context, CellAnchor anchor, bool create)
    {
        var document = (SpreadsheetDocument)context.Package.Package;
        var sheet = SpreadsheetPartUtility.ResolveSheet(document, anchor.SheetId);
        if (sheet is null) return null;
        return (document, sheet.Value.Sheet, sheet.Value.Part,
            SpreadsheetPartUtility.Cell(sheet.Value.Part, anchor.Address, create));
    }

    public static (SpreadsheetDocument Document, uint SheetId, WorksheetPart Part, TableDefinitionPart Table)? Table(
        ApplyContext context, NodeAnchor anchor)
    {
        if (!TryTablePath(anchor.Path, out var sheetId, out var name)) return null;
        var document = (SpreadsheetDocument)context.Package.Package;
        var sheet = SpreadsheetPartUtility.ResolveSheet(document, sheetId);
        if (sheet is null) return null;
        var table = sheet.Value.Part.TableDefinitionParts.FirstOrDefault(candidate =>
            string.Equals(candidate.Table?.DisplayName?.Value, name, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(candidate.Table?.Name?.Value, name, StringComparison.OrdinalIgnoreCase));
        return table is null ? null : (document, sheetId, sheet.Value.Part, table);
    }

    public static bool TryTablePath(string path, out uint sheetId, out string name)
    {
        sheetId = 0;
        name = string.Empty;
        var value = path.StartsWith("table#", StringComparison.OrdinalIgnoreCase) ? path.Substring(6) : path;
        var slash = value.IndexOf('/');
        if (slash <= 0 || !uint.TryParse(value.Substring(0, slash), out sheetId)) return false;
        name = value.Substring(slash + 1);
        return name.Length > 0;
    }
}

internal sealed class SetCellHandler : IOperationHandler
{
    public bool CanHandle(PlanOperation operation) => operation is SetCellOp { Target: CellAnchor };

    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (SetCellOp)operation;
        var anchor = (CellAnchor)op.Target;
        if (!SpreadsheetPartUtility.TryParseCell(anchor.Address, out _, out _))
            return Fail("setCell needs one valid A1 cell address.", anchor);
        if (op.Value is { } typedValue)
        {
            try { SpreadsheetPartUtility.WriteValue(new S.Cell(), typedValue, op.ValueKind); }
            catch (ArgumentException error) { return Fail(error.Message, anchor); }
        }
        var target = ExcelTargets.Cell(context, anchor, create: false);
        if (target is null)
            return Fail($"Worksheet id {anchor.SheetId} does not exist.", anchor, ValidationErrorCodes.AnchorNotFound);
        var before = target.Value.Cell is null ? string.Empty : SpreadsheetPartUtility.DisplayValue(target.Value.Document, target.Value.Cell);
        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "setCell",
            Before = before,
            After = op.Formula is { Length: > 0 } ? "=" + op.Formula : op.Value ?? string.Empty,
            Context = $"{target.Value.Sheet.Name?.Value}!{anchor.Address}"
        });
    }

    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (SetCellOp)operation;
        var anchor = (CellAnchor)op.Target;
        var target = ExcelTargets.Cell(context, anchor, create: true)
            ?? throw new InvalidOperationException($"Worksheet id {anchor.SheetId} vanished before apply.");
        var cell = target.Cell ?? throw new InvalidOperationException($"Cell '{anchor.Address}' could not be created.");

        cell.CellFormula = null;
        cell.CellValue = null;
        cell.InlineString = null;
        cell.DataType = null;
        if (op.Formula is { Length: > 0 } formula)
        {
            cell.CellFormula = new S.CellFormula(formula.StartsWith("=", StringComparison.Ordinal) ? formula.Substring(1) : formula);
            if (op.Value is { } cached) cell.CellValue = new S.CellValue(cached);
            SpreadsheetPartUtility.RecalculateOnOpen(target.Document);
        }
        else if (op.Value is { } value)
            SpreadsheetPartUtility.WriteValue(cell, value, op.ValueKind);
        target.Part.Worksheet.Save();
        target.Document.WorkbookPart!.Workbook.Save();
    }

    private static OperationPreview Fail(string message, Anchor anchor, string code = ValidationErrorCodes.InvalidOperation) =>
        OperationPreview.Fail(new ValidationError(code, message, anchor));
}

internal sealed class AppendTableRowsHandler : IOperationHandler
{
    public bool CanHandle(PlanOperation operation) =>
        operation is AppendTableRowsOp { Target: NodeAnchor { Kind: "spreadsheetTable" } };

    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (AppendTableRowsOp)operation;
        var anchor = (NodeAnchor)op.Target;
        var target = ExcelTargets.Table(context, anchor);
        if (target is null)
            return Fail($"No Excel table with path '{anchor.Path}'.", anchor, ValidationErrorCodes.AnchorNotFound);
        var table = target.Value.Table.Table!;
        if (!SpreadsheetPartUtility.TryParseRange(table.Reference?.Value ?? string.Empty,
                out var left, out _, out var right, out var bottom))
            return Fail("The table has an invalid range.", anchor);
        if ((table.TotalsRowCount?.Value ?? 0) > 0 || table.TotalsRowShown?.Value == true)
            return Fail("appendTableRows does not append to a table with a totals row.", anchor);
        var columns = right - left + 1;
        if (op.Rows.Count == 0 || op.Rows.Any(row => row.Count != columns))
            return Fail($"Each appended row must contain exactly {columns} values.", anchor);
        for (var offset = 1; offset <= op.Rows.Count; offset++)
            for (var column = left; column <= right; column++)
            {
                var address = SpreadsheetPartUtility.ColumnName(column) + (bottom + (uint)offset).ToString(CultureInfo.InvariantCulture);
                var occupied = SpreadsheetPartUtility.Cell(target.Value.Part, address, create: false);
                if (occupied is not null && (occupied.CellValue is not null || occupied.InlineString is not null || occupied.CellFormula is not null))
                    return Fail($"Appending would overwrite populated cell {address} below the table.", anchor);
            }
        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "appendTableRows",
            Before = table.Reference?.Value ?? string.Empty,
            After = $"{op.Rows.Count} row(s) appended",
            Context = table.DisplayName?.Value ?? string.Empty,
            BlastRadius = op.Rows.Count
        });
    }

    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (AppendTableRowsOp)operation;
        var anchor = (NodeAnchor)op.Target;
        var target = ExcelTargets.Table(context, anchor)
            ?? throw new InvalidOperationException($"Table '{anchor.Path}' vanished before apply.");
        var table = target.Table.Table!;
        SpreadsheetPartUtility.TryParseRange(table.Reference!.Value!, out var left, out var top, out var right, out var bottom);

        for (var rowOffset = 0; rowOffset < op.Rows.Count; rowOffset++)
            for (var columnOffset = 0; columnOffset < op.Rows[rowOffset].Count; columnOffset++)
            {
                var column = left + columnOffset;
                var row = bottom + (uint)rowOffset + 1U;
                var address = SpreadsheetPartUtility.ColumnName(column) + row.ToString(CultureInfo.InvariantCulture);
                var cell = SpreadsheetPartUtility.Cell(target.Part, address, create: true)!;
                var templateAddress = SpreadsheetPartUtility.ColumnName(column) + bottom.ToString(CultureInfo.InvariantCulture);
                var template = SpreadsheetPartUtility.Cell(target.Part, templateAddress, create: false);
                if (template?.StyleIndex is not null) cell.StyleIndex = template.StyleIndex.Value;
                SpreadsheetPartUtility.WriteValue(cell, op.Rows[rowOffset][columnOffset] ?? string.Empty);
            }

        var newBottom = bottom + (uint)op.Rows.Count;
        var updated = SpreadsheetPartUtility.ColumnName(left) + top.ToString(CultureInfo.InvariantCulture) + ":" +
            SpreadsheetPartUtility.ColumnName(right) + newBottom.ToString(CultureInfo.InvariantCulture);
        table.Reference = updated;
        if (table.AutoFilter is not null) table.AutoFilter.Reference = updated;
        target.Part.Worksheet.Save();
        table.Save();
    }

    private static OperationPreview Fail(string message, Anchor anchor, string code = ValidationErrorCodes.InvalidOperation) =>
        OperationPreview.Fail(new ValidationError(code, message, anchor));
}

internal sealed class CellCommentHandler : IOperationHandler
{
    public bool CanHandle(PlanOperation operation) => operation is CommentOp
    {
        Target: CellAnchor or NodeAnchor { Kind: "cellComment" },
        Action: CommentAction.Add or CommentAction.Remove
    };

    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (CommentOp)operation;
        if (!TryAnchor(op.Target, out var anchor))
            return OperationPreview.Fail(new ValidationError(ValidationErrorCodes.InvalidOperation,
                "Excel comments need a cell anchor, or a cellComment node returned by inspection.", op.Target));
        var target = ExcelTargets.Cell(context, anchor, create: false);
        if (target is null)
            return OperationPreview.Fail(new ValidationError(ValidationErrorCodes.AnchorNotFound,
                $"Worksheet id {anchor.SheetId} does not exist.", anchor));
        var existing = target.Value.Part.WorksheetCommentsPart?.Comments?.CommentList?
            .Elements<S.Comment>().FirstOrDefault(c =>
                string.Equals(c.Reference?.Value, anchor.Address, StringComparison.OrdinalIgnoreCase));
        if (op.Action == CommentAction.Add && existing is not null)
            return OperationPreview.Fail(new ValidationError(ValidationErrorCodes.InvalidOperation,
                $"Cell {anchor.Address} already has a comment; remove it before adding a replacement.", anchor));
        if (op.Action == CommentAction.Remove && existing is null)
            return OperationPreview.Fail(new ValidationError(ValidationErrorCodes.AnchorNotFound,
                $"Cell {anchor.Address} has no comment.", anchor));
        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "comment",
            Before = existing is null ? string.Empty : CommentText(existing),
            After = op.Action == CommentAction.Add ? op.Text : string.Empty,
            Context = $"{target.Value.Sheet.Name?.Value}!{anchor.Address}"
        });
    }

    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (CommentOp)operation;
        TryAnchor(op.Target, out var anchor);
        var target = ExcelTargets.Cell(context, anchor, create: true)
            ?? throw new InvalidOperationException("Comment worksheet vanished before apply.");
        if (op.Action == CommentAction.Add) Add(target.Part, anchor.Address, op.Author, op.Text);
        else Remove(target.Part, anchor.Address);
        target.Part.Worksheet.Save();
    }

    private static void Add(WorksheetPart part, string address, string author, string text)
    {
        var commentsPart = part.WorksheetCommentsPart ?? part.AddNewPart<WorksheetCommentsPart>();
        commentsPart.Comments ??= new S.Comments(new S.Authors(), new S.CommentList());
        var authors = commentsPart.Comments.Authors ?? commentsPart.Comments.PrependChild(new S.Authors());
        var authorList = authors.Elements<S.Author>().ToList();
        var authorId = authorList.FindIndex(a => string.Equals(a.Text, author, StringComparison.Ordinal));
        if (authorId < 0) { authorId = authorList.Count; authors.Append(new S.Author(author)); }
        commentsPart.Comments.CommentList ??= new S.CommentList();
        commentsPart.Comments.CommentList.Append(new S.Comment(
            new S.CommentText(new S.Run(new S.Text(text))))
        {
            Reference = address.ToUpperInvariant(),
            AuthorId = (uint)authorId
        });
        commentsPart.Comments.Save();
        EnsureVml(part, address, add: true);
    }

    private static void Remove(WorksheetPart part, string address)
    {
        var commentsPart = part.WorksheetCommentsPart;
        var comment = commentsPart?.Comments?.CommentList?.Elements<S.Comment>().FirstOrDefault(c =>
            string.Equals(c.Reference?.Value, address, StringComparison.OrdinalIgnoreCase));
        comment?.Remove();
        commentsPart?.Comments?.Save();
        EnsureVml(part, address, add: false);
    }

    private static void EnsureVml(WorksheetPart part, string address, bool add)
    {
        var vmlPart = part.VmlDrawingParts.FirstOrDefault();
        if (vmlPart is null)
        {
            if (!add) return;
            vmlPart = part.AddNewPart<VmlDrawingPart>();
            using var initial = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(
                "<xml xmlns:v='urn:schemas-microsoft-com:vml' xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:x='urn:schemas-microsoft-com:office:excel'><o:shapelayout v:ext='edit'><o:idmap v:ext='edit' data='1'/></o:shapelayout><v:shapetype id='_x0000_t202' coordsize='21600,21600' o:spt='202' path='m,l,21600r21600,l21600,xe'><v:stroke joinstyle='miter'/><v:path gradientshapeok='t' o:connecttype='rect'/></v:shapetype></xml>"));
            vmlPart.FeedData(initial);
            var legacy = new S.LegacyDrawing { Id = part.GetIdOfPart(vmlPart) };
            var tableParts = part.Worksheet.GetFirstChild<S.TableParts>();
            if (tableParts is null) part.Worksheet.Append(legacy); else part.Worksheet.InsertBefore(legacy, tableParts);
        }

        XDocument xml;
        using (var input = vmlPart.GetStream(FileMode.Open, FileAccess.Read)) xml = XDocument.Load(input);
        XNamespace x = "urn:schemas-microsoft-com:office:excel";
        XNamespace v = "urn:schemas-microsoft-com:vml";
        SpreadsheetPartUtility.TryParseCell(address, out var column, out var row);
        var matching = xml.Descendants(v + "shape").Where(shape =>
        {
            var data = shape.Element(x + "ClientData");
            return data?.Element(x + "Row")?.Value == (row - 1).ToString(CultureInfo.InvariantCulture) &&
                   data.Element(x + "Column")?.Value == (column - 1).ToString(CultureInfo.InvariantCulture);
        }).ToList();
        foreach (var shape in matching) shape.Remove();
        if (add)
        {
            var id = "_x0000_s" + (1025 + xml.Descendants(v + "shape").Count()).ToString(CultureInfo.InvariantCulture);
            xml.Root!.Add(new XElement(v + "shape",
                new XAttribute("id", id), new XAttribute("type", "#_x0000_t202"),
                new XAttribute("style", "position:absolute;margin-left:80pt;margin-top:5pt;width:108pt;height:59.25pt;z-index:1;visibility:hidden"),
                new XAttribute("fillcolor", "#ffffe1"),
                new XElement(v + "fill", new XAttribute("color2", "#ffffe1")),
                new XElement(v + "shadow", new XAttribute("on", "t"), new XAttribute("color", "black"), new XAttribute("obscured", "t")),
                new XElement(v + "path", new XAttribute(XName.Get("connecttype", "urn:schemas-microsoft-com:office:office"), "none")),
                new XElement(v + "textbox", new XAttribute("style", "mso-direction-alt:auto"), new XElement("div", new XAttribute("style", "text-align:left"))),
                new XElement(x + "ClientData", new XAttribute("ObjectType", "Note"),
                    new XElement(x + "MoveWithCells"), new XElement(x + "SizeWithCells"),
                    new XElement(x + "AutoFill", "False"), new XElement(x + "Row", row - 1),
                    new XElement(x + "Column", column - 1))));
        }
        using var output = vmlPart.GetStream(FileMode.Create, FileAccess.Write);
        xml.Save(output, SaveOptions.DisableFormatting);
    }

    private static bool TryAnchor(Anchor source, out CellAnchor anchor)
    {
        if (source is CellAnchor cell) { anchor = cell; return true; }
        if (source is NodeAnchor { Kind: "cellComment" } node)
        {
            var value = node.Path.StartsWith("comment#", StringComparison.OrdinalIgnoreCase) ? node.Path.Substring(8) : node.Path;
            var slash = value.IndexOf('/');
            if (slash > 0 && uint.TryParse(value.Substring(0, slash), out var sheetId))
            {
                anchor = new CellAnchor { Id = node.Id, SheetId = sheetId, Address = value.Substring(slash + 1) };
                return true;
            }
        }
        anchor = new CellAnchor();
        return false;
    }

    private static string CommentText(S.Comment comment) =>
        string.Concat(comment.CommentText?.Descendants<S.Text>().Select(t => t.Text) ?? Enumerable.Empty<string>());
}
