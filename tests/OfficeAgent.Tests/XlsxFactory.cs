using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeAgent.Tests;

internal static class XlsxFactory
{
    public static byte[] WorkbookWithTable()
    {
        using var stream = new MemoryStream();
        using (var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new S.Workbook();
            var styles = workbookPart.AddNewPart<WorkbookStylesPart>();
            styles.Stylesheet = new S.Stylesheet(
                new S.Fonts(new S.Font()) { Count = 1 },
                new S.Fills(new S.Fill(new S.PatternFill { PatternType = S.PatternValues.None }),
                    new S.Fill(new S.PatternFill { PatternType = S.PatternValues.Gray125 })) { Count = 2 },
                new S.Borders(new S.Border()) { Count = 1 },
                new S.CellStyleFormats(new S.CellFormat()) { Count = 1 },
                new S.CellFormats(new S.CellFormat(), new S.CellFormat { FillId = 1U, ApplyFill = true }) { Count = 2 });

            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            var tablePart = worksheetPart.AddNewPart<TableDefinitionPart>();
            tablePart.Table = new S.Table(
                new S.AutoFilter { Reference = "A1:B2" },
                new S.TableColumns(
                    new S.TableColumn { Id = 1U, Name = "Region" },
                    new S.TableColumn { Id = 2U, Name = "Revenue" }) { Count = 2U },
                new S.TableStyleInfo { Name = "TableStyleMedium2", ShowRowStripes = true })
            {
                Id = 1U,
                Name = "Sales",
                DisplayName = "Sales",
                Reference = "A1:B2",
                TotalsRowShown = false
            };
            var data = new S.SheetData(
                new S.Row(
                    Inline("A1", "Region"), Inline("B1", "Revenue"), Inline("C1", "Unrelated")) { RowIndex = 1U },
                new S.Row(
                    Inline("A2", "EMEA"), new S.Cell(new S.CellValue("10")) { CellReference = "B2", StyleIndex = 1U },
                    new S.Cell(new S.CellFormula("SUM(B2)"), new S.CellValue("10")) { CellReference = "C2" }) { RowIndex = 2U });
            worksheetPart.Worksheet = new S.Worksheet(
                new S.SheetDimension { Reference = "A1:C2" }, data,
                new S.TableParts(new S.TablePart { Id = worksheetPart.GetIdOfPart(tablePart) }) { Count = 1U });

            workbookPart.Workbook.AppendChild(new S.Sheets()).Append(new S.Sheet
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = 7U,
                Name = "Data"
            });
            styles.Stylesheet.Save();
            tablePart.Table.Save();
            worksheetPart.Worksheet.Save();
            workbookPart.Workbook.Save();
        }
        return stream.ToArray();
    }

    private static S.Cell Inline(string reference, string value) => new()
    {
        CellReference = reference,
        DataType = S.CellValues.InlineString,
        InlineString = new S.InlineString(new S.Text(value))
    };
}
