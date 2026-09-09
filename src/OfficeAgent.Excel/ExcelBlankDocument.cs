using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeAgent.Excel;

internal static class ExcelBlankDocument
{
    public static byte[] Create()
    {
        using var stream = new MemoryStream();
        using (var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new S.Workbook();
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new S.Worksheet(new S.SheetData());
            workbookPart.Workbook.AppendChild(new S.Sheets()).Append(new S.Sheet
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1U,
                Name = "Sheet1"
            });
            worksheetPart.Worksheet.Save();
            workbookPart.Workbook.Save();
        }
        return stream.ToArray();
    }
}
