using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Excel;
using S = DocumentFormat.OpenXml.Spreadsheet;
using DocFormat = OfficeAgent.Abstractions.DocumentFormat;

namespace OfficeAgent.Tests;

public class ExcelModuleTests
{
    [Fact]
    public void Blank_workbook_is_inspectable_and_range_is_bounded()
    {
        var module = new ExcelModule();
        var client = new OfficeAgentClient(module);
        var inspection = client.Inspect(module.CreateBlank());

        Assert.Equal(DocFormat.Excel, inspection.Format);
        Assert.Equal(1U, Assert.Single(inspection.Worksheets).SheetId);
        Assert.Empty(inspection.Cells);

        var bounded = client.Inspect(Handle(XlsxFactory.WorkbookWithTable()), new InspectOptions
        {
            SheetId = 7U,
            Range = "A2:B2",
            MaximumCells = 1
        });
        var cell = Assert.Single(bounded.Cells);
        Assert.Equal("A2", cell.Anchor.Address);
    }

    [Fact]
    public void Find_can_choose_displayed_or_raw_cell_values()
    {
        var client = new OfficeAgentClient(new ExcelModule());
        var bytes = XlsxFactory.WorkbookWithTable();

        var hit = Assert.Single(client.Find(Handle(bytes), new FindQuery("EMEA")
        {
            Options = new MatchOptions { SpreadsheetValueView = SpreadsheetValueView.Displayed }
        }));

        var anchor = Assert.IsType<CellAnchor>(hit.Anchor);
        Assert.Equal(7U, anchor.SheetId);
        Assert.Equal("A2", anchor.Address);
    }

    [Fact]
    public void Cell_formula_preserves_style_and_requests_recalculation()
    {
        var client = new OfficeAgentClient(new ExcelModule());
        using var applied = client.Commit(Handle(XlsxFactory.WorkbookWithTable()), new DocumentPlan
        {
            Format = DocFormat.Excel,
            Operations = new PlanOperation[]
            {
                new SetCellOp
                {
                    Target = new CellAnchor { SheetId = 7U, Address = "B2" },
                    Formula = "20+22"
                }
            }
        });
        Assert.True(applied.Committed, string.Join("; ", applied.Report.Errors.Select(e => e.Message)));

        using var stream = new MemoryStream(applied.ToBytes());
        using var document = SpreadsheetDocument.Open(stream, false);
        var part = document.WorkbookPart!.WorksheetParts.Single();
        var cell = part.Worksheet.Descendants<S.Cell>().Single(c => c.CellReference?.Value == "B2");
        Assert.Equal(1U, cell.StyleIndex?.Value);
        Assert.Equal("20+22", cell.CellFormula?.Text);
        Assert.Null(cell.CellValue);
        Assert.True(document.WorkbookPart.Workbook.CalculationProperties?.FullCalculationOnLoad?.Value);
        Assert.Equal("SUM(B2)", part.Worksheet.Descendants<S.Cell>().Single(c => c.CellReference?.Value == "C2").CellFormula?.Text);
    }

    [Fact]
    public void Table_rows_and_comments_are_managed_without_rebuilding_the_sheet()
    {
        var client = new OfficeAgentClient(new ExcelModule());
        using var applied = client.Commit(Handle(XlsxFactory.WorkbookWithTable()), new DocumentPlan
        {
            Format = DocFormat.Excel,
            Operations = new PlanOperation[]
            {
                new AppendTableRowsOp
                {
                    Target = new NodeAnchor { Kind = "spreadsheetTable", Path = "table#7/Sales" },
                    Rows = new[] { new[] { "APAC", "15" } }
                },
                new CommentOp
                {
                    Target = new CellAnchor { SheetId = 7U, Address = "A2" },
                    Action = CommentAction.Add,
                    Author = "Reviewer",
                    Text = "Confirm region"
                }
            }
        });
        Assert.True(applied.Committed, string.Join("; ", applied.Report.Errors.Select(e => e.Message)));

        using var stream = new MemoryStream(applied.ToBytes());
        using var document = SpreadsheetDocument.Open(stream, false);
        var sheet = document.WorkbookPart!.WorksheetParts.Single();
        Assert.Equal("A1:B3", sheet.TableDefinitionParts.Single().Table!.Reference!.Value);
        Assert.Equal("APAC", string.Concat(sheet.Worksheet.Descendants<S.Cell>()
            .Single(c => c.CellReference?.Value == "A3").Descendants<S.Text>().Select(t => t.Text)));
        var revenue = sheet.Worksheet.Descendants<S.Cell>().Single(c => c.CellReference?.Value == "B3");
        Assert.Null(revenue.DataType);
        Assert.Equal("15", revenue.CellValue?.Text);
        Assert.Equal("SUM(B2)", sheet.Worksheet.Descendants<S.Cell>().Single(c => c.CellReference?.Value == "C2").CellFormula?.Text);
        Assert.Equal("Confirm region", string.Concat(sheet.WorksheetCommentsPart!.Comments!.Descendants<S.Text>().Select(t => t.Text)));
        Assert.Single(sheet.VmlDrawingParts);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(document).ToList();
        Assert.True(errors.Count == 0, string.Join("; ", errors.Select(e => e.Description)));

        using var removed = client.Commit(Handle(applied.ToBytes()), new DocumentPlan
        {
            Format = DocFormat.Excel,
            Operations = new PlanOperation[]
            {
                new CommentOp
                {
                    Target = new NodeAnchor
                    {
                        Kind = "cellComment",
                        Path = "comment#7/A2"
                    },
                    Action = CommentAction.Remove
                }
            }
        });
        Assert.True(removed.Committed, string.Join("; ", removed.Report.Errors.Select(e => e.Message)));
        var afterRemoval = client.Inspect(Handle(removed.ToBytes()));
        Assert.DoesNotContain(afterRemoval.Nodes, node => node.Kind == "cellComment");
    }

    private static StreamHandle Handle(byte[] bytes) => new(new MemoryStream(bytes, writable: false));
}
