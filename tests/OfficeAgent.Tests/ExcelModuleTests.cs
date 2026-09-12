using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml;
using System.Xml.Linq;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
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
    public void Distinct_cell_addresses_do_not_conflict_without_optional_anchor_ids()
    {
        var client = new OfficeAgentClient(new ExcelModule());
        using var applied = client.Commit(Handle(XlsxFactory.WorkbookWithTable()), new DocumentPlan
        {
            Format = DocFormat.Excel,
            Operations = new PlanOperation[]
            {
                new SetCellOp
                {
                    Target = new CellAnchor { SheetId = 7U, Address = "A2" },
                    Value = "APAC"
                },
                new SetCellOp
                {
                    Target = new CellAnchor { SheetId = 7U, Address = "B2" },
                    Value = "25"
                }
            }
        });

        Assert.True(applied.Committed, string.Join("; ", applied.Report.Errors.Select(e => e.Message)));
        using var stream = new MemoryStream(applied.ToBytes());
        using var document = SpreadsheetDocument.Open(stream, false);
        var sheet = document.WorkbookPart!.WorksheetParts.Single();
        Assert.Equal("APAC", string.Concat(sheet.Worksheet.Descendants<S.Cell>()
            .Single(c => c.CellReference?.Value == "A2").Descendants<S.Text>().Select(t => t.Text)));
        Assert.Equal("25", sheet.Worksheet.Descendants<S.Cell>()
            .Single(c => c.CellReference?.Value == "B2").CellValue?.Text);
        Assert.True(document.WorkbookPart.Workbook.CalculationProperties?.FullCalculationOnLoad?.Value);
    }

    [Fact]
    public void Equivalent_cell_addresses_conflict_after_normalization()
    {
        var client = new OfficeAgentClient(new ExcelModule());
        using var applied = client.Commit(Handle(XlsxFactory.WorkbookWithTable()), new DocumentPlan
        {
            Format = DocFormat.Excel,
            Operations = new PlanOperation[]
            {
                new SetCellOp
                {
                    Target = new CellAnchor { SheetId = 7U, Address = "A2" },
                    Value = "one"
                },
                new SetCellOp
                {
                    Target = new CellAnchor { SheetId = 7U, Address = "$a$2" },
                    Value = "two"
                }
            }
        });

        Assert.False(applied.Committed);
        Assert.Contains(applied.Report.Errors,
            error => error.Code == ValidationErrorCodes.OperationConflict);
    }

    [Fact]
    public void The_same_address_on_different_sheets_does_not_conflict()
    {
        var client = new OfficeAgentClient(new ExcelModule());
        using var applied = client.Commit(Handle(XlsxFactory.WorkbookWithTwoSheets()), new DocumentPlan
        {
            Format = DocFormat.Excel,
            Operations = new PlanOperation[]
            {
                new SetCellOp
                {
                    Target = new CellAnchor { SheetId = 1U, Address = "A1" },
                    Value = "one"
                },
                new SetCellOp
                {
                    Target = new CellAnchor { SheetId = 2U, Address = "A1" },
                    Value = "two"
                }
            }
        });

        Assert.True(applied.Committed, string.Join("; ", applied.Report.Errors.Select(e => e.Message)));
    }

    [Fact]
    public void Replacing_one_shared_formula_member_is_rejected()
    {
        var client = new OfficeAgentClient(new ExcelModule());
        using var applied = client.Commit(Handle(XlsxFactory.WorkbookWithSharedFormula()), new DocumentPlan
        {
            Format = DocFormat.Excel,
            Operations = new PlanOperation[]
            {
                new SetCellOp
                {
                    Target = new CellAnchor { SheetId = 1U, Address = "A1" },
                    Value = "99"
                }
            }
        });

        Assert.False(applied.Committed);
        Assert.Contains(applied.Report.Errors,
            error => error.Message.Contains("shared-formula", StringComparison.Ordinal));
    }

    [Fact]
    public void Replacing_one_array_formula_member_is_rejected()
    {
        var client = new OfficeAgentClient(new ExcelModule());
        using var applied = client.Commit(Handle(XlsxFactory.WorkbookWithArrayFormula()), new DocumentPlan
        {
            Format = DocFormat.Excel,
            Operations = new PlanOperation[]
            {
                new SetCellOp
                {
                    Target = new CellAnchor { SheetId = 1U, Address = "A2" },
                    Value = "99"
                }
            }
        });

        Assert.False(applied.Committed);
        Assert.Contains(applied.Report.Errors,
            error => error.Message.Contains("array formula", StringComparison.Ordinal));
    }

    [Fact]
    public void Shared_string_changes_invalidate_the_snapshot()
    {
        var client = new OfficeAgentClient(new ExcelModule());
        var original = XlsxFactory.WorkbookWithSharedString("old");
        var changed = XlsxFactory.WorkbookWithSharedString("new");
        var originalSnapshot = client.Inspect(Handle(original)).Snapshot;
        var changedSnapshot = client.Inspect(Handle(changed)).Snapshot;
        Assert.NotEqual(originalSnapshot, changedSnapshot);

        using var applied = client.Commit(Handle(changed), new DocumentPlan
        {
            Format = DocFormat.Excel,
            Snapshot = originalSnapshot,
            Operations = new PlanOperation[]
            {
                new SetCellOp
                {
                    Target = new CellAnchor { SheetId = 1U, Address = "A1" },
                    Value = "replacement"
                }
            }
        });

        Assert.False(applied.Committed);
        Assert.Contains(applied.Report.Errors,
            error => error.Code == "stale-snapshot");
    }

    [Fact]
    public async Task Provider_apply_rejects_a_shared_string_change_after_inspection()
    {
        var root = Path.Combine(Path.GetTempPath(), "officeagent-excel-snapshot-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try
        {
            var path = Path.Combine(root, "shared.xlsx");
            await File.WriteAllBytesAsync(path, XlsxFactory.WorkbookWithSharedString("old"));
            var provider = new FileSystemDocumentProvider(new FileSystemDocumentProviderOptions
            {
                ConnectionId = "workspace",
                RootPath = root,
                AllowedExtensions = new[] { ".xlsx" }
            });
            var client = new OfficeAgentClient(new DocumentProviderRegistry(new[] { provider }), new ExcelModule());
            var reference = await client.RegisterAsync("workspace", path);
            var snapshot = (await client.InspectAsync(reference)).Snapshot;
            await File.WriteAllBytesAsync(path, XlsxFactory.WorkbookWithSharedString("new"));
            var changedReference = await client.RegisterAsync("workspace", path);

            var applied = await client.CommitAsync(changedReference, new DocumentPlan
            {
                Format = DocFormat.Excel,
                Snapshot = snapshot,
                Operations = new PlanOperation[]
                {
                    new SetCellOp
                    {
                        Target = new CellAnchor { SheetId = 1U, Address = "A1" },
                        Value = "replacement"
                    }
                }
            });

            Assert.False(applied.Committed);
            Assert.Contains(applied.Report.Errors, error => error.Code == ValidationErrorCodes.StaleSnapshot);
        }
        finally
        {
            if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
        }
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

    [Fact]
    public void Comment_shape_ids_remain_unique_after_delete_and_add()
    {
        var client = new OfficeAgentClient(new ExcelModule());
        var bytes = XlsxFactory.WorkbookWithTable();
        bytes = Apply(bytes, new CommentOp
        {
            Target = new CellAnchor { SheetId = 7U, Address = "A1" },
            Action = CommentAction.Add,
            Author = "Reviewer",
            Text = "First"
        });
        bytes = Apply(bytes, new CommentOp
        {
            Target = new CellAnchor { SheetId = 7U, Address = "B1" },
            Action = CommentAction.Add,
            Author = "Reviewer",
            Text = "Second"
        });
        bytes = Apply(bytes, new CommentOp
        {
            Target = new CellAnchor { SheetId = 7U, Address = "A2" },
            Action = CommentAction.Add,
            Author = "Reviewer",
            Text = "Third"
        });
        bytes = Apply(bytes, new CommentOp
        {
            Target = new NodeAnchor { Kind = "cellComment", Path = "comment#7/A1" },
            Action = CommentAction.Remove
        });
        bytes = Apply(bytes, new CommentOp
        {
            Target = new CellAnchor { SheetId = 7U, Address = "B2" },
            Action = CommentAction.Add,
            Author = "Reviewer",
            Text = "Fourth"
        });

        using var stream = new MemoryStream(bytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        var vml = document.WorkbookPart!.WorksheetParts.Single().VmlDrawingParts.Single();
        using var reader = new StreamReader(vml.GetStream());
        var xml = XDocument.Parse(reader.ReadToEnd());
        var ids = xml.Descendants().Attributes("id")
            .Select(attribute => attribute.Value)
            .Where(value => value.StartsWith("_x0000_s", StringComparison.Ordinal))
            .ToList();
        Assert.Equal(ids.Count, ids.Distinct(StringComparer.Ordinal).Count());
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(document).ToList();
        Assert.True(errors.Count == 0, string.Join("; ", errors.Select(e => e.Description)));
    }

    private static byte[] Apply(byte[] input, PlanOperation operation)
    {
        var client = new OfficeAgentClient(new ExcelModule());
        using var result = client.Commit(Handle(input), new DocumentPlan
        {
            Format = DocFormat.Excel,
            Operations = new[] { operation }
        });
        Assert.True(result.Committed, string.Join("; ", result.Report.Errors.Select(e => e.Message)));
        return result.ToBytes();
    }

    private static StreamHandle Handle(byte[] bytes) => new(new MemoryStream(bytes, writable: false));
}
