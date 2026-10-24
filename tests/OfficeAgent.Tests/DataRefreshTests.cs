using System.Security.Cryptography;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.Abstractions;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Excel;
using OfficeAgent.PowerPoint;
using OfficeAgent.Word;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;
using S = DocumentFormat.OpenXml.Spreadsheet;
using DocFormat = OfficeAgent.Abstractions.DocumentFormat;

namespace OfficeAgent.Tests;

/// <summary>
/// Report refresh: a Word report or PowerPoint deck carries markers naming workbook ranges, and
/// one refresh writes this month's values, tables and charts into it.
/// </summary>
public class DataRefreshTests
{
    private static OfficeAgentClient Office() => new(new WordModule(), new PowerPointModule(), new ExcelModule());

    // ── reading the workbook ────────────────────────────────────────────────

    [Fact]
    public void Inspection_shows_cells_as_Excel_displays_them()
    {
        var cells = Office().Inspect(FinanceWorkbook(), new InspectOptions { SheetId = 1, Range = "B2:B6" }).Cells
            .ToDictionary(cell => cell.Anchor.Address, cell => cell.FormattedValue);

        Assert.Equal("$4,200,000", cells["B2"]);
        Assert.Equal("12.50%", cells["B3"]);
        Assert.Equal("10/31/2026", cells["B4"]);
        Assert.Equal("$8,400,000", cells["B5"]);
        Assert.Null(cells["B6"]);
    }

    [Fact]
    public void Inspection_lists_the_workbook_s_defined_names()
    {
        var names = Office().Inspect(FinanceWorkbook(), new InspectOptions { Fidelity = Fidelity.Structure }).DefinedNames;

        Assert.Contains(names, name => name.Name == "Revenue" && name.SheetId == 1 && name.Range == "B2");
        Assert.Contains(names, name => name.Name == "SalesData" && name.SheetId == 2 && name.Range == "A1:C4");
        Assert.DoesNotContain(names, name => name.Name.StartsWith("_xlnm", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData("Summary!B2", "Summary", "B2", null)]
    [InlineData("'Monthly Sales'!$A$1:$C$4", "Monthly Sales", "A1:C4", null)]
    [InlineData("'It''s'!a1", "It's", "A1", null)]
    [InlineData("=Revenue", null, null, "Revenue")]
    public void References_read_as_Excel_writes_them(string text, string? sheet, string? range, string? name)
    {
        Assert.True(WorkbookReference.TryParse(text, out var parsed));
        Assert.Equal(sheet, parsed.Sheet);
        Assert.Equal(range, parsed.Range);
        Assert.Equal(name, parsed.Name);
    }

    [Theory]
    [InlineData("B2")]
    [InlineData("Summary!")]
    [InlineData("Summary!ZZZZ1")]
    [InlineData("two words")]
    public void A_reference_with_no_sheet_or_no_range_is_refused(string text) =>
        Assert.False(WorkbookReference.TryParse(text, out _));

    // ── Word ─────────────────────────────────────────────────────────────────

    [Fact]
    public void A_bound_Word_report_takes_this_month_s_values_and_table_and_can_be_refreshed_again()
    {
        var office = Office();
        var report = BoundReport(office);

        var refresh = office.PlanDataRefresh(report, FinanceWorkbook());
        Assert.True(refresh.IsComplete, string.Join("; ", refresh.Bindings.Select(b => $"{b.Marker}: {b.Message}")));
        Assert.Equal(new[] { DataBindingKind.Value, DataBindingKind.Value, DataBindingKind.Table },
            refresh.Bindings.Select(binding => binding.Kind));
        Assert.Equal("Summary!B2", refresh.Bindings[0].Range);
        Assert.Equal("$4,200,000", refresh.Bindings[0].Summary);

        using var applied = office.Commit(new StreamHandle(new MemoryStream(report)), refresh.Plan);
        Assert.True(applied.Committed, Errors(applied));
        var refreshed = applied.ToBytes();

        var paragraphs = Paragraphs(refreshed);
        Assert.Contains("Revenue reached $4,200,000 this quarter.", paragraphs);
        Assert.Contains("Margin: 12.50%", paragraphs);

        var table = WordTable(refreshed);
        Assert.Equal(new[] { "Month|North|South", "Jan|10|20", "Feb|12|18", "Mar|15|25" }, table.Rows);
        Assert.True(table.HeaderBold, "the header row keeps its bold run");
        Assert.Empty(Validate(refreshed));

        // The markers survive the refresh, so next month is the same call.
        var again = office.PlanDataRefresh(refreshed, FinanceWorkbook());
        Assert.True(again.IsComplete);
        Assert.Equal(3, again.Bindings.Count);
    }

    [Fact]
    public void Binding_the_same_range_twice_gives_the_second_a_suffix_and_both_refresh()
    {
        var office = Office();
        var memo = Memo(Para("Revenue 1.0M."), Para("Again: 1.0M."));
        var inspect = office.Inspect(memo);
        using var bound = office.Commit(new StreamHandle(new MemoryStream(memo)), Plan(
            Bind(inspect.Paragraphs[0].ParaId, "1.0M", "Revenue"),
            Bind(inspect.Paragraphs[1].ParaId, "1.0M", "Revenue")));
        Assert.True(bound.Committed, Errors(bound));

        var tags = office.Inspect(bound.ToBytes()).StructuralAnchors.Select(anchor => anchor.Tag).ToList();
        Assert.Equal(new[] { "xl:Revenue", "xl:Revenue#2" }, tags);

        var refresh = office.PlanDataRefresh(bound.ToBytes(), FinanceWorkbook());
        Assert.True(refresh.IsComplete);
        using var applied = office.Commit(new StreamHandle(new MemoryStream(bound.ToBytes())), refresh.Plan);
        Assert.Equal(new[] { "Revenue $4,200,000.", "Again: $4,200,000." }, Paragraphs(applied.ToBytes()));
    }

    [Theory]
    [InlineData("Summary!B6", DataRefreshErrorCodes.NotCalculated)]
    [InlineData("Nowhere!A1", DataRefreshErrorCodes.RangeNotFound)]
    [InlineData("NoSuchName", DataRefreshErrorCodes.RangeNotFound)]
    [InlineData("Summary!B2:B3", DataRefreshErrorCodes.NotOneCell)]
    public void A_binding_that_cannot_be_read_stops_the_whole_refresh(string source, string code)
    {
        var office = Office();
        var memo = Memo(Para("Revenue 1.0M."), Para("Margin 11.0%."));
        var inspect = office.Inspect(memo);
        using var bound = office.Commit(new StreamHandle(new MemoryStream(memo)), Plan(
            Bind(inspect.Paragraphs[0].ParaId, "1.0M", "Revenue"),
            Bind(inspect.Paragraphs[1].ParaId, "11.0%", source)));
        Assert.True(bound.Committed, Errors(bound));

        var refresh = office.PlanDataRefresh(bound.ToBytes(), FinanceWorkbook());

        Assert.False(refresh.IsComplete);
        Assert.Empty(refresh.Plan.Operations);
        Assert.Null(refresh.Bindings[0].ErrorCode);
        Assert.Equal(code, refresh.Bindings[1].ErrorCode);
        Assert.False(string.IsNullOrEmpty(refresh.Bindings[1].Message));
    }

    [Fact]
    public void A_document_edited_after_planning_is_refused_as_stale()
    {
        var office = Office();
        var report = BoundReport(office);
        var refresh = office.PlanDataRefresh(report, FinanceWorkbook());

        var inspect = office.Inspect(report);
        var closing = inspect.Paragraphs.Single(p => p.Text == "Prepared by Finance.");
        using var edited = office.Commit(new StreamHandle(new MemoryStream(report)), Plan(new ChangeTextOp
        {
            Target = new TextSpanAnchor { ParaId = closing.ParaId, Expect = "Finance" },
            With = "Treasury",
            Mode = ChangeMode.Direct
        }));

        using var applied = office.Commit(new StreamHandle(new MemoryStream(edited.ToBytes())), refresh.Plan);
        Assert.False(applied.Committed);
        Assert.Contains(applied.Report.Errors, error => error.Code == ValidationErrorCodes.StaleSnapshot);
    }

    [Fact]
    public void Fill_table_refuses_rows_that_do_not_match_the_columns()
    {
        var office = Office();
        var report = Memo(Para("x"), SalesTable());

        var report2 = office.Preview(new StreamHandle(new MemoryStream(report)), Plan(new FillTableOp
        {
            Target = new NodeAnchor { Kind = "table", Path = "table#0" },
            Rows = new[] { new[] { "a", "b" } }
        }));

        Assert.Contains("3 column", Assert.Single(report2.Errors).Message);
    }

    // ── PowerPoint ───────────────────────────────────────────────────────────

    [Fact]
    public void A_bound_deck_refreshes_its_value_table_and_chart_and_the_chart_keeps_its_colours()
    {
        var office = Office();
        var deck = BoundDeck(office);

        var refresh = office.PlanDataRefresh(deck, FinanceWorkbook());
        Assert.True(refresh.IsComplete, string.Join("; ", refresh.Bindings.Select(b => $"{b.Marker}: {b.Message}")));
        Assert.Equal(new[] { DataBindingKind.Value, DataBindingKind.Table, DataBindingKind.Chart },
            refresh.Bindings.Select(binding => binding.Kind).OrderBy(kind => kind));

        using var applied = office.Commit(new StreamHandle(new MemoryStream(deck)), new DocumentPlan
        {
            Format = DocFormat.PowerPoint,
            Snapshot = refresh.Plan.Snapshot,
            Operations = refresh.Plan.Operations
        });
        Assert.True(applied.Committed, Errors(applied));

        using var stream = new MemoryStream(applied.ToBytes());
        using var document = PresentationDocument.Open(stream, isEditable: false);
        var slide = document.PresentationPart!.SlideParts.First();

        var title = slide.Slide.Descendants<P.Shape>().First(shape =>
            shape.NonVisualShapeProperties!.NonVisualDrawingProperties!.Name!.Value!.StartsWith("xl:", StringComparison.Ordinal));
        Assert.Equal("$4,200,000", title.TextBody!.InnerText);

        var rows = slide.Slide.Descendants<A.Table>().Single().Elements<A.TableRow>()
            .Select(row => string.Join("|", row.Elements<A.TableCell>().Select(cell => cell.TextBody!.InnerText)))
            .ToList();
        Assert.Equal(new[] { "Metric|Value", "Revenue|$4,200,000", "Margin|12.50%", "Closed|10/31/2026", "Total|$8,400,000" }, rows);

        var chart = slide.ChartParts.Single();
        var series = chart.ChartSpace!.Descendants<C.BarChartSeries>().ToList();
        Assert.Equal(2, series.Count);
        Assert.Equal(new[] { "North", "South" },
            series.Select(s => s.GetFirstChild<C.SeriesText>()!.Descendants<C.StringPoint>().Single().InnerText));
        Assert.Equal(new[] { "Jan", "Feb", "Mar" },
            series[0].GetFirstChild<C.CategoryAxisData>()!.Descendants<C.StringPoint>().Select(p => p.InnerText));
        Assert.Equal(new[] { "10", "12", "15" },
            series[0].GetFirstChild<C.Values>()!.Descendants<C.NumericPoint>().Select(p => p.InnerText));
        Assert.Equal(new[] { "20", "18", "25" },
            series[1].GetFirstChild<C.Values>()!.Descendants<C.NumericPoint>().Select(p => p.InnerText));
        Assert.All(series, s => Assert.Equal("C00000",
            s.GetFirstChild<C.ChartShapeProperties>()?.Descendants<A.RgbColorModelHex>().First().Val?.Value));
        Assert.Equal(new uint[] { 0, 1 }, series.Select(s => s.GetFirstChild<C.Index>()!.Val!.Value));

        // "Edit Data" in PowerPoint opens the numbers the chart shows.
        var embedded = chart.EmbeddedPackagePart!;
        using var workbookStream = embedded.GetStream();
        using var workbook = SpreadsheetDocument.Open(workbookStream, isEditable: false);
        var sheetText = workbook.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<S.Cell>()
            .ToDictionary(cell => cell.CellReference!.Value!, cell => cell.InnerText);
        Assert.Equal("Mar", sheetText["A4"]);
        Assert.Equal("25", sheetText["C4"]);

        Assert.DoesNotContain(new OpenXmlValidator(FileFormatVersions.Office2019).Validate(document),
            error => error.Part?.Uri.ToString().Contains("/slides/", StringComparison.Ordinal) == true ||
                     error.Part?.Uri.ToString().Contains("/charts/", StringComparison.Ordinal) == true);
    }

    [Fact]
    public void Chart_data_with_fewer_series_drops_the_extra_ones()
    {
        var office = Office();
        var deck = DeckWithChart(office, seriesCount: 3);
        var path = office.Inspect(deck).Nodes.Single(node => node.Kind == "chart").Path;

        using var applied = office.Commit(new StreamHandle(new MemoryStream(deck)), new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new SetChartDataOp
                {
                    Target = new NodeAnchor { Kind = "chart", Path = path },
                    Categories = new[] { "Q1" },
                    Series = new[] { new ChartSeries { Name = "Only", Values = new double?[] { 5 } } }
                }
            }
        });

        Assert.True(applied.Committed, Errors(applied));
        using var stream = new MemoryStream(applied.ToBytes());
        using var document = PresentationDocument.Open(stream, isEditable: false);
        Assert.Single(document.PresentationPart!.SlideParts.First().ChartParts.Single()
            .ChartSpace!.Descendants<C.BarChartSeries>());
    }

    [Fact]
    public void Accounting_and_currency_formats_display_as_Excel_shows_them()
    {
        var book = WorkbookWithFormats(("A1", "4200000.5", 44U), ("A2", "-1234", 7U), ("A3", "1234", 42U));
        var cells = Office().Inspect(book, new InspectOptions { SheetId = 1, Range = "A1:A3" }).Cells
            .ToDictionary(cell => cell.Anchor.Address, cell => cell.FormattedValue);

        Assert.Equal("$4,200,000.50", cells["A1"]);
        Assert.Equal("($1,234.00)", cells["A2"]);
        Assert.Equal("$1,234", cells["A3"]);
    }

    [Fact]
    public void Growing_a_slide_table_does_not_copy_its_row_id()
    {
        var office = Office();
        using var stream = new MemoryStream();
        stream.Write(PptxFactory.DeckWithTable());
        using (var document = PresentationDocument.Open(stream, isEditable: true))
        {
            var slide = document.PresentationPart!.SlideParts.First();
            foreach (var (row, id) in slide.Slide.Descendants<A.TableRow>().Select((row, i) => (row, i)))
                row.Append(new A.ExtensionList(new A.Extension { Uri = "{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}" }));
            slide.Slide.Save();
        }

        var deck = stream.ToArray();
        var path = office.Inspect(deck).Nodes.Single(node => node.Kind == "table").Path;
        using var applied = office.Commit(new StreamHandle(new MemoryStream(deck)), Plan(new FillTableOp
        {
            Target = new NodeAnchor { Kind = "table", Path = path },
            Rows = new[] { new[] { "Region", "Q1" }, new[] { "North", "1" }, new[] { "South", "2" }, new[] { "West", "3" } }
        }));

        Assert.True(applied.Committed, Errors(applied));
        using var output = new MemoryStream(applied.ToBytes());
        using var result = PresentationDocument.Open(output, isEditable: false);
        var extensions = result.PresentationPart!.SlideParts.First().Slide.Descendants<A.TableRow>()
            .Select(row => row.Elements<A.ExtensionList>().Count()).ToList();
        Assert.Equal(new[] { 1, 1, 0, 0 }, extensions);
    }

    [Fact]
    public void Chart_data_that_would_empty_one_type_of_a_combination_chart_is_refused()
    {
        var office = Office();
        var deck = DeckWithChart(office, seriesCount: 2);
        using var stream = new MemoryStream();
        stream.Write(deck);
        using (var document = PresentationDocument.Open(stream, isEditable: true))
        {
            // Move the second column series into a line group: columns and a line on one chart.
            var chart = document.PresentationPart!.SlideParts.First().ChartParts.Single();
            var plot = chart.ChartSpace!.Descendants<C.PlotArea>().Single();
            var bar = plot.GetFirstChild<C.BarChart>()!;
            var second = bar.Elements<C.BarChartSeries>().Last();
            second.Remove();
            var line = new C.LineChart(new C.Grouping { Val = C.GroupingValues.Standard }, new C.VaryColors { Val = false },
                new C.LineChartSeries(second.Elements().Where(e => e is C.Index or C.Order or C.SeriesText or C.CategoryAxisData or C.Values)
                    .Select(e => e.CloneNode(true))),
                new C.Marker(), bar.Elements<C.AxisId>().First().CloneNode(true), bar.Elements<C.AxisId>().Last().CloneNode(true));
            bar.InsertAfterSelf(line);
            chart.ChartSpace.Save();
        }

        var combined = stream.ToArray();
        var path = office.Inspect(combined).Nodes.Single(node => node.Kind == "chart").Path;
        var report = office.Preview(new StreamHandle(new MemoryStream(combined)), Plan(new SetChartDataOp
        {
            Target = new NodeAnchor { Kind = "chart", Path = path },
            Categories = new[] { "Q1" },
            Series = new[] { new ChartSeries { Name = "Only", Values = new double?[] { 1 } } }
        }));

        Assert.Contains("combines chart types", Assert.Single(report.Errors).Message);
    }

    [Fact]
    public void Filling_a_Word_table_keeps_a_comment_on_its_cell_and_does_not_copy_it_into_new_rows()
    {
        var office = Office();
        var commented = new TableCell(new Paragraph(
            new CommentRangeStart { Id = "0" }, new Run(new Text("9")), new CommentRangeEnd { Id = "0" },
            new Run(new CommentReference { Id = "0" })));
        var table = new Table(
            new TableProperties(new TableWidth { Type = TableWidthUnitValues.Auto, Width = "0" }),
            new TableGrid(new GridColumn { Width = "2000" }, new GridColumn { Width = "2000" }),
            new TableRow(new TableCell(new Paragraph(new Run(new Text("Month")))), new TableCell(new Paragraph(new Run(new Text("North"))))),
            new TableRow(new TableCell(new Paragraph(new Run(new Text("Dec")))), commented));
        var report = Memo(Para("x"), table);

        using var applied = office.Commit(new StreamHandle(new MemoryStream(report)), Plan(new FillTableOp
        {
            Target = new NodeAnchor { Kind = "table", Path = "table#0" },
            Rows = new[] { new[] { "Month", "North" }, new[] { "Jan", "10" }, new[] { "Feb", "12" } }
        }));

        Assert.True(applied.Committed, Errors(applied));
        using var stream = new MemoryStream(applied.ToBytes());
        using var document = WordprocessingDocument.Open(stream, isEditable: false);
        var rows = document.MainDocumentPart!.Document.Body!.Elements<Table>().Single().Elements<TableRow>().ToList();
        Assert.Equal("10", rows[1].Elements<TableCell>().Last().InnerText);
        Assert.Single(rows[1].Descendants<CommentReference>());
        Assert.Single(rows[1].Descendants<CommentRangeStart>());
        Assert.Empty(rows[2].Descendants<CommentReference>());
        Assert.Empty(rows[2].Descendants<CommentRangeStart>());
    }

    [Fact]
    public void A_bound_OfficeAgent_chart_can_still_be_rebuilt_with_update_chart()
    {
        var office = Office();
        var deck = BoundDeck(office);
        var chart = office.Inspect(deck).Nodes.Single(node => node.Kind == "chart");
        Assert.StartsWith("xl:SalesData#OfficeAgent Chart ", chart.Name);

        using var applied = office.Commit(new StreamHandle(new MemoryStream(deck)), Plan(new UpdateChartOp
        {
            Target = new NodeAnchor { Kind = "chart", Path = chart.Path },
            Kind = ChartKind.Line,
            Categories = new[] { "A" },
            Series = new[] { new ChartSeries { Name = "S", Values = new double?[] { 1 } } }
        }));

        Assert.True(applied.Committed, Errors(applied));
    }

    // ── refresh_document ─────────────────────────────────────────────────────

    [Fact]
    public async Task Refresh_document_previews_then_saves_with_a_receipt_and_lists_its_bindings()
    {
        using var workspace = new Workspace();
        workspace.Stage("report.docx", BoundReport(Office()));
        workspace.Stage("finance.xlsx", FinanceWorkbook());
        var before = workspace.Hash("report.docx");

        var preview = Parse(await workspace.Tools.RefreshDocument("workspace", "report.docx", "finance.xlsx", preview: true));
        Assert.True(preview.GetProperty("isValid").GetBoolean(), preview.GetRawText());
        Assert.False(preview.GetProperty("committed").GetBoolean());
        Assert.Equal(3, preview.GetProperty("bindings").GetArrayLength());
        Assert.Equal("Value", preview.GetProperty("bindings")[0].GetProperty("kind").GetString());
        Assert.Equal(before, workspace.Hash("report.docx"));

        var saved = Parse(await workspace.Tools.RefreshDocument("workspace", "report.docx", "finance.xlsx"));
        Assert.True(saved.GetProperty("committed").GetBoolean(), saved.GetRawText());
        Assert.False(string.IsNullOrEmpty(saved.GetProperty("receipt").GetProperty("outputSha256").GetString()));
        Assert.Contains("Revenue reached $4,200,000 this quarter.", Paragraphs(workspace.Read("report.docx")));
    }

    [Fact]
    public async Task Refresh_document_without_bindings_says_how_to_add_them_and_writes_nothing()
    {
        using var workspace = new Workspace();
        workspace.Stage("plain.docx", Memo(Para("Nothing linked here.")));
        workspace.Stage("finance.xlsx", FinanceWorkbook());
        var before = workspace.Hash("plain.docx");

        var response = Parse(await workspace.Tools.RefreshDocument("workspace", "plain.docx", "finance.xlsx"));

        var error = response.GetProperty("errors")[0];
        Assert.Equal(DataRefreshErrorCodes.NoBindings, error.GetProperty("code").GetString());
        Assert.Contains("bindData", error.GetProperty("message").GetString());
        Assert.Equal(before, workspace.Hash("plain.docx"));
    }

    [Fact]
    public async Task Refresh_document_reports_every_unreadable_binding_and_writes_nothing()
    {
        using var workspace = new Workspace();
        var office = Office();
        var memo = Memo(Para("Pending 1.0M."));
        using var bound = office.Commit(new StreamHandle(new MemoryStream(memo)),
            Plan(Bind(office.Inspect(memo).Paragraphs[0].ParaId, "1.0M", "Summary!B6")));
        workspace.Stage("report.docx", bound.ToBytes());
        workspace.Stage("finance.xlsx", FinanceWorkbook());
        var before = workspace.Hash("report.docx");

        var response = Parse(await workspace.Tools.RefreshDocument("workspace", "report.docx", "finance.xlsx"));

        Assert.False(response.GetProperty("committed").GetBoolean());
        Assert.Equal(DataRefreshErrorCodes.NotCalculated, response.GetProperty("errors")[0].GetProperty("code").GetString());
        Assert.Equal(DataRefreshErrorCodes.NotCalculated, response.GetProperty("bindings")[0].GetProperty("errorCode").GetString());
        Assert.Equal(before, workspace.Hash("report.docx"));
    }

    // ── fixtures ─────────────────────────────────────────────────────────────

    /// <summary>A report with two numbers and a table, each bound to the finance workbook.</summary>
    private static byte[] BoundReport(OfficeAgentClient office)
    {
        var report = Memo(
            Para("Revenue reached 3.1M this quarter."),
            Para("Margin: 11.0%"),
            SalesTable(),
            Para("Prepared by Finance."));
        var inspect = office.Inspect(report);
        using var bound = office.Commit(new StreamHandle(new MemoryStream(report)), Plan(
            Bind(inspect.Paragraphs.First(p => p.Text.StartsWith("Revenue", StringComparison.Ordinal)).ParaId, "3.1M", "Revenue"),
            Bind(inspect.Paragraphs.First(p => p.Text.StartsWith("Margin", StringComparison.Ordinal)).ParaId, "11.0%", "Summary!B3"),
            new BindDataOp { Target = new NodeAnchor { Kind = "table", Path = "table#0" }, Source = "SalesData" }));
        Assert.True(bound.Committed, Errors(bound));
        return bound.ToBytes();
    }

    /// <summary>A deck whose title shape, table and red-coloured chart are bound to the workbook.</summary>
    private static byte[] BoundDeck(OfficeAgentClient office)
    {
        var deck = DeckWithChart(office, seriesCount: 1);
        var nodes = office.Inspect(deck).Nodes;
        var chart = nodes.Single(node => node.Kind == "chart").Path;
        var table = nodes.Single(node => node.Kind == "table").Path;
        using var bound = office.Commit(new StreamHandle(new MemoryStream(deck)), new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new BindDataOp { Target = new NodeAnchor { Kind = "shape", Path = "shape#256/2" }, Source = "Revenue" },
                new BindDataOp { Target = new NodeAnchor { Kind = "table", Path = table }, Source = "Summary!A1:B5" },
                new BindDataOp { Target = new NodeAnchor { Kind = "chart", Path = chart }, Source = "SalesData" }
            }
        });
        Assert.True(bound.Committed, Errors(bound));
        return bound.ToBytes();
    }

    /// <summary>The table deck plus a native chart whose series are filled red, as an author would style one.</summary>
    private static byte[] DeckWithChart(OfficeAgentClient office, int seriesCount)
    {
        var deck = PptxFactory.DeckWithTable();
        using var withChart = office.Commit(new StreamHandle(new MemoryStream(deck)), new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new InsertChartOp
                {
                    Target = new NodeAnchor { Kind = "slide", Path = "slide#256" },
                    Kind = ChartKind.ClusteredColumn,
                    Categories = new[] { "Oct", "Nov" },
                    Series = Enumerable.Range(1, seriesCount)
                        .Select(i => new ChartSeries { Name = $"Old {i}", Values = new double?[] { 1, 2 } }).ToArray(),
                    WidthPx = 400, HeightPx = 300
                }
            }
        });
        Assert.True(withChart.Committed, Errors(withChart));

        using var stream = new MemoryStream();
        stream.Write(withChart.ToBytes());
        using (var document = PresentationDocument.Open(stream, isEditable: true))
        {
            var chart = document.PresentationPart!.SlideParts.First().ChartParts.Single();
            foreach (var series in chart.ChartSpace!.Descendants<C.BarChartSeries>())
                series.SeriesText!.InsertAfterSelf(new C.ChartShapeProperties(
                    new A.SolidFill(new A.RgbColorModelHex { Val = "C00000" })));
            chart.ChartSpace.Save();
        }

        return stream.ToArray();
    }

    private static Table SalesTable()
    {
        static TableCell Cell(string text, bool bold = false)
        {
            var run = new Run();
            if (bold) run.AppendChild(new RunProperties(new Bold()));
            run.AppendChild(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
            return new TableCell(new Paragraph(run));
        }

        return new Table(
            new TableProperties(new TableWidth { Type = TableWidthUnitValues.Auto, Width = "0" }),
            new TableGrid(new GridColumn { Width = "2000" }, new GridColumn { Width = "2000" }, new GridColumn { Width = "2000" }),
            new TableRow(Cell("Month", bold: true), Cell("North", bold: true), Cell("South", bold: true)),
            new TableRow(Cell("Dec"), Cell("9"), Cell("19")));
    }

    /// <summary>
    /// A finance workbook: Summary holds formatted figures, one formula with a stored result and
    /// one without; 'Monthly Sales' holds a small table. Two defined names point into them.
    /// </summary>
    private static byte[] FinanceWorkbook()
    {
        using var stream = new MemoryStream();
        using (var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var workbook = document.AddWorkbookPart();
            workbook.Workbook = new S.Workbook();
            var styles = workbook.AddNewPart<WorkbookStylesPart>();
            styles.Stylesheet = new S.Stylesheet(
                new S.NumberingFormats(new S.NumberingFormat { NumberFormatId = 164U, FormatCode = "\"$\"#,##0" }) { Count = 1 },
                new S.Fonts(new S.Font()) { Count = 1 },
                new S.Fills(new S.Fill(new S.PatternFill { PatternType = S.PatternValues.None }),
                    new S.Fill(new S.PatternFill { PatternType = S.PatternValues.Gray125 })) { Count = 2 },
                new S.Borders(new S.Border()) { Count = 1 },
                new S.CellStyleFormats(new S.CellFormat()) { Count = 1 },
                new S.CellFormats(
                    new S.CellFormat(),
                    new S.CellFormat { NumberFormatId = 164U, ApplyNumberFormat = true },
                    new S.CellFormat { NumberFormatId = 10U, ApplyNumberFormat = true },
                    new S.CellFormat { NumberFormatId = 14U, ApplyNumberFormat = true }) { Count = 4 });

            var summary = workbook.AddNewPart<WorksheetPart>();
            summary.Worksheet = new S.Worksheet(new S.SheetData(
                Row(1, Inline("A1", "Metric"), Inline("B1", "Value")),
                Row(2, Inline("A2", "Revenue"), Number("B2", "4200000", 1)),
                Row(3, Inline("A3", "Margin"), Number("B3", "0.125", 2)),
                Row(4, Inline("A4", "Closed"), Number("B4", "46326", 3)),
                Row(5, Inline("A5", "Total"), new S.Cell(new S.CellFormula("B2*2"), new S.CellValue("8400000")) { CellReference = "B5", StyleIndex = 1U }),
                Row(6, Inline("A6", "Pending"), new S.Cell(new S.CellFormula("B2*3")) { CellReference = "B6", StyleIndex = 1U })));

            var sales = workbook.AddNewPart<WorksheetPart>();
            sales.Worksheet = new S.Worksheet(new S.SheetData(
                Row(1, Inline("A1", "Month"), Inline("B1", "North"), Inline("C1", "South")),
                Row(2, Inline("A2", "Jan"), Number("B2", "10"), Number("C2", "20")),
                Row(3, Inline("A3", "Feb"), Number("B3", "12"), Number("C3", "18")),
                Row(4, Inline("A4", "Mar"), Number("B4", "15"), Number("C4", "25"))));

            workbook.Workbook.Append(
                new S.Sheets(
                    new S.Sheet { Id = workbook.GetIdOfPart(summary), SheetId = 1U, Name = "Summary" },
                    new S.Sheet { Id = workbook.GetIdOfPart(sales), SheetId = 2U, Name = "Monthly Sales" }),
                new S.DefinedNames(
                    new S.DefinedName("Summary!$B$2") { Name = "Revenue" },
                    new S.DefinedName("'Monthly Sales'!$A$1:$C$4") { Name = "SalesData" },
                    new S.DefinedName("Summary!$A$1:$B$6") { Name = "_xlnm.Print_Area", LocalSheetId = 0U }));
            styles.Stylesheet.Save();
            summary.Worksheet.Save();
            sales.Worksheet.Save();
            workbook.Workbook.Save();
        }

        return stream.ToArray();
    }

    /// <summary>One sheet of numbers, each cell styled with a built-in number format id and no format code.</summary>
    private static byte[] WorkbookWithFormats(params (string Address, string Value, uint FormatId)[] cells)
    {
        using var stream = new MemoryStream();
        using (var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var workbook = document.AddWorkbookPart();
            workbook.Workbook = new S.Workbook();
            var styles = workbook.AddNewPart<WorkbookStylesPart>();
            styles.Stylesheet = new S.Stylesheet(
                new S.Fonts(new S.Font()) { Count = 1 },
                new S.Fills(new S.Fill(new S.PatternFill { PatternType = S.PatternValues.None }),
                    new S.Fill(new S.PatternFill { PatternType = S.PatternValues.Gray125 })) { Count = 2 },
                new S.Borders(new S.Border()) { Count = 1 },
                new S.CellStyleFormats(new S.CellFormat()) { Count = 1 },
                new S.CellFormats(new[] { new S.CellFormat() }.Concat(cells.Select(cell =>
                    new S.CellFormat { NumberFormatId = cell.FormatId, ApplyNumberFormat = true }))) { Count = (uint)cells.Length + 1 });
            var sheet = workbook.AddNewPart<WorksheetPart>();
            sheet.Worksheet = new S.Worksheet(new S.SheetData(cells.Select((cell, i) =>
                Row((uint)(i + 1), Number(cell.Address, cell.Value, (uint)(i + 1))))));
            workbook.Workbook.Append(new S.Sheets(new S.Sheet { Id = workbook.GetIdOfPart(sheet), SheetId = 1U, Name = "Data" }));
            styles.Stylesheet.Save();
            sheet.Worksheet.Save();
            workbook.Workbook.Save();
        }

        return stream.ToArray();
    }

    private static S.Row Row(uint index, params S.Cell[] cells) => new(cells) { RowIndex = index };

    private static S.Cell Inline(string address, string text) =>
        new(new S.InlineString(new S.Text(text))) { CellReference = address, DataType = S.CellValues.InlineString };

    private static S.Cell Number(string address, string value, uint style = 0) =>
        new(new S.CellValue(value)) { CellReference = address, StyleIndex = style == 0 ? null : style };

    private static BindDataOp Bind(string paraId, string expect, string source) => new()
    {
        Target = new TextSpanAnchor { ParaId = paraId, Expect = expect },
        Source = source
    };

    private static DocumentPlan Plan(params PlanOperation[] operations) => new() { Operations = operations };

    private static Paragraph Para(string text) =>
        new(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));

    private static byte[] Memo(params OpenXmlElement[] blocks)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(blocks));
            main.Document.Save();
        }

        return stream.ToArray();
    }

    private static IReadOnlyList<string> Paragraphs(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, isEditable: false);
        return document.MainDocumentPart!.Document.Body!.Elements<Paragraph>()
            .Select(p => string.Concat(p.Descendants<Text>().Select(t => t.Text)))
            .ToList();
    }

    private static (IReadOnlyList<string> Rows, bool HeaderBold) WordTable(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, isEditable: false);
        var table = document.MainDocumentPart!.Document.Body!.Elements<Table>().Single();
        var rows = table.Elements<TableRow>()
            .Select(row => string.Join("|", row.Elements<TableCell>().Select(cell => cell.InnerText)))
            .ToList();
        var bold = table.Elements<TableRow>().First().Descendants<Run>().All(run => run.RunProperties?.Bold is not null);
        return (rows, bold);
    }

    private static IReadOnlyList<ValidationErrorInfo> Validate(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, isEditable: false);
        return new OpenXmlValidator(FileFormatVersions.Office2019).Validate(document).ToList();
    }

    private static string Errors(ApplyResult result) =>
        string.Join("; ", result.Report.Errors.Select(e => $"{e.Code}: {e.Message}"));

    private static JsonElement Parse(string json) => JsonDocument.Parse(json).RootElement.Clone();

    private sealed class Workspace : IDisposable
    {
        private readonly ServiceProvider _services;

        public Workspace()
        {
            Root = Path.Combine(Path.GetTempPath(), $"officeagent-refresh-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Root);
            var services = new ServiceCollection();
            services.AddWordFormat();
            services.AddPowerPointFormat();
            services.AddExcelFormat();
            services.AddFileSystemDocumentProvider("workspace", Root, options =>
                options.AllowedExtensions = new[] { ".docx", ".pptx", ".xlsx" });
            services.AddOfficeAgent();
            _services = services.BuildServiceProvider();
            Tools = new OfficeAgentTools(_services.GetRequiredService<OfficeAgentClient>());
        }

        public string Root { get; }

        public OfficeAgentTools Tools { get; }

        public void Stage(string name, byte[] bytes) => File.WriteAllBytes(Path.Combine(Root, name), bytes);

        public byte[] Read(string name) => File.ReadAllBytes(Path.Combine(Root, name));

        public string Hash(string name) => Convert.ToHexString(SHA256.HashData(Read(name)));

        public void Dispose()
        {
            _services.Dispose();
            if (Directory.Exists(Root)) Directory.Delete(Root, recursive: true);
        }
    }
}
