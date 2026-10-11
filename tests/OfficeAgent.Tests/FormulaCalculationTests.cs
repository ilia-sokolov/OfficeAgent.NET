using System.Globalization;
using System.IO.Compression;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Excel;
using OfficeAgent.Excel.Calculation;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeAgent.Tests;

/// <summary>
/// Optional formula calculation (#47) and formula errors in the report (#49). The evaluator
/// calculates only allowlisted functions it was verified against, reports everything else as
/// not calculated, and the engine writes results into cached values and nothing else.
/// </summary>
public class FormulaCalculationTests
{
    private static readonly ClosedXmlFormulaEvaluator Evaluator = new();

    // ── golden values: one or more cases for every allowlisted function ──────

    /// <summary>
    /// Formula, expected result. Expected values are computed independently - by hand or with
    /// .NET arithmetic - never taken from an engine. Data lives on sheet Data (see <see cref="Golden"/>).
    /// </summary>
    public static TheoryData<string, object> GoldenCases()
    {
        double D(int y, int m, int d) => new DateTime(y, m, d).ToOADate();
        var c1 = new DateTime(2026, 10, 6);
        var c2 = new DateTime(2026, 1, 31);
        var workdays = Enumerable.Range(0, (c1 - c2).Days + 1).Select(i => c2.AddDays(i))
            .Count(d => d.DayOfWeek is not DayOfWeek.Saturday and not DayOfWeek.Sunday);
        var workday5 = c2;
        for (var left = 5; left > 0;) { workday5 = workday5.AddDays(1); if (workday5.DayOfWeek is not DayOfWeek.Saturday and not DayOfWeek.Sunday) left--; }
        var r = 0.05 / 12;

        return new TheoryData<string, object>
        {
            { "SUM(Data!A1:A5)", 15d }, { "AVERAGE(Data!A1:A5)", 3d }, { "MIN(Data!A1:A5)", 1d }, { "MAX(Data!A1:A5)", 5d },
            { "COUNT(Data!A1:A6)", 6d }, { "COUNTA(Data!B1:B4)", 4d }, { "COUNTBLANK(Data!A7:A9)", 3d },
            { "ROUND(2.567,2)", 2.57 }, { "ROUNDUP(2.561,2)", 2.57 }, { "ROUNDDOWN(2.569,2)", 2.56 }, { "INT(-2.5)", -3d },
            { "MOD(10,3)", 1d }, { "ABS(-4)", 4d }, { "POWER(2,10)", 1024d }, { "SQRT(16)", 4d }, { "PRODUCT(Data!A1:A5)", 120d },
            { "SUMPRODUCT(Data!A1:A3,Data!A3:A5)", 26d }, { "SUMIF(Data!A1:A5,\">2\")", 12d },
            { "SUMIFS(Data!F1:F3,Data!E1:E3,\"y\")", 200d }, { "COUNTIF(Data!A1:A5,\">=3\")", 3d },
            { "COUNTIFS(Data!A1:A5,\">1\",Data!A1:A5,\"<5\")", 3d }, { "CEILING(4.3,1)", 5d }, { "FLOOR(4.7,1)", 4d },
            { "CEILING.MATH(4.3)", 5d }, { "FLOOR.MATH(4.7)", 4d }, { "TRUNC(4.79)", 4d }, { "SIGN(-3)", -1d }, { "EXP(0)", 1d },
            { "LN(1)", 0d }, { "LOG10(1000)", 3d }, { "LOG(8,2)", 3d }, { "PI()", Math.PI }, { "GCD(12,18)", 6d }, { "LCM(4,6)", 12d },
            { "FACT(5)", 120d }, { "COMBIN(5,2)", 10d }, { "QUOTIENT(7,2)", 3d }, { "EVEN(3)", 4d }, { "ODD(2)", 3d },
            { "MROUND(10,3)", 9d }, { "SUBTOTAL(9,Data!A1:A5)", 15d },
            { "MEDIAN(Data!A1:A5)", 3d }, { "LARGE(Data!A1:A5,2)", 4d }, { "STDEV.S(Data!A1:A5)", Math.Sqrt(2.5) },
            { "VAR.S(Data!A1:A5)", 2.5 }, { "STDEV.P(Data!A1:A5)", Math.Sqrt(2) },
            { "IF(Data!A1>0,\"pos\",\"neg\")", "pos" }, { "AND(TRUE,FALSE)", false }, { "OR(TRUE,FALSE)", true }, { "NOT(FALSE)", true },
            { "IFERROR(1/0,\"div\")", "div" }, { "ISBLANK(Data!A9)", true }, { "ISNUMBER(Data!A1)", true }, { "ISTEXT(Data!B1)", true },
            { "ISERROR(1/0)", true }, { "ISNA(NA())", true }, { "ISEVEN(4)", true }, { "ISODD(4)", false }, { "N(TRUE)", 1d }, { "TYPE(1)", 1d },
            { "CONCATENATE(\"a\",\"b\")", "ab" }, { "CONCAT(Data!B1,\"-\",Data!B2)", "apple-Banana" },
            { "TEXTJOIN(\",\",TRUE,\"a\",\"\",\"b\")", "a,b" }, { "LEFT(\"hello\",2)", "he" }, { "RIGHT(\"hello\",3)", "llo" },
            { "MID(\"hello\",2,3)", "ell" }, { "LEN(\"hello\")", 5d }, { "UPPER(\"abc\")", "ABC" }, { "LOWER(\"ABC\")", "abc" },
            { "PROPER(\"hello world\")", "Hello World" }, { "TRIM(Data!B3)", "cherry" }, { "SUBSTITUTE(\"a-b-c\",\"-\",\"+\")", "a+b+c" },
            { "REPLACE(\"hello\",1,1,\"J\")", "Jello" }, { "FIND(\"l\",\"hello\")", 3d }, { "SEARCH(\"L\",\"hello\")", 3d },
            { "TEXT(1234.5,\"#,##0.00\")", "1,234.50" }, { "TEXT(Data!C1,\"yyyy-mm-dd\")", "2026-10-06" }, { "VALUE(\"42\")", 42d },
            { "REPT(\"ab\",3)", "ababab" }, { "EXACT(\"a\",\"A\")", false }, { "CHAR(65)", "A" }, { "CODE(\"A\")", 65d },
            { "NUMBERVALUE(\"1.5\")", 1.5 }, { "FIXED(1234.567,1)", "1,234.6" }, { "T(\"x\")", "x" },
            { "VLOOKUP(\"y\",Data!E1:F3,2,FALSE)", 200d }, { "HLOOKUP(\"b\",Data!H6:I7,2,FALSE)", 2d }, { "INDEX(Data!F1:F3,2)", 200d },
            { "MATCH(\"z\",Data!E1:E3,0)", 3d }, { "ROW(Data!A3)", 3d }, { "COLUMN(Data!C1)", 3d }, { "ROWS(Data!A1:A5)", 5d },
            { "COLUMNS(Data!A1:F1)", 6d },
            { "DATE(2026,10,6)", D(2026, 10, 6) }, { "YEAR(Data!C1)", 2026d }, { "MONTH(Data!C1)", 10d }, { "DAY(Data!C1)", 6d },
            { "WEEKDAY(Data!C1)", 3d }, { "EDATE(Data!C2,1)", D(2026, 2, 28) }, { "EOMONTH(Data!C1,0)", D(2026, 10, 31) },
            { "DAYS(Data!C1,Data!C2)", (c1 - c2).TotalDays }, { "NETWORKDAYS(Data!C2,Data!C1)", (double)workdays },
            { "WORKDAY(Data!C2,5)", workday5.ToOADate() }, { "TIME(12,0,0)", 0.5 }, { "HOUR(0.75)", 18d },
            { "MINUTE(TIME(1,30,0))", 30d }, { "YEARFRAC(DATE(2026,1,1),DATE(2026,7,1))", 0.5 },
            { "PMT(0.05/12,60,10000)", -(10000 * r) / (1 - Math.Pow(1 + r, -60)) },
            { "FV(0.05,10,-100)", 100 * (Math.Pow(1.05, 10) - 1) / 0.05 },
            { "1/0", "#DIV/0!" }, { "Data!A1+\"x\"", "#VALUE!" }, { "VLOOKUP(\"q\",Data!E1:F3,2,FALSE)", "#N/A" }
        };
    }

    [Theory]
    [MemberData(nameof(GoldenCases))]
    public void An_allowlisted_function_calculates_its_known_value(string formula, object expected)
    {
        var result = Evaluator.Evaluate(Golden(formula)).Cells.Single(c => c.SheetId == 2 && c.Address == "A1");

        Assert.Equal(FormulaOutcome.Calculated, result.Outcome);
        switch (expected)
        {
            case double number:
                Assert.Equal(FormulaValueKind.Number, result.ValueKind);
                var actual = double.Parse(result.Value!, CultureInfo.InvariantCulture);
                Assert.True(Math.Abs(actual - number) <= 1e-9 * Math.Max(1, Math.Abs(number)), $"={formula}: expected {number}, got {actual}");
                break;
            case bool flag:
                Assert.Equal(FormulaValueKind.Boolean, result.ValueKind);
                Assert.Equal(flag ? "TRUE" : "FALSE", result.Value);
                break;
            case string text when text.StartsWith("#", StringComparison.Ordinal):
                Assert.Equal(FormulaValueKind.Error, result.ValueKind);
                Assert.Equal(text, result.Value);
                break;
            case string text:
                Assert.Equal(FormulaValueKind.Text, result.ValueKind);
                Assert.Equal(text, result.Value);
                break;
        }
    }

    [Fact]
    public void Every_allowlisted_function_has_a_golden_case()
    {
        var used = GoldenCases().Select(row => (string)row[0])
            .SelectMany(formula => System.Text.RegularExpressions.Regex.Matches(formula, @"([A-Z][A-Z0-9.]*)\(").Select(m => m.Groups[1].Value))
            .ToHashSet(StringComparer.Ordinal);

        Assert.Empty(Evaluator.SupportedFunctions.Where(function => !used.Contains(function)));
        Assert.DoesNotContain("NOW", Evaluator.SupportedFunctions);
        Assert.DoesNotContain("XLOOKUP", Evaluator.SupportedFunctions);
    }

    [Fact]
    public void Text_results_do_not_depend_on_the_server_locale()
    {
        var culture = CultureInfo.CurrentCulture;
        CultureInfo.CurrentCulture = new CultureInfo("de-DE");
        try
        {
            var cells = Evaluator.Evaluate(Golden("TEXT(1234.5,\"#,##0.00\")", "NUMBERVALUE(\"1.5\")")).Cells
                .Where(c => c.SheetId == 2).ToDictionary(c => c.Address, c => c.Value);

            Assert.Equal("1,234.50", cells["A1"]);
            Assert.Equal("1.5", cells["B1"]);
            Assert.Equal("de-DE", CultureInfo.CurrentCulture.Name);
        }
        finally
        {
            CultureInfo.CurrentCulture = culture;
        }
    }

    // ── what is not calculated, and why ─────────────────────────────────

    [Theory]
    [InlineData("XLOOKUP(\"z\",Data!E1:E3,Data!F1:F3)", "XLOOKUP")]
    [InlineData("NOW()", "volatile")]
    [InlineData("OFFSET(Data!A1,1,0)", "volatile")]
    [InlineData("SUM(Sales[Amount])", "structured")]
    [InlineData("Missing*2", "does not define")]
    public void A_formula_outside_what_is_verified_is_not_calculated(string formula, string reason)
    {
        var result = Evaluator.Evaluate(Golden(formula)).Cells.Single(c => c.SheetId == 2 && c.Address == "A1");

        Assert.Equal(FormulaOutcome.NotCalculated, result.Outcome);
        Assert.Null(result.Value);
        Assert.Contains(reason, result.Reason);
    }

    [Fact]
    public void A_formula_depending_on_an_uncalculated_cell_is_not_calculated_either()
    {
        var cells = Evaluator.Evaluate(Golden("XLOOKUP(\"z\",Data!E1:E3,Data!F1:F3)", "A1*2", "B1+1", "SUM(Data!A1:A5)")).Cells
            .Where(c => c.SheetId == 2).ToDictionary(c => c.Address);

        Assert.Equal(FormulaOutcome.NotCalculated, cells["B1"].Outcome);
        Assert.Contains("Calc!A1", cells["B1"].Reason);
        Assert.Equal(FormulaOutcome.NotCalculated, cells["C1"].Outcome);
        Assert.Contains("Calc!B1", cells["C1"].Reason);
        Assert.Equal(FormulaOutcome.Calculated, cells["D1"].Outcome);
        Assert.Equal("15", cells["D1"].Value);
    }

    [Fact]
    public void A_structured_reference_never_yields_the_silent_zero()
    {
        // ClosedXML evaluates SUM(Sales[Amount]) to 0 with no error; it must not reach the file.
        var result = Evaluator.Evaluate(Golden("SUM(Sales[Amount])")).Cells.Single(c => c.SheetId == 2 && c.Address == "A1");

        Assert.Equal(FormulaOutcome.NotCalculated, result.Outcome);
    }

    [Fact]
    public void A_circular_reference_leaves_the_whole_workbook_uncalculated()
    {
        var evaluation = Evaluator.Evaluate(Golden("B1+1", "A1+1", "SUM(Data!A1:A5)"));

        Assert.True(evaluation.CircularReference);
        Assert.NotNull(evaluation.NotCalculatedReason);
        Assert.All(evaluation.Cells, cell => Assert.Equal(FormulaOutcome.NotCalculated, cell.Outcome));
    }

    // ── functions newer than the 2007 format ────────────────────────────

    /// <summary>
    /// Excel stores functions added after 2007 with a prefix and reads an unprefixed one as an
    /// unknown name: before 1.2, setCell wrote <c>STDEV.S(…)</c> verbatim and Excel showed #NAME?.
    /// </summary>
    [Theory]
    [InlineData("STDEV.S(B2:B3)", "_xlfn.STDEV.S(B2:B3)")]
    [InlineData("=TEXTJOIN(\",\",TRUE,B2:B3)", "_xlfn.TEXTJOIN(\",\",TRUE,B2:B3)")]
    [InlineData("_xlfn.STDEV.S(B2:B3)", "_xlfn.STDEV.S(B2:B3)")]
    [InlineData("FILTER(B2:B3,B2:B3>8)", "_xlfn._xlws.FILTER(B2:B3,B2:B3>8)")]
    [InlineData("CONCAT(\"STDEV.S(\",B2)", "_xlfn.CONCAT(\"STDEV.S(\",B2)")]
    [InlineData("IFNA(XLOOKUP(1,B2:B3,B2:B3),0)", "_xlfn.IFNA(_xlfn.XLOOKUP(1,B2:B3,B2:B3),0)")]
    [InlineData("SUM(B2:B3)+NETWORKDAYS.INTL(B2,B3)", "SUM(B2:B3)+NETWORKDAYS.INTL(B2,B3)")]
    public void SetCell_stores_a_newer_function_the_way_Excel_reads_it(string formula, string stored)
    {
        using var document = SpreadsheetDocument.Open(new MemoryStream(Formula(formula)), false);
        var cell = document.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<S.Cell>()
            .Single(c => c.CellReference?.Value == "B4");

        Assert.Equal(stored, cell.CellFormula!.Text);
    }

    [Fact]
    public void A_newer_function_stored_without_its_prefix_is_not_calculated()
    {
        // Written by some other tool: Excel shows #NAME? here, so no value may be cached.
        var workbook = Formula("SUM(B2:B3)");
        using (var stream = new MemoryStream())
        {
            stream.Write(workbook, 0, workbook.Length);
            using (var document = SpreadsheetDocument.Open(stream, true))
            {
                document.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<S.Cell>()
                    .Single(c => c.CellReference?.Value == "B4").CellFormula!.Text = "STDEV.S(B2:B3)";
            }
            workbook = stream.ToArray();
        }

        var result = Evaluator.Evaluate(workbook).Cells.Single(c => c.Address == "B4");

        Assert.Equal(FormulaOutcome.NotCalculated, result.Outcome);
        Assert.Contains("_xlfn.", result.Reason);
    }

    [Fact]
    public async Task A_newer_function_written_by_setCell_is_calculated()
    {
        using var applied = await Commit(Formula("STDEV.S(B2:B3)"), calculate: true);

        var cell = applied.Report.Calculation!.Cells.Single(c => c.Address == "B4");
        Assert.Equal(FormulaOutcome.Calculated, cell.Outcome);
        Assert.Equal(Math.Sqrt(12.5), double.Parse(cell.Value!, CultureInfo.InvariantCulture), 9);
    }

    // ── the engine: write-back, report, preview, rejection ──────────────

    [Fact]
    public async Task A_commit_that_asks_for_calculation_writes_the_cached_value()
    {
        using var applied = await Commit(Formula("SUM(B2:B3)"), calculate: true);

        Assert.True(applied.Committed);
        var cell = applied.Report.Calculation!.Cells.Single(c => c.Address == "B4");
        Assert.Equal(("B4", "20"), (cell.Address, cell.Value));
        Assert.StartsWith("ClosedXML", applied.Report.Calculation.Engine);
        Assert.Equal("20", CachedValue(applied.ToBytes(), "B4"));
    }

    [Fact]
    public async Task Without_calculation_the_formula_keeps_no_cached_value()
    {
        using var applied = await Commit(Formula("SUM(B2:B3)"), calculate: false);

        Assert.True(applied.Committed);
        Assert.Null(applied.Report.Calculation);
        Assert.Null(CachedValue(applied.ToBytes(), "B4"));
    }

    [Fact]
    public async Task Calculation_changes_nothing_but_cached_values()
    {
        var input = Formula("SUM(B2:B3)");
        using var plain = await Commit(input, calculate: false);
        using var calculated = await Commit(input, calculate: true);

        var a = Parts(plain.ToBytes());
        var b = Parts(calculated.ToBytes());
        Assert.Equal(a.Keys.OrderBy(k => k), b.Keys.OrderBy(k => k));
        foreach (var (name, bytes) in a)
        {
            if (name.StartsWith("xl/worksheets/", StringComparison.Ordinal))
                Assert.Equal(WithoutCachedValues(bytes), WithoutCachedValues(b[name]));
            else
                Assert.True(bytes.AsSpan().SequenceEqual(b[name]), $"{name} changed");
        }
    }

    [Fact]
    public async Task A_preview_shows_the_result_without_writing()
    {
        var client = Client();
        var input = Formula("SUM(B2:B3)");

        var preview = await client.ApplyAsync(Handle(input), Plan("B4", "SUM(B2:B3)*2"), new ApplyOptions { DryRun = true, CalculateFormulas = true });

        Assert.False(preview.Committed);
        Assert.True(preview.Report.IsValid);
        Assert.Contains(preview.Report.Calculation!.Cells, c => c.Address == "B4" && c.Value == "40");
    }

    [Fact]
    public async Task A_new_formula_error_is_reported_and_still_committed_by_default()
    {
        using var applied = await Commit(Formula("B2/0"), calculate: true);

        Assert.True(applied.Committed);
        var error = Assert.Single(applied.Report.Calculation!.NewErrors);
        Assert.Equal(("B4", "#DIV/0!"), (error.Address, error.Value));
        Assert.True(applied.Report.Calculation.HasFormulaErrors);
    }

    [Fact]
    public async Task Reject_formula_errors_refuses_the_plan_and_writes_nothing()
    {
        var client = Client();

        using var applied = await client.ApplyAsync(Handle(Formula("SUM(B2:B3)")), Plan("B4", "B2/0"),
            new ApplyOptions { DryRun = false, CalculateFormulas = true, RejectFormulaErrors = true });

        Assert.False(applied.Committed);
        Assert.Null(applied.Output);
        var error = Assert.Single(applied.Report.Errors);
        Assert.Equal(ValidationErrorCodes.FormulaError, error.Code);
        Assert.Contains("#DIV/0!", error.Message);
    }

    [Fact]
    public async Task An_error_the_workbook_already_carried_is_not_new()
    {
        var withError = await Commit(Formula("NA()"), calculate: true);
        var input = withError.ToBytes();
        withError.Dispose();

        // A later edit elsewhere in the workbook: the existing #N/A is not the plan's doing.
        using var applied = await Client().ApplyAsync(Handle(input), Plan("C2", "B2*2"),
            new ApplyOptions { DryRun = false, CalculateFormulas = true, RejectFormulaErrors = true });

        Assert.True(applied.Committed, string.Join("; ", applied.Report.Errors.Select(e => e.Message)));
        Assert.Empty(applied.Report.Calculation!.NewErrors);
    }

    [Fact]
    public void Discovery_reports_the_engine_and_its_functions()
    {
        var capabilities = Client().DescribeCapabilities();

        Assert.NotNull(capabilities.FormulaCalculation);
        Assert.StartsWith("ClosedXML", capabilities.FormulaCalculation!.Engine);
        Assert.Contains("SUM", capabilities.FormulaCalculation.SupportedFunctions);
        Assert.DoesNotContain("NOW", capabilities.FormulaCalculation.SupportedFunctions);
    }

    [Fact]
    public void Without_an_evaluator_discovery_reports_no_calculation()
    {
        Assert.Null(new OfficeAgentClient(new ExcelModule()).DescribeCapabilities().FormulaCalculation);
    }

    // ── fixtures ─────────────────────────────────────────────────────────

    private static OfficeAgentClient Client() =>
        new ServiceCollection().AddExcelFormat().AddExcelFormulaCalculation().AddOfficeAgent()
            .BuildServiceProvider().GetRequiredService<OfficeAgentClient>();

    private static StreamHandle Handle(byte[] bytes) => new(new MemoryStream(bytes));

    private static DocumentPlan Plan(string address, string formula) => new()
    {
        Format = OfficeAgent.Abstractions.DocumentFormat.Excel,
        Operations = new PlanOperation[] { new SetCellOp { Target = new CellAnchor { SheetId = 1, Address = address }, Formula = formula } }
    };

    /// <summary>A blank workbook with 12.5 and 7.5 in B2:B3, then the formula committed into B4 without calculation.</summary>
    private static byte[] Formula(string formula)
    {
        var client = new OfficeAgentClient(new ExcelModule());
        using var applied = client.Commit(Handle(new ExcelModule().CreateBlank()), new DocumentPlan
        {
            Format = OfficeAgent.Abstractions.DocumentFormat.Excel,
            Operations = new PlanOperation[]
            {
                new SetCellOp { Target = new CellAnchor { SheetId = 1, Address = "B2" }, Value = "12.5", ValueKind = SpreadsheetCellValueKind.Number },
                new SetCellOp { Target = new CellAnchor { SheetId = 1, Address = "B3" }, Value = "7.5", ValueKind = SpreadsheetCellValueKind.Number },
                new SetCellOp { Target = new CellAnchor { SheetId = 1, Address = "B4" }, Formula = formula }
            }
        });
        Assert.True(applied.Committed);
        return applied.ToBytes();
    }

    private static Task<ApplyResult> Commit(byte[] input, bool calculate) =>
        Client().ApplyAsync(Handle(input), Plan("C2", "B2*2"), new ApplyOptions { DryRun = false, CalculateFormulas = calculate });

    private static string? CachedValue(byte[] workbook, string address)
    {
        using var document = SpreadsheetDocument.Open(new MemoryStream(workbook), false);
        return document.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<S.Cell>()
            .Single(c => c.CellReference?.Value == address).CellValue?.Text;
    }

    /// <summary>
    /// The coverage workbook: data on sheet Data (sheet id 1) as in the spike, a table Sales,
    /// a defined name Rate, and the formulas under test in Calc!A1, B1, C1, … (sheet id 2).
    /// </summary>
    private static byte[] Golden(params string[] formulas)
    {
        using var book = new XLWorkbook();
        var data = book.AddWorksheet("Data");
        for (var i = 1; i <= 5; i++) { data.Cell(i, 1).Value = i; data.Cell(i, 7).Value = i * 2; }
        data.Cell(6, 1).Value = 10.5;
        data.Cell(1, 2).Value = "apple"; data.Cell(2, 2).Value = "Banana"; data.Cell(3, 2).Value = "  cherry  "; data.Cell(4, 2).Value = "a,b,c";
        data.Cell(1, 3).Value = new DateTime(2026, 10, 6); data.Cell(2, 3).Value = new DateTime(2026, 1, 31);
        string[] keys = { "x", "y", "z" };
        for (var i = 0; i < 3; i++) { data.Cell(i + 1, 5).Value = keys[i]; data.Cell(i + 1, 6).Value = (i + 1) * 100; }
        data.Cell("H6").Value = "a"; data.Cell("I6").Value = "b"; data.Cell("H7").Value = 1; data.Cell("I7").Value = 2;
        data.Cell("M1").Value = "Region"; data.Cell("N1").Value = "Amount";
        string[] regions = { "EMEA", "APAC", "AMER" };
        for (var i = 0; i < 3; i++) { data.Cell(i + 2, 13).Value = regions[i]; data.Cell(i + 2, 14).Value = (i + 1) * 10; }
        data.Range("M1:N4").CreateTable("Sales");
        book.DefinedNames.Add("Rate", "0.1");

        var calc = book.AddWorksheet("Calc");
        for (var i = 0; i < formulas.Length; i++) calc.Cell(1, i + 1).FormulaA1 = formulas[i];
        var buffer = new MemoryStream();
        book.SaveAs(buffer);
        return buffer.ToArray();
    }

    private static Dictionary<string, byte[]> Parts(byte[] package)
    {
        using var zip = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        return zip.Entries.ToDictionary(e => e.FullName, e =>
        {
            using var stream = e.Open();
            var copy = new MemoryStream();
            stream.CopyTo(copy);
            return copy.ToArray();
        });
    }

    private static string WithoutCachedValues(byte[] worksheet)
    {
        var xml = System.Xml.Linq.XDocument.Parse(System.Text.Encoding.UTF8.GetString(worksheet));
        System.Xml.Linq.XNamespace m = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        foreach (var c in xml.Descendants(m + "c").Where(c => c.Element(m + "f") is not null))
        {
            c.Element(m + "v")?.Remove();
            c.Attribute("t")?.Remove();
        }
        return xml.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
    }
}
