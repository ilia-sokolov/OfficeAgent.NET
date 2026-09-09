using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using S = DocumentFormat.OpenXml.Spreadsheet;
using DocFormat = OfficeAgent.Abstractions.DocumentFormat;

namespace OfficeAgent.Excel;

/// <summary>Inspects and edits Open XML workbooks without evaluating formulas.</summary>
public sealed class ExcelModule : IFormatModule, IBlankDocumentFactory, IApplyTimeProvider
{
    /// <summary>Initializes an Excel module using the system clock.</summary>
    public ExcelModule() : this(TimeProvider.System) { }

    /// <summary>Initializes an Excel module with a trusted clock.</summary>
    public ExcelModule(TimeProvider clock)
    {
        Clock = clock ?? throw new ArgumentNullException(nameof(clock));
        Handlers = new IOperationHandler[]
        {
            new SetCellHandler(),
            new AppendTableRowsHandler(),
            new CellCommentHandler()
        };
    }

    /// <inheritdoc />
    public DocFormat Format => DocFormat.Excel;
    /// <inheritdoc />
    public string Extension => ".xlsx";
    /// <inheritdoc />
    public TimeProvider Clock { get; }
    /// <inheritdoc />
    public IReadOnlyList<IOperationHandler> Handlers { get; }
    /// <inheritdoc />
    public bool CanHandle(IOpenXmlPackage package) => package.Format == DocFormat.Excel;
    /// <inheritdoc />
    public byte[] CreateBlank() => ExcelBlankDocument.Create();
    /// <inheritdoc />
    public IReadOnlyDictionary<string, string> Stabilize(IOpenXmlPackage package) =>
        new Dictionary<string, string>(0, StringComparer.Ordinal);

    /// <inheritdoc />
    public InspectResult Inspect(IOpenXmlPackage package, InspectOptions options)
    {
        var document = (SpreadsheetDocument)package.Package;
        var workbookPart = document.WorkbookPart
            ?? throw new InvalidOperationException("Workbook has no workbook part.");
        var worksheets = new List<WorksheetInfo>();
        var cells = new List<CellInfo>();
        var nodes = new List<NodeInfo>();
        var anchors = new List<Anchor>();
        var limit = Math.Max(1, Math.Min(options.MaximumCells, 10000));

        foreach (var sheet in workbookPart.Workbook.Sheets?.Elements<S.Sheet>() ?? Enumerable.Empty<S.Sheet>())
        {
            var sheetId = sheet.SheetId?.Value ?? 0;
            if (options.SheetId is { } selected && selected != sheetId) continue;
            if (sheet.Id?.Value is not { Length: > 0 } rel ||
                workbookPart.GetPartById(rel) is not WorksheetPart part) continue;

            var tables = part.TableDefinitionParts.Select(tablePart => new SpreadsheetTableInfo
            {
                Name = tablePart.Table?.Name?.Value ?? string.Empty,
                DisplayName = tablePart.Table?.DisplayName?.Value ?? string.Empty,
                Reference = tablePart.Table?.Reference?.Value ?? string.Empty
            }).ToList();
            worksheets.Add(new WorksheetInfo
            {
                SheetId = sheetId,
                Name = sheet.Name?.Value ?? string.Empty,
                Dimension = part.Worksheet.SheetDimension?.Reference?.Value,
                Tables = tables
            });

            var sheetPath = $"sheet#{sheetId}";
            nodes.Add(Node("worksheet", sheetPath, sheet.Name?.Value ?? sheetPath));
            foreach (var table in tables)
                nodes.Add(Node("spreadsheetTable", $"table#{sheetId}/{table.DisplayName}",
                    $"{table.DisplayName} ({table.Reference})"));

            var range = options.Range;
            var left = 0;
            var right = 0;
            uint top = 0;
            uint bottom = 0;
            var hasRange = !string.IsNullOrWhiteSpace(range) &&
                SpreadsheetPartUtility.TryParseRange(range!, out left, out top, out right, out bottom);
            if (!string.IsNullOrWhiteSpace(range) && !hasRange)
                throw new ArgumentException($"'{range}' is not a valid A1 range.", nameof(options));

            if (options.Fidelity == Fidelity.Content)
            {
                foreach (var cell in part.Worksheet.Descendants<S.Cell>())
                {
                    if (cells.Count >= limit) break;
                    var address = cell.CellReference?.Value;
                    if (string.IsNullOrEmpty(address) ||
                        !SpreadsheetPartUtility.TryParseCell(address, out var column, out var row)) continue;
                    if (hasRange && (column < left || column > right || row < top || row > bottom)) continue;
                    var anchor = new CellAnchor { Id = $"cell#{sheetId}/{address}", SheetId = sheetId, Address = address };
                    cells.Add(new CellInfo
                    {
                        Anchor = anchor,
                        RawValue = cell.CellValue?.InnerText,
                        DisplayValue = SpreadsheetPartUtility.DisplayValue(document, cell),
                        Formula = cell.CellFormula?.Text
                    });
                    anchors.Add(anchor);
                }
            }

            foreach (var comment in part.WorksheetCommentsPart?.Comments?.CommentList?.Elements<S.Comment>()
                         ?? Enumerable.Empty<S.Comment>())
            {
                var address = comment.Reference?.Value ?? string.Empty;
                nodes.Add(Node("cellComment", $"comment#{sheetId}/{address}",
                    string.Concat(comment.CommentText?.Descendants<S.Text>().Select(t => t.Text)
                        ?? Enumerable.Empty<string>())));
            }
        }

        foreach (var node in nodes)
            if (node.Anchor is not null) anchors.Add(node.Anchor);
        return new InspectResult
        {
            Format = DocFormat.Excel,
            Snapshot = new SnapshotToken(Snapshot(document)),
            Worksheets = worksheets,
            Cells = cells,
            Nodes = nodes,
            Anchors = anchors,
            Styles = new StyleCatalog()
        };
    }

    /// <inheritdoc />
    public IReadOnlyList<FindHit> Find(IOpenXmlPackage package, FindQuery query)
    {
        var document = (SpreadsheetDocument)package.Package;
        var workbook = document.WorkbookPart;
        if (workbook is null || string.IsNullOrEmpty(query.Pattern)) return Array.Empty<FindHit>();
        var hits = new List<FindHit>();
        Regex? regex = null;
        if (query.Options.Regex)
            regex = new Regex(query.Pattern,
                (query.Options.CaseSensitive ? RegexOptions.None : RegexOptions.IgnoreCase) | RegexOptions.CultureInvariant,
                TimeSpan.FromMilliseconds(250));
        var comparison = query.Options.CaseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

        foreach (var sheet in workbook.Workbook.Sheets?.Elements<S.Sheet>() ?? Enumerable.Empty<S.Sheet>())
        {
            var sheetId = sheet.SheetId?.Value ?? 0;
            if (sheet.Id?.Value is not { Length: > 0 } rel || workbook.GetPartById(rel) is not WorksheetPart part) continue;
            foreach (var cell in part.Worksheet.Descendants<S.Cell>())
            {
                var address = cell.CellReference?.Value;
                if (string.IsNullOrEmpty(address)) continue;
                var raw = cell.CellValue?.InnerText ?? string.Empty;
                var displayed = SpreadsheetPartUtility.DisplayValue(document, cell);
                IEnumerable<string> candidates = query.Options.SpreadsheetValueView switch
                {
                    SpreadsheetValueView.Raw => new[] { raw },
                    SpreadsheetValueView.Displayed => new[] { displayed },
                    _ => new[] { displayed, raw }.Distinct(StringComparer.Ordinal)
                };
                var match = candidates.FirstOrDefault(candidate =>
                    regex?.IsMatch(candidate) ?? candidate.IndexOf(query.Pattern, comparison) >= 0);
                if (match is null) continue;
                hits.Add(new FindHit
                {
                    Anchor = new CellAnchor { Id = $"cell#{sheetId}/{address}", SheetId = sheetId, Address = address },
                    Text = match,
                    Context = $"{sheet.Name?.Value ?? sheetId.ToString(CultureInfo.InvariantCulture)}!{address}"
                });
            }
        }
        return hits;
    }

    private static NodeInfo Node(string kind, string path, string summary) => new()
    {
        Kind = kind,
        Path = path,
        Summary = summary,
        Anchor = new NodeAnchor { Id = path, Kind = kind, Path = path }
    };

    private static string Snapshot(SpreadsheetDocument document)
    {
        using var hash = SHA256.Create();
        var parts = new List<string> { document.WorkbookPart?.Workbook.OuterXml ?? string.Empty };
        if (document.WorkbookPart is { } workbook)
            foreach (var sheet in workbook.WorksheetParts)
            {
                parts.Add(sheet.Worksheet.OuterXml);
                parts.AddRange(sheet.TableDefinitionParts.Select(p => p.Table?.OuterXml ?? string.Empty));
                if (sheet.WorksheetCommentsPart?.Comments is { } comments) parts.Add(comments.OuterXml);
            }
        var bytes = Encoding.UTF8.GetBytes(string.Join("\n", parts));
        return Convert.ToBase64String(hash.ComputeHash(bytes));
    }
}
