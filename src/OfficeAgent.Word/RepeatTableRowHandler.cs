using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;

namespace OfficeAgent.Word;

/// <summary>Expands one explicitly selected Word table row from named placeholders.</summary>
internal sealed class RepeatTableRowHandler : IOperationHandler
{
    private static readonly Regex Placeholder = new(
        @"\{\{(?<name>[A-Za-z_][A-Za-z0-9_.-]*)\}\}",
        RegexOptions.CultureInvariant,
        TimeSpan.FromSeconds(1));
    private readonly TableNodeProvider _tables = new();

    public bool CanHandle(PlanOperation operation) =>
        operation is RepeatTableRowOp { Target: NodeAnchor { Kind: "table" } };

    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (RepeatTableRowOp)operation;
        var anchor = (NodeAnchor)op.Target;
        var row = ResolveTemplate(context, anchor, op.TemplateRowIndex, out var error);
        if (row is null) return OperationPreview.Fail(error!);
        if (row.Descendants<OpenXmlElement>().Any(element => WordRevisions.TagOf(element) is not null))
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "The selected template row contains tracked revisions; resolve them before repeating the row.", anchor));

        var names = PlaceholderNames(row);
        if (names.Count == 0)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "The selected template row contains no {{Field}} placeholders.", anchor));

        for (var recordIndex = 0; recordIndex < op.Records.Count; recordIndex++)
        {
            var record = op.Records[recordIndex];
            var missing = names.Where(name => !record.ContainsKey(name)).ToArray();
            if (missing.Length > 0 && op.MissingValueBehavior == MissingTemplateValueBehavior.Fail)
                return OperationPreview.Fail(new ValidationError(
                    ValidationErrorCodes.InvalidOperation,
                    $"Repeating record {recordIndex} is missing: {string.Join(", ", missing)}.", anchor));
            var unknown = record.Keys.Where(key => !names.Contains(key)).ToArray();
            if (unknown.Length > 0 && op.RejectUnknownValues)
                return OperationPreview.Fail(new ValidationError(
                    ValidationErrorCodes.InvalidOperation,
                    $"Repeating record {recordIndex} supplies fields absent from the row: {string.Join(", ", unknown)}.", anchor));
        }

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "repeatTableRow",
            Before = $"template row {op.TemplateRowIndex}",
            After = $"{op.Records.Count} populated row(s)",
            Context = anchor.Path,
            BlastRadius = Math.Max(1, op.Records.Count)
        });
    }

    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (RepeatTableRowOp)operation;
        var anchor = (NodeAnchor)op.Target;
        var template = ResolveTemplate(context, anchor, op.TemplateRowIndex, out _)
            ?? throw new InvalidOperationException("The repeating template row vanished before apply.");
        var names = PlaceholderNames(template);
        var rows = op.Records.Select(record => BuildRow(context, template, names, record, op.MissingValueBehavior)).ToList();

        OpenXmlElement cursor = template;
        foreach (var row in rows)
        {
            cursor.InsertAfterSelf(row);
            cursor = row;
        }

        if (WordRevisionMarker.IsTracked(op.Mode))
        {
            var marker = new WordRevisionMarker(context.Package, context.Revision);
            marker.MarkRowDeleted(template);
            foreach (var row in rows) marker.MarkRowInserted(row);
        }
        else
        {
            template.Remove();
        }
    }

    private TableRow? ResolveTemplate(
        ApplyContext context,
        NodeAnchor anchor,
        int index,
        out ValidationError? error)
    {
        var node = _tables.Resolve(anchor, new WordObjectMap(context.Package));
        if (node is null)
        {
            error = new ValidationError(ValidationErrorCodes.AnchorNotFound,
                $"No table at path '{anchor.Path}'.", anchor);
            return null;
        }
        var rows = ((Table)node.Elements[0]).Elements<TableRow>().ToList();
        if (index < 0 || index >= rows.Count)
        {
            error = new ValidationError(ValidationErrorCodes.InvalidOperation,
                $"Template row {index} is outside the table's 0..{Math.Max(0, rows.Count - 1)} range.", anchor);
            return null;
        }
        error = null;
        return rows[index];
    }

    private static HashSet<string> PlaceholderNames(TableRow row) => new(
        row.Descendants<Paragraph>()
            .SelectMany(p => Placeholder.Matches(WordModel.Text.GetLogicalText(p)).Cast<Match>())
            .Select(match => match.Groups["name"].Value),
        StringComparer.Ordinal);

    private static TableRow BuildRow(
        ApplyContext context,
        TableRow template,
        HashSet<string> names,
        IReadOnlyDictionary<string, string?> record,
        MissingTemplateValueBehavior missingBehavior)
    {
        var row = (TableRow)template.CloneNode(true);
        ClearParagraphIds(row);
        foreach (var paragraph in row.Descendants<Paragraph>())
        {
            var text = WordModel.Text.GetLogicalText(paragraph);
            var matches = Placeholder.Matches(text).Cast<Match>().OrderByDescending(m => m.Index).ToArray();
            foreach (var match in matches)
            {
                var name = match.Groups["name"].Value;
                var replacement = record.TryGetValue(name, out var value)
                    ? value ?? string.Empty
                    : missingBehavior == MissingTemplateValueBehavior.Empty ? string.Empty : match.Value;
                Replace(paragraph, match.Index, match.Length, replacement);
            }
        }
        RefreshDrawingIds(context, row);
        return row;
    }

    private static void Replace(Paragraph paragraph, int start, int length, string replacement)
    {
        var covered = WordModel.Text.IsolateSpan(paragraph, start, length);
        if (covered.Count == 0) return;
        WordModel.Dialect.SetRunText(covered[0], replacement);
        for (var i = 1; i < covered.Count; i++) covered[i].Remove();
    }

    private static void ClearParagraphIds(OpenXmlElement row)
    {
        foreach (var paragraph in row.Descendants<Paragraph>())
        {
            paragraph.RemoveAttribute("paraId", "http://schemas.microsoft.com/office/word/2010/wordml");
            paragraph.RemoveAttribute("textId", "http://schemas.microsoft.com/office/word/2010/wordml");
        }
    }

    private static void RefreshDrawingIds(ApplyContext context, OpenXmlElement row)
    {
        var ids = ImageNodeProvider.EnumerateDrawings(context.Package)
            .SelectMany(drawing => drawing.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>())
            .Select(p => p.Id?.Value ?? 0U).ToList();
        var next = ids.Count == 0 ? 1U : ids.Max() + 1U;
        foreach (var properties in row.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>())
            properties.Id = next++;
    }
}
