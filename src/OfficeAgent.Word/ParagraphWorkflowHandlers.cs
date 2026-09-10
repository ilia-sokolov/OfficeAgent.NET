using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.Word;

internal sealed class InsertParagraphsHandler : IOperationHandler
{
    public bool CanHandle(PlanOperation operation) =>
        operation is InsertParagraphsOp { Target: TextSpanAnchor };

    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (InsertParagraphsOp)operation;
        var anchor = (TextSpanAnchor)op.Target;
        var target = WordModel.ResolveParagraph(context, anchor.ParaId);
        if (target is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No paragraph with id '{anchor.ParaId}'.", anchor));
        if (target.Ancestors<Table>().Any())
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "insertParagraphs supports free-flowing Word paragraphs; use table-row operations inside a table.", anchor));
        if (op.Paragraphs.Count == 0)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "insertParagraphs requires at least one paragraph.", anchor));

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "insertParagraphs",
            Before = string.Empty,
            After = string.Join(" / ", op.Paragraphs.Take(3).Select(p => p.Text)),
            Context = $"{op.Position.ToString().ToLowerInvariant()} paragraph '{anchor.ParaId}'",
            BlastRadius = op.Paragraphs.Count
        });
    }

    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (InsertParagraphsOp)operation;
        var anchor = (TextSpanAnchor)op.Target;
        var target = WordModel.ResolveParagraph(context, anchor.ParaId)
            ?? throw new InvalidOperationException($"Paragraph '{anchor.ParaId}' vanished before apply.");

        var paragraphs = op.Paragraphs.Select(p => Build(target, p)).ToList();
        if (op.Position == InsertPosition.Before)
        {
            foreach (var paragraph in paragraphs)
                target.InsertBeforeSelf(paragraph);
        }
        else
        {
            OpenXmlElement cursor = target;
            foreach (var paragraph in paragraphs)
            {
                cursor.InsertAfterSelf(paragraph);
                cursor = paragraph;
            }
        }

        if (WordRevisionMarker.IsTracked(op.Mode))
        {
            var marker = new WordRevisionMarker(context.Package, context.Revision);
            foreach (var paragraph in paragraphs)
                marker.MarkParagraphInserted(paragraph);
        }
    }

    private static Paragraph Build(Paragraph neighbor, ParagraphData data)
    {
        var paragraph = new Paragraph();
        if (neighbor.ParagraphProperties is not null)
            paragraph.ParagraphProperties = (ParagraphProperties)neighbor.ParagraphProperties.CloneNode(true);
        if (data.StyleId is not null)
        {
            var properties = paragraph.ParagraphProperties ??= new ParagraphProperties();
            properties.ParagraphStyleId = new ParagraphStyleId { Val = data.StyleId };
        }
        paragraph.AppendChild(new Run(new Text(data.Text) { Space = SpaceProcessingModeValues.Preserve }));
        return paragraph;
    }
}

internal sealed class RemoveParagraphHandler : IOperationHandler
{
    public bool CanHandle(PlanOperation operation) =>
        operation is RemoveParagraphOp { Target: TextSpanAnchor };

    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (RemoveParagraphOp)operation;
        var anchor = (TextSpanAnchor)op.Target;
        var paragraph = WordModel.ResolveParagraph(context, anchor.ParaId);
        if (paragraph is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No paragraph with id '{anchor.ParaId}'.", anchor));
        var text = WordModel.Text.GetLogicalText(paragraph);
        if (!string.Equals(text, anchor.Expect, StringComparison.Ordinal) || anchor.Occurrence != 0)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.ExpectMismatch,
                $"removeParagraph requires the complete current paragraph text for '{anchor.ParaId}'.", anchor));
        if (!WordRevisionMarker.IsTracked(op.Mode) &&
            paragraph.Parent is Body body && body.Elements<Paragraph>().Count() == 1)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "A direct edit cannot remove the last body paragraph.", anchor));

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "removeParagraph",
            Before = text,
            After = string.Empty,
            Context = $"paragraph '{anchor.ParaId}'",
            BlastRadius = 1
        });
    }

    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (RemoveParagraphOp)operation;
        var anchor = (TextSpanAnchor)op.Target;
        var paragraph = WordModel.ResolveParagraph(context, anchor.ParaId)
            ?? throw new InvalidOperationException($"Paragraph '{anchor.ParaId}' vanished before apply.");
        if (WordRevisionMarker.IsTracked(op.Mode))
            new WordRevisionMarker(context.Package, context.Revision).MarkParagraphDeleted(paragraph);
        else
            paragraph.Remove();
    }
}
