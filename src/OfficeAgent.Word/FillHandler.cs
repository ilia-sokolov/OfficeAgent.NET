using DocumentFormat.OpenXml;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.Word;

/// <summary>
/// Populates a Word content control (by tag) without disturbing surrounding styles.
/// </summary>
internal sealed class FillHandler : IOperationHandler
{
    public bool CanHandle(PlanOperation operation) =>
        operation is FillOp { Target: StructuralAnchor };

    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (FillOp)operation;
        var anchor = (StructuralAnchor)op.Target;

        var sdt = FindContentControl(context, anchor.Tag);
        if (sdt is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No content control with tag '{anchor.Tag}'.", anchor));

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "fill",
            Before = CurrentText(sdt),
            After = op.Value,
            Context = $"content control '{anchor.Tag}'",
            BlastRadius = 1
        });
    }

    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (FillOp)operation;
        var anchor = (StructuralAnchor)op.Target;

        var sdt = FindContentControl(context, anchor.Tag)
            ?? throw new InvalidOperationException($"Content control '{anchor.Tag}' not found at apply time.");

        SetText(sdt, op.Value);
    }

    private static SdtElement? FindContentControl(ApplyContext context, string tag)
    {
        foreach (var (root, _) in WordModel.TextHosts(context.Package))
        {
            var match = root.Descendants<SdtElement>()
                .FirstOrDefault(sdt => string.Equals(
                    sdt.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value,
                    tag,
                    StringComparison.Ordinal));
            if (match is not null) return match;
        }
        return null;
    }

    private static OpenXmlElement? ContentOf(SdtElement sdt) =>
        sdt.Descendants<SdtContentRun>().FirstOrDefault() as OpenXmlElement
        ?? sdt.Descendants<SdtContentBlock>().FirstOrDefault();

    private static string CurrentText(SdtElement sdt)
    {
        var content = ContentOf(sdt);
        if (content is null) return string.Empty;
        return string.Concat(content.Descendants<Text>().Select(t => t.Text));
    }

    private static void SetText(SdtElement sdt, string value)
    {
        var content = ContentOf(sdt);
        if (content is null)
            return;

        var texts = content.Descendants<Text>().ToList();
        if (texts.Count > 0)
        {
            texts[0].Text = value;
            texts[0].Space = SpaceProcessingModeValues.Preserve;
            for (int i = 1; i < texts.Count; i++)
                texts[i].Text = string.Empty;
            return;
        }

        var run = new Run(new Text(value) { Space = SpaceProcessingModeValues.Preserve });
        if (content is SdtContentBlock)
            content.AppendChild(new Paragraph(run));
        else
            content.AppendChild(run);
    }
}
