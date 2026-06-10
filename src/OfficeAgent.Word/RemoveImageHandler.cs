using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeAgent.Word;

/// <summary>
/// Removes an inline image addressed by a <see cref="NodeAnchor"/> with
/// <c>Kind="image"</c> and <c>Path="image#N"</c>. The drawing element is removed along
/// with its parent run when the run is left empty, and the underlying
/// <see cref="ImagePart"/> resource is released once no other drawing in the same host
/// part still references it. Removing an image therefore leaves no orphaned image
/// bytes behind in the saved package, while images shared by another drawing are kept
/// intact.
/// </summary>
internal sealed class RemoveImageHandler : IOperationHandler
{
    private readonly ImageNodeProvider _images = new();

    public bool CanHandle(PlanOperation operation) =>
        operation is RemoveImageOp { Target: NodeAnchor { Kind: "image" } };

    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (RemoveImageOp)operation;
        var anchor = (NodeAnchor)op.Target;
        var node = _images.ResolveWithHost(anchor, new WordObjectMap(context.Package));
        if (node is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No image at path '{anchor.Path}'.", anchor));

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "removeImage",
            Before = "(inline image)",
            After = "(removed)",
            Context = anchor.Path,
            BlastRadius = 1
        });
    }

    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (RemoveImageOp)operation;
        var anchor = (NodeAnchor)op.Target;
        var located = _images.ResolveWithHost(anchor, new WordObjectMap(context.Package))
            ?? throw new InvalidOperationException($"Image '{anchor.Path}' vanished before apply.");

        var (drawing, host) = located;

        // The drawing references its bytes through a blip relationship scoped to the
        // hosting part; capture it before the drawing is detached.
        var relationshipId = drawing.Descendants<A.Blip>().FirstOrDefault()?.Embed?.Value;

        var run = drawing.Parent as Run;
        drawing.Remove();
        if (run is not null && !run.HasChildren)
            run.Remove();

        // Release the underlying image resource once nothing else points at it, so the
        // removal cleans up all related resources rather than orphaning the ImagePart.
        if (!string.IsNullOrEmpty(relationshipId))
            ReleaseImagePartIfOrphaned(host, relationshipId!);
    }

    private static void ReleaseImagePartIfOrphaned(OpenXmlPart host, string relationshipId)
    {
        if (host.GetPartById(relationshipId) is not ImagePart) return;

        // Relationship ids are scoped to the hosting part, so a surviving reference can
        // only come from a drawing that still lives in this same part.
        var stillReferenced = host.RootElement?
            .Descendants<A.Blip>()
            .Any(blip => string.Equals(blip.Embed?.Value, relationshipId, StringComparison.Ordinal)) ?? false;

        if (!stillReferenced)
            host.DeletePart(relationshipId);
    }
}
