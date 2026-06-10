using DocumentFormat.OpenXml;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.Word;

/// <summary>
/// Accepts or rejects tracked revisions addressed by a revision <see cref="NodeAnchor"/>.
/// Accept: keep inserted runs, drop deleted runs. Reject: the inverse. Deterministic.
/// </summary>
internal sealed class RevisionHandler : IOperationHandler
{
    private readonly RevisionNodeProvider _provider = new();

    public bool CanHandle(PlanOperation operation) =>
        operation is RevisionOp { Target: NodeAnchor { Kind: "revision" } };

    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (RevisionOp)operation;
        var anchor = (NodeAnchor)op.Target;
        var node = _provider.Resolve(anchor, new WordObjectMap(context.Package));

        if (node is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No revision matches '{anchor.Path}'.", anchor));

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "revision",
            Before = "tracked revision(s)",
            After = op.Action == RevisionAction.Accept ? "accepted" : "rejected",
            Context = anchor.Path,
            BlastRadius = Math.Max(1, node.Elements.Count),
            Capability = Capability.Deterministic
        });
    }

    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (RevisionOp)operation;
        var anchor = (NodeAnchor)op.Target;

        var node = _provider.Resolve(anchor, new WordObjectMap(context.Package))
            ?? throw new InvalidOperationException($"Revision '{anchor.Path}' vanished before apply.");

        foreach (var element in node.Elements.ToList())
        {
            switch (element)
            {
                case InsertedRun ins when op.Action == RevisionAction.Accept:
                    Unwrap(ins);
                    break;
                case InsertedRun ins:
                    ins.Remove();
                    break;
                case DeletedRun del when op.Action == RevisionAction.Accept:
                    del.Remove();
                    break;
                case DeletedRun del:
                    RestoreDeleted(del);
                    Unwrap(del);
                    break;
            }
        }
    }

    private static void Unwrap(OpenXmlElement wrapper)
    {
        var parent = wrapper.Parent ?? throw new InvalidOperationException("Revision has no parent.");
        foreach (var child in wrapper.ChildElements.ToList())
        {
            child.Remove();
            parent.InsertBefore(child, wrapper);
        }
        wrapper.Remove();
    }

    private static void RestoreDeleted(DeletedRun del)
    {
        foreach (var delText in del.Descendants<DeletedText>().ToList())
        {
            var text = new Text(delText.Text ?? string.Empty) { Space = SpaceProcessingModeValues.Preserve };
            delText.InsertAfterSelf(text);
            delText.Remove();
        }
    }
}
