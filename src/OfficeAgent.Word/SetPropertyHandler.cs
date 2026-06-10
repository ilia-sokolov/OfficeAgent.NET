using DocumentFormat.OpenXml.Packaging;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.Word;

/// <summary>
/// Sets a named property on an addressed node, including supported document properties and document-level settings.
/// </summary>
internal sealed class SetPropertyHandler : IOperationHandler
{
    public bool CanHandle(PlanOperation operation) =>
        operation is SetPropertyOp { Target: NodeAnchor };

    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (SetPropertyOp)operation;
        var anchor = (NodeAnchor)op.Target;
        var map = new WordObjectMap(context.Package);

        switch (anchor.Kind)
        {
            case "docProperty":
            {
                var name = DocPropertyNodeProvider.NameOf(anchor.Path);
                if (!DocPropertyNodeProvider.IsCore(name))
                    return OperationPreview.Fail(new ValidationError(
                        ValidationErrorCodes.AnchorNotFound,
                        $"Unknown document property '{anchor.Path}'.", anchor));

                return OperationPreview.Ok(new ProposedChange
                {
                    Target = anchor,
                    Verb = "setProperty",
                    Before = DocPropertyNodeProvider.Read(map, name) ?? string.Empty,
                    After = op.Value ?? string.Empty,
                    Context = anchor.Path,
                    BlastRadius = 1,
                    Capability = Capability.Deterministic
                });
            }

            case "field" when op.Name == "updateOnOpen":
                return OperationPreview.Ok(new ProposedChange
                {
                    Target = anchor,
                    Verb = "setProperty",
                    Before = "fields static",
                    After = "update fields on open",
                    Context = "document settings",
                    BlastRadius = 1,
                    Capability = Capability.DeferredToWordOnOpen
                });

            case "field" when op.Name == "refresh":
                return OperationPreview.Ok(new ProposedChange
                {
                    Target = anchor,
                    Verb = "setProperty",
                    Before = "field result",
                    After = "recomputed value",
                    Context = anchor.Path,
                    BlastRadius = 1,
                    Capability = Capability.NeedsRenderer
                });

            default:
                return OperationPreview.Fail(new ValidationError(
                    ValidationErrorCodes.UnsupportedOperation,
                    $"setProperty does not support node kind '{anchor.Kind}' with name '{op.Name}'.", anchor));
        }
    }

    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (SetPropertyOp)operation;
        var anchor = (NodeAnchor)op.Target;
        var map = new WordObjectMap(context.Package);

        switch (anchor.Kind)
        {
            case "docProperty":
                DocPropertyNodeProvider.Write(map, DocPropertyNodeProvider.NameOf(anchor.Path), op.Value);
                break;

            case "field" when op.Name == "updateOnOpen":
                SetUpdateFieldsOnOpen(map);
                break;

            default:
                throw new InvalidOperationException(
                    $"setProperty cannot apply node kind '{anchor.Kind}' / name '{op.Name}' in the pure engine.");
        }
    }

    private static void SetUpdateFieldsOnOpen(WordObjectMap map)
    {
        var settingsPart = map.Main.DocumentSettingsPart
            ?? map.Main.AddNewPart<DocumentSettingsPart>();
        settingsPart.Settings ??= new Settings();

        var existing = settingsPart.Settings.GetFirstChild<UpdateFieldsOnOpen>();
        if (existing is null)
            settingsPart.Settings.PrependChild(new UpdateFieldsOnOpen { Val = true });
        else
            existing.Val = true;

        settingsPart.Settings.Save();
    }
}
