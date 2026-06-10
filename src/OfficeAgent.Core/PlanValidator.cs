using OfficeAgent.Abstractions;

namespace OfficeAgent.Core;

/// <summary>
/// Validates a plan against the live document by previewing every operation
/// through its handler. Produces the dry-run change report. No mutation.
/// </summary>
internal sealed class PlanValidator
{
    public ChangeReport Validate(ApplyContext context, DocumentPlan plan, IFormatModule module)
    {
        var changes = new List<ProposedChange>();
        var errors = new List<ValidationError>();

        // contractVersion is informational while the wire schema is pre-1.0. LLMs
        // routinely supply "1.0" or omit the field; failing the entire plan over a
        // string mismatch is anti-ergonomic. The field is preserved on DocumentPlan
        // for forward compatibility; once the schema has real breaking versions, this
        // check can come back behind a stricter policy.

        if (plan.Format != module.Format)
        {
            errors.Add(new ValidationError(
                ValidationErrorCodes.ContractMismatch,
                $"Plan targets {plan.Format} but the document is {module.Format}."));
        }

        if (plan.Snapshot is { ETag.Length: > 0 } planSnap
            && !string.Equals(planSnap.ETag, context.Inspection.Snapshot.ETag, StringComparison.Ordinal))
        {
            errors.Add(new ValidationError(
                ValidationErrorCodes.StaleSnapshot,
                $"Plan was authored against snapshot '{planSnap.ETag}' but the live document is at '{context.Inspection.Snapshot.ETag}'. Re-inspect and rebuild the plan."));
        }

        DetectConflicts(plan, errors);

        foreach (var operation in plan.Operations)
        {
            var handler = module.Handlers.FirstOrDefault(h => h.CanHandle(operation));
            if (handler is null)
            {
                errors.Add(new ValidationError(
                    ValidationErrorCodes.UnsupportedOperation,
                    $"No handler for operation '{operation.GetType().Name}' in the {module.Format} module.",
                    operation.Target));
                continue;
            }

            var preview = handler.Preview(context, operation);
            if (preview.Error is not null)
                errors.Add(preview.Error);
            if (preview.Change is not null)
            {
                changes.Add(preview.Change);

                if (preview.Change.Capability == Capability.NeedsRenderer)
                    errors.Add(new ValidationError(
                        ValidationErrorCodes.RequiresRenderer,
                        $"Operation '{operation.GetType().Name}' needs a layout/rendering engine " +
                        "(field/cross-reference value, pagination); the OOXML engine cannot complete it.",
                        operation.Target));
            }
        }

        return new ChangeReport
        {
            IsValid = errors.Count == 0,
            Changes = changes,
            Errors = errors
        };
    }





    private static void DetectConflicts(DocumentPlan plan, List<ValidationError> errors)
    {
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (var operation in plan.Operations)
        {
            var key = Identity(operation);
            if (key is null) continue;
            if (!seen.Add(key))
                errors.Add(new ValidationError(
                    ValidationErrorCodes.OperationConflict,
                    $"Multiple operations target the same location ('{key}'); merge them or apply in separate plans.",
                    operation.Target));
        }
    }

    // Conflict keys are scoped to the op type so unrelated verbs at the same anchor don't
    // collide (e.g. format + insertImage on the same paragraph). Insert-style ops also
    // include their positional slot, since "before" and "after" the same paragraph are
    // independent destinations.
    private static string? Identity(PlanOperation operation)
    {
        var anchor = AnchorKey(operation.Target);
        if (anchor is null) return null;
        var slot = operation switch
        {
            InsertOp i => $":{i.Position}",
            InsertImageOp i => $":{i.Position}",
            InsertTableRowsOp i => $":{i.Position}:{i.RowIndex}",
            InsertTableColumnsOp i => $":{i.Position}:{i.ColumnIndex}",
            _ => string.Empty
        };
        return $"{operation.GetType().Name}:{anchor}{slot}";
    }

    private static string? AnchorKey(Anchor? anchor) => anchor switch
    {
        TextSpanAnchor t => $"text:{t.ParaId}:{t.Expect}:{t.Occurrence}",
        NodeAnchor n => $"node:{n.Kind}:{n.Path}:{n.Occurrence}",
        StructuralAnchor s => $"struct:{s.Tag}",
        StyleAnchor st => $"style:{st.StyleId}",
        null => null,
        _ => $"id:{anchor.Id}"
    };
}
