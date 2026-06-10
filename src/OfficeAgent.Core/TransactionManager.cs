using OfficeAgent.Abstractions;

namespace OfficeAgent.Core;

/// <summary>
/// Applies a plan as an all-or-nothing transaction. Mutations land on an
/// in-memory editable package; each operation is re-validated against live state
/// immediately before it is applied, so an earlier op that shifts offsets cannot
/// silently corrupt a later op's target.
/// </summary>
internal sealed class TransactionManager
{
    /// <summary>
    /// Result of an atomic apply: either all operations succeeded, or none did and
    /// the failing operation's <see cref="ValidationError"/> is reported.
    /// </summary>
    internal sealed class ApplyOutcome
    {
        public bool Success { get; init; }
        public ValidationError? Error { get; init; }

        public static ApplyOutcome Ok() => new() { Success = true };
        public static ApplyOutcome Fail(ValidationError error) => new() { Success = false, Error = error };
    }

    public ApplyOutcome ApplyAtomic(ApplyContext context, DocumentPlan plan, IFormatModule module, CancellationToken cancellationToken = default)
    {
        foreach (var operation in plan.Operations)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var handler = module.Handlers.FirstOrDefault(h => h.CanHandle(operation));
            if (handler is null)
                return ApplyOutcome.Fail(new ValidationError(
                    ValidationErrorCodes.UnsupportedOperation,
                    $"No handler for operation '{operation.GetType().Name}'.",
                    operation.Target));

            var preview = handler.Preview(context, operation);
            if (preview.Error is not null)
                return ApplyOutcome.Fail(preview.Error);

            try
            {
                handler.Apply(context, operation);
            }
            catch (Exception ex)
            {
                return ApplyOutcome.Fail(new ValidationError(
                    ValidationErrorCodes.InvalidOperation,
                    $"Apply failed for '{operation.GetType().Name}': {ex.Message}",
                    operation.Target));
            }
        }
        return ApplyOutcome.Ok();
    }
}
