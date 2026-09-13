using OfficeAgent.Abstractions;

namespace OfficeAgent.Core;

/// <summary>Central compatibility check for every typed document-plan entry point.</summary>
internal static class DocumentPlanCompatibility
{
    private const string Remediation =
        "Omit contractVersion for legacy 0.2 behavior or set it to \"0.2\"; " +
        "null, empty, malformed, and unknown versions are rejected.";

    public static ValidationError? Error(DocumentPlan plan)
    {
        if (plan is null) throw new ArgumentNullException(nameof(plan));

        return string.Equals(
            plan.ContractVersion,
            DocumentPlan.CurrentContractVersion,
            StringComparison.Ordinal)
            ? null
            : new ValidationError(
                ValidationErrorCodes.ContractMismatch,
                $"Plan contractVersion must be \"{DocumentPlan.CurrentContractVersion}\". {Remediation}");
    }

    public static ChangeReport? InvalidReport(DocumentPlan plan) =>
        Error(plan) is { } error ? ChangeReport.Invalid(error) : null;

    public static ApplyResult? RejectedResult(DocumentPlan plan) =>
        InvalidReport(plan) is { } report
            ? new ApplyResult { Report = report, Committed = false, Output = null, Receipt = null }
            : null;
}
