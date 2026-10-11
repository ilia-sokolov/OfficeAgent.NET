namespace OfficeAgent.Abstractions;

/// <summary>
/// Calculates the formula cells of a workbook, so an edit can leave correct cached values in
/// the file instead of none. Register one to enable <see cref="ApplyOptions.CalculateFormulas"/>.
/// </summary>
/// <remarks>
/// An evaluator calculates only what it can calculate correctly. Every other formula cell is
/// reported as <see cref="FormulaOutcome.NotCalculated"/> with a reason, and its cached value
/// is left as it was: a guessed value is worse than none, because readers trust a cached
/// value. The workbook stays marked for full recalculation, so Excel's own result always has
/// the final say when the file is opened.
/// </remarks>
public interface IFormulaEvaluator
{
    /// <summary>Gets the engine's name and version, as capability discovery reports it.</summary>
    string Name { get; }

    /// <summary>Gets the functions the evaluator calculates, in upper case.</summary>
    IReadOnlyList<string> SupportedFunctions { get; }

    /// <summary>Calculates every formula cell of the workbook, which is not modified.</summary>
    /// <param name="workbook">The workbook package bytes.</param>
    /// <param name="cancellationToken">Cancels the calculation.</param>
    /// <returns>One result per formula cell, or a reason the workbook could not be calculated.</returns>
    FormulaEvaluation Evaluate(byte[] workbook, CancellationToken cancellationToken = default);
}

/// <summary>The outcome of calculating a workbook.</summary>
public sealed class FormulaEvaluation
{
    /// <summary>Gets one result per formula cell.</summary>
    public IReadOnlyList<FormulaCellResult> Cells { get; init; } = Array.Empty<FormulaCellResult>();

    /// <summary>
    /// Gets why no cell was calculated, when the workbook as a whole could not be - for example
    /// because it contains a circular reference. Null when cells were evaluated one by one.
    /// </summary>
    public string? NotCalculatedReason { get; init; }

    /// <summary>Gets whether the workbook contains a circular reference.</summary>
    public bool CircularReference { get; init; }
}

/// <summary>Whether a formula cell was calculated.</summary>
public enum FormulaOutcome
{
    /// <summary>The cell was calculated; <see cref="FormulaCellResult.Value"/> is its result.</summary>
    Calculated,

    /// <summary>The cell was not calculated; <see cref="FormulaCellResult.Reason"/> says why.</summary>
    NotCalculated
}

/// <summary>The type of a calculated formula result.</summary>
public enum FormulaValueKind
{
    /// <summary>A number, including dates and times as Excel serial numbers.</summary>
    Number,

    /// <summary>Text.</summary>
    Text,

    /// <summary>TRUE or FALSE.</summary>
    Boolean,

    /// <summary>An Excel error value such as <c>#DIV/0!</c>.</summary>
    Error
}

/// <summary>The calculation result of one formula cell.</summary>
public sealed class FormulaCellResult
{
    /// <summary>Gets the durable sheet id of the cell's worksheet.</summary>
    public uint SheetId { get; init; }

    /// <summary>Gets the cell's A1 address.</summary>
    public string Address { get; init; } = string.Empty;

    /// <summary>Gets the cell's formula, without the leading <c>=</c>.</summary>
    public string Formula { get; init; } = string.Empty;

    /// <summary>Gets whether the cell was calculated.</summary>
    public FormulaOutcome Outcome { get; init; }

    /// <summary>Gets the result's type, for a calculated cell.</summary>
    public FormulaValueKind? ValueKind { get; init; }

    /// <summary>
    /// Gets the result as invariant text, for a calculated cell: a number in round-trip form,
    /// <c>TRUE</c>/<c>FALSE</c>, the text itself, or the error value such as <c>#DIV/0!</c>.
    /// </summary>
    public string? Value { get; init; }

    /// <summary>Gets why the cell was not calculated, for a cell that was not.</summary>
    public string? Reason { get; init; }
}

/// <summary>
/// What calculation did to a plan's workbook, reported on the <see cref="ChangeReport"/> of
/// a preview or commit that asked for <see cref="ApplyOptions.CalculateFormulas"/>.
/// </summary>
public sealed class FormulaCalculationReport
{
    /// <summary>Gets the evaluator's name and version.</summary>
    public string Engine { get; init; } = string.Empty;

    /// <summary>Gets the result of every formula cell in the edited workbook.</summary>
    public IReadOnlyList<FormulaCellResult> Cells { get; init; } = Array.Empty<FormulaCellResult>();

    /// <summary>Gets why nothing was calculated, when the workbook as a whole could not be.</summary>
    public string? NotCalculatedReason { get; init; }

    /// <summary>
    /// Gets the formula errors the plan introduced: cells whose result is an Excel error value
    /// they did not already carry, and circular references. A workbook's existing errors,
    /// such as a lookup that is meant to miss, are not listed.
    /// </summary>
    public IReadOnlyList<FormulaCellResult> NewErrors { get; init; } = Array.Empty<FormulaCellResult>();

    /// <summary>
    /// Gets whether the workbook contains a circular reference. The engine cannot tell one the
    /// plan introduced from one the workbook already had, so it is reported either way.
    /// </summary>
    public bool CircularReference { get; init; }

    /// <summary>Gets whether the plan introduced formula errors or the workbook has a circular reference.</summary>
    public bool HasFormulaErrors => NewErrors.Count > 0 || CircularReference;
}

/// <summary>What the registered formula evaluator offers, as capability discovery reports it.</summary>
public sealed class FormulaCalculationCapabilities
{
    /// <summary>Gets the evaluator's name and version.</summary>
    public string Engine { get; init; } = string.Empty;

    /// <summary>Gets the functions the evaluator calculates; formulas using others are reported as not calculated.</summary>
    public IReadOnlyList<string> SupportedFunctions { get; init; } = Array.Empty<string>();
}
