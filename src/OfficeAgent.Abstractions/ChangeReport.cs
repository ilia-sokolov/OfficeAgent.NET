namespace OfficeAgent.Abstractions;

/// <summary>
/// Contains the result of validating or applying a document plan.
/// </summary>
public sealed class ChangeReport
{
    /// <summary>
    /// Gets a value indicating whether the plan passed validation.
    /// </summary>
    public bool IsValid { get; init; }

    /// <summary>
    /// Gets the proposed or applied changes in plan order.
    /// </summary>
    public IReadOnlyList<ProposedChange> Changes { get; init; } = Array.Empty<ProposedChange>();

    /// <summary>
    /// Gets validation errors that prevented the plan from being committed.
    /// </summary>
    public IReadOnlyList<ValidationError> Errors { get; init; } = Array.Empty<ValidationError>();

    /// <summary>
    /// Creates an invalid report from one or more validation errors.
    /// </summary>
    /// <param name="errors">The validation errors to include in the report.</param>
    /// <returns>An invalid change report.</returns>
    public static ChangeReport Invalid(params ValidationError[] errors) =>
        new() { IsValid = false, Errors = errors };
}

/// <summary>
/// Describes one proposed or applied document change.
/// </summary>
public sealed class ProposedChange
{
    /// <summary>
    /// Gets the operation target.
    /// </summary>
    public Anchor? Target { get; init; }

    /// <summary>
    /// Gets the plan verb that produced the change.
    /// </summary>
    public string Verb { get; init; } = string.Empty;

    /// <summary>
    /// Gets the target content before the change.
    /// </summary>
    public string Before { get; init; } = string.Empty;

    /// <summary>
    /// Gets the target content after the change.
    /// </summary>
    public string After { get; init; } = string.Empty;

    /// <summary>
    /// Gets nearby text or object context for display in a preview.
    /// </summary>
    public string Context { get; init; } = string.Empty;

    /// <summary>
    /// Gets a relative estimate of how much document content the change affects.
    /// </summary>
    public int BlastRadius { get; init; } = 1;

    /// <summary>
    /// Gets the engine capability required to complete the change.
    /// </summary>
    public Capability Capability { get; init; } = Capability.Deterministic;
}

/// <summary>
/// Describes a validation error that prevents safe application of a plan.
/// </summary>
public sealed class ValidationError
{
    /// <summary>
    /// Gets a stable machine-readable error code (string form, suitable for the wire).
    /// </summary>
    public string Code { get; init; } = string.Empty;

    /// <summary>
    /// Gets the strongly typed error code for switch/exhaustive consumers.
    /// </summary>
    public ValidationErrorCode CodeKind => ValidationErrorCodes.Parse(Code);

    /// <summary>
    /// Gets a human-readable error message.
    /// </summary>
    public string Message { get; init; } = string.Empty;

    /// <summary>
    /// Gets the target associated with the error, when one is available.
    /// </summary>
    public Anchor? Target { get; init; }

    /// <summary>
    /// Initializes a new instance of the <see cref="ValidationError"/> class.
    /// </summary>
    public ValidationError()
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="ValidationError"/> class.
    /// </summary>
    /// <param name="code">The stable error code.</param>
    /// <param name="message">The display message.</param>
    /// <param name="target">The optional target associated with the error.</param>
    public ValidationError(string code, string message, Anchor? target = null)
    {
        Code = code;
        Message = message;
        Target = target;
    }

    /// <summary>
    /// Initializes a new instance from the strongly typed <see cref="ValidationErrorCode"/>.
    /// </summary>
    public ValidationError(ValidationErrorCode code, string message, Anchor? target = null)
        : this(ValidationErrorCodes.ToWireCode(code), message, target)
    {
    }
}

/// <summary>
/// Strongly typed enumeration of engine-emitted validation error codes. Consumers can
/// switch exhaustively; the wire still uses the string in <see cref="ValidationErrorCodes"/>.
/// </summary>
public enum ValidationErrorCode
{
    /// <summary>An unrecognised string code (forward compatibility).</summary>
    Unknown,

    /// <summary>The inspected document snapshot no longer matches the document being edited.</summary>
    StaleSnapshot,

    /// <summary>The target anchor could not be resolved.</summary>
    AnchorNotFound,

    /// <summary>The live content did not match the expected anchor content.</summary>
    ExpectMismatch,

    /// <summary>The target address is not specific enough to edit safely.</summary>
    AmbiguousAnchor,

    /// <summary>No registered handler supports the requested operation.</summary>
    UnsupportedOperation,

    /// <summary>The plan contract version does not match the engine contract version.</summary>
    ContractMismatch,

    /// <summary>The operation is malformed or invalid for the target.</summary>
    InvalidOperation,

    /// <summary>The operation requires a renderer or calculation engine.</summary>
    RequiresRenderer,

    /// <summary>Multiple operations in the same plan target the same location.</summary>
    OperationConflict
}

/// <summary>
/// Provides stable validation error codes returned by the engine. Use the string
/// constants on the wire and <see cref="ValidationErrorCode"/> in switch expressions.
/// </summary>
public static class ValidationErrorCodes
{
    /// <summary>The inspected document snapshot no longer matches the document being edited.</summary>
    public const string StaleSnapshot = "stale-snapshot";

    /// <summary>The target anchor could not be resolved.</summary>
    public const string AnchorNotFound = "anchor-not-found";

    /// <summary>The live content did not match the expected anchor content.</summary>
    public const string ExpectMismatch = "expect-mismatch";

    /// <summary>The target address is not specific enough to edit safely.</summary>
    public const string AmbiguousAnchor = "ambiguous-anchor";

    /// <summary>No registered handler supports the requested operation.</summary>
    public const string UnsupportedOperation = "unsupported-operation";

    /// <summary>The plan contract version does not match the engine contract version.</summary>
    public const string ContractMismatch = "contract-mismatch";

    /// <summary>The operation is malformed or invalid for the target.</summary>
    public const string InvalidOperation = "invalid-operation";

    /// <summary>The operation requires a renderer or calculation engine.</summary>
    public const string RequiresRenderer = "requires-renderer";

    /// <summary>Multiple operations in the same plan target the same location.</summary>
    public const string OperationConflict = "operation-conflict";

    /// <summary>Parses a wire code to its strongly typed enum value.</summary>
    public static ValidationErrorCode Parse(string code) => code switch
    {
        StaleSnapshot => ValidationErrorCode.StaleSnapshot,
        AnchorNotFound => ValidationErrorCode.AnchorNotFound,
        ExpectMismatch => ValidationErrorCode.ExpectMismatch,
        AmbiguousAnchor => ValidationErrorCode.AmbiguousAnchor,
        UnsupportedOperation => ValidationErrorCode.UnsupportedOperation,
        ContractMismatch => ValidationErrorCode.ContractMismatch,
        InvalidOperation => ValidationErrorCode.InvalidOperation,
        RequiresRenderer => ValidationErrorCode.RequiresRenderer,
        OperationConflict => ValidationErrorCode.OperationConflict,
        _ => ValidationErrorCode.Unknown
    };

    /// <summary>Returns the stable wire string for an enum value.</summary>
    public static string ToWireCode(ValidationErrorCode code) => code switch
    {
        ValidationErrorCode.StaleSnapshot => StaleSnapshot,
        ValidationErrorCode.AnchorNotFound => AnchorNotFound,
        ValidationErrorCode.ExpectMismatch => ExpectMismatch,
        ValidationErrorCode.AmbiguousAnchor => AmbiguousAnchor,
        ValidationErrorCode.UnsupportedOperation => UnsupportedOperation,
        ValidationErrorCode.ContractMismatch => ContractMismatch,
        ValidationErrorCode.InvalidOperation => InvalidOperation,
        ValidationErrorCode.RequiresRenderer => RequiresRenderer,
        ValidationErrorCode.OperationConflict => OperationConflict,
        _ => "unknown"
    };
}
