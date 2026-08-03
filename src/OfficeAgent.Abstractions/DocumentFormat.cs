namespace OfficeAgent.Abstractions;

/// <summary>
/// Identifies the Office document format handled by a plan or inspection result.
/// </summary>
public enum DocumentFormat
{
    /// <summary>
    /// Microsoft Word Open XML document format.
    /// </summary>
    Word,

    /// <summary>
    /// Microsoft Excel Open XML workbook format.
    /// </summary>
    Excel,

    /// <summary>
    /// Microsoft PowerPoint Open XML presentation format.
    /// </summary>
    PowerPoint,

    /// <summary>
    /// No particular format: the plan applies to whichever format the document turns out
    /// to be. This is the default for <see cref="DocumentPlan.Format"/>, so a plan that
    /// simply lists operations - the shape the agent tools document - is not tied to one
    /// format by omission.
    /// </summary>
    /// <remarks>
    /// Declared last so the existing members keep their numeric values, which persisted
    /// plans and stored inspection results may already carry.
    /// </remarks>
    Unspecified
}
