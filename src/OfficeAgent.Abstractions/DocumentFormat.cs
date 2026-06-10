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
    PowerPoint
}
