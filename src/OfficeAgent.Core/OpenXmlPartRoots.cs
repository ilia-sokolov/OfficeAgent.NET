using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.Core;

/// <summary>
/// The root element of a part the engine has already validated. The Open XML SDK types each
/// root as nullable because a part can be empty; a package that reaches a format module has a
/// root in every part it addresses, so a missing one is a malformed package and fails with a
/// message naming the part, as <c>WordObjectMap.Main</c> does for a missing main part, rather
/// than as a null dereference somewhere later.
/// </summary>
internal static class OpenXmlPartRoots
{
    /// <summary>The slide part's <c>p:sld</c> root.</summary>
    public static Slide SlideRoot(this SlidePart part) =>
        part.Slide ?? throw new InvalidOperationException("The slide part has no slide element.");

    /// <summary>The worksheet part's <c>x:worksheet</c> root.</summary>
    public static Worksheet WorksheetRoot(this WorksheetPart part) =>
        part.Worksheet ?? throw new InvalidOperationException("The worksheet part has no worksheet element.");

    /// <summary>The workbook part's <c>x:workbook</c> root.</summary>
    public static Workbook WorkbookRoot(this WorkbookPart part) =>
        part.Workbook ?? throw new InvalidOperationException("The workbook part has no workbook element.");

    /// <summary>The main document part's <c>w:document</c> root.</summary>
    public static Document DocumentRoot(this MainDocumentPart part) =>
        part.Document ?? throw new InvalidOperationException("The main document part has no document element.");
}
