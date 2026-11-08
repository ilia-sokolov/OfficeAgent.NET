using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeAgent.Abstractions;

namespace OfficeAgent.Core;

/// <summary>
/// The root element of a part the engine has already validated. The Open XML SDK types each
/// root as nullable because a part can be empty; a package that reaches a format module has a
/// root in every part it addresses, so a missing one is a malformed package and is refused as
/// one, with a message naming the part, rather than failing as a null dereference somewhere
/// later. A part whose root is some other element reads as having none.
/// </summary>
internal static class OpenXmlPartRoots
{
    /// <summary>The slide part's <c>p:sld</c> root.</summary>
    public static Slide SlideRoot(this SlidePart part) =>
        part.Slide ?? throw Missing("slide", "slide");

    /// <summary>The worksheet part's <c>x:worksheet</c> root.</summary>
    public static Worksheet WorksheetRoot(this WorksheetPart part) =>
        part.Worksheet ?? throw Missing("worksheet", "worksheet");

    /// <summary>The workbook part's <c>x:workbook</c> root.</summary>
    public static Workbook WorkbookRoot(this WorkbookPart part) =>
        part.Workbook ?? throw Missing("workbook", "workbook");

    /// <summary>The main document part's <c>w:document</c> root.</summary>
    public static Document DocumentRoot(this MainDocumentPart part) =>
        part.Document ?? throw Missing("main document", "document");

    /// <summary>The presentation part's <c>p:presentation</c> root.</summary>
    public static Presentation PresentationRoot(this PresentationPart part) =>
        part.Presentation ?? throw Missing("presentation", "presentation");

    private static OpenXmlPackageRejectedException Missing(string part, string element) =>
        new($"The package could not be read: the {part} part has no {element} element.");
}
