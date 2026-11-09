using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeAgent.Abstractions;
using DocFormat = OfficeAgent.Abstractions.DocumentFormat;

namespace OfficeAgent.Core;

/// <summary>
/// Turns the bytes of a template into the bytes of a document a new file can start from.
/// </summary>
/// <remarks>
/// <para>
/// What makes a <c>.dotx</c>, <c>.potx</c> or <c>.xltx</c> a template is the content type of
/// its main part, recorded in <c>[Content_Types].xml</c>. A copy that keeps it, saved under a
/// <c>.docx</c> name, is a template wearing a document's extension: Word refuses to open it or
/// opens a new document based on it, and either way the user is not looking at the file that
/// was made. The main part is therefore retyped as a document's, which is what Office does when
/// it creates a document from a template. A slideshow (<c>.ppsx</c>) is retyped as a
/// presentation for the same reason. An ordinary document is copied unchanged.
/// </para>
/// <para>
/// A macro-enabled package is refused rather than converted: dropping the VBA project would
/// hand back a different template than the one asked for, and keeping it would put macros in a
/// <c>.docx</c>, which Office then refuses.
/// </para>
/// </remarks>
internal static class TemplatePackages
{
    /// <summary>The bytes a new document starts from, and the format they hold.</summary>
    internal sealed class Converted
    {
        public Converted(byte[] original, byte[] bytes, DocFormat format, bool wasRetyped)
        {
            Original = original;
            Bytes = bytes;
            Format = format;
            WasRetyped = wasRetyped;
        }

        /// <summary>Gets the template's own bytes, exactly as stored.</summary>
        public byte[] Original { get; }

        /// <summary>Gets the document bytes: the template's own when nothing needed retyping.</summary>
        public byte[] Bytes { get; }

        /// <summary>Gets the format family of the template.</summary>
        public DocFormat Format { get; }

        /// <summary>Gets whether the main part was a template's or a slideshow's and was retyped.</summary>
        public bool WasRetyped { get; }

        /// <summary>Gets the extension a document of this format is saved under.</summary>
        public string DocumentExtension => DocumentExtensionOf(Format);
    }

    /// <summary>The extension a document of <paramref name="format"/> is saved under.</summary>
    public static string DocumentExtensionOf(DocFormat format) => format switch
    {
        DocFormat.Word => ".docx",
        DocFormat.PowerPoint => ".pptx",
        DocFormat.Excel => ".xlsx",
        _ => string.Empty
    };

    /// <summary>
    /// Returns the bytes of a document equivalent to <paramref name="template"/>, refusing a
    /// protected or macro-enabled package. The input array is never modified.
    /// </summary>
    /// <exception cref="EncryptedDocumentException">The package is encrypted.</exception>
    /// <exception cref="MacroEnabledDocumentException">The package is macro-enabled or carries a VBA project.</exception>
    /// <exception cref="OpenXmlPackageRejectedException">The bytes are not a readable Office package.</exception>
    /// <exception cref="OpenXmlIngestionLimitException">The package exceeds the ingestion limits.</exception>
    public static Converted AsDocument(byte[] template, OpenXmlIngestionLimits limits)
    {
        if (template is null) throw new ArgumentNullException(nameof(template));

        // Checked before the package service sees the bytes: it does not recognise a
        // macro-enabled main part at all, and would report the file as merely malformed.
        if (EncryptedPackageDetector.IsEncrypted(template, template.Length))
            throw new EncryptedDocumentException();
        if (IsMacroEnabled(template, limits))
            throw new MacroEnabledDocumentException();

        // The package service holds the template to the host's ingestion limits before the SDK
        // copies any part, and opens a copy, so the caller's array is left as it was.
        using var package = new OpenXmlPackageService(limits).Open(template, editable: true);
        var retyped = OpenXmlPackageService.Guarded(() => Retype(package.Package));
        return new Converted(template, retyped ? package.ToBytes() : template, package.Format, retyped);
    }

    /// <summary>Changes a template's or slideshow's main part to a document's; false when it already is one.</summary>
    private static bool Retype(OpenXmlPackage package)
    {
        switch (package)
        {
            case WordprocessingDocument word when word.DocumentType != WordprocessingDocumentType.Document:
                word.ChangeDocumentType(WordprocessingDocumentType.Document);
                return true;
            case PresentationDocument deck when deck.DocumentType != PresentationDocumentType.Presentation:
                deck.ChangeDocumentType(PresentationDocumentType.Presentation);
                return true;
            case SpreadsheetDocument book when book.DocumentType != SpreadsheetDocumentType.Workbook:
                book.ChangeDocumentType(SpreadsheetDocumentType.Workbook);
                return true;
            default:
                return false;
        }
    }

    /// <summary>
    /// Whether the package declares a macro-enabled main part or carries VBA. Anything that
    /// cannot be read as an archive is not judged here; the package service refuses it next.
    /// </summary>
    private static bool IsMacroEnabled(byte[] bytes, OpenXmlIngestionLimits limits)
    {
        try
        {
            using var zip = new ZipArchive(new MemoryStream(bytes, writable: false), ZipArchiveMode.Read);
            if (zip.Entries.Any(entry => entry.Name.EndsWith(".bin", StringComparison.OrdinalIgnoreCase) &&
                                         entry.Name.StartsWith("vba", StringComparison.OrdinalIgnoreCase)))
                return true;

            var contentTypes = zip.GetEntry("[Content_Types].xml");
            if (contentTypes is null) return false;

            var settings = new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                MaxCharactersInDocument = limits.MaximumXmlCharacters,
                MaxCharactersFromEntities = 0
            };
            using var stream = contentTypes.Open();
            using var reader = XmlReader.Create(stream, settings);
            var manifest = XDocument.Load(reader);
            return manifest.Root?.Elements().Any(element =>
                (string?)element.Attribute("ContentType") is { } type &&
                (type.IndexOf("macroEnabled", StringComparison.OrdinalIgnoreCase) >= 0 ||
                 type.IndexOf("vbaProject", StringComparison.OrdinalIgnoreCase) >= 0)) ?? false;
        }
        catch (Exception ex) when (ex is InvalidDataException or XmlException)
        {
            return false;
        }
    }
}
