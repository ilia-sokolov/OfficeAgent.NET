using System.IO.Compression;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocFormat = OfficeAgent.Abstractions.DocumentFormat;

namespace OfficeAgent.Core;

/// <summary>
/// Opens bytes as an in-memory Open XML package. Avoids a second copy when the
/// caller already owns the bytes and detects the format from <c>[Content_Types].xml</c>
/// in a single zip pass.
/// </summary>
internal sealed class OpenXmlPackageService
{
    public IOpenXmlPackage Open(byte[] bytes, bool editable)
    {
        var memory = new MemoryStream();
        memory.Write(bytes, 0, bytes.Length);
        memory.Position = 0;
        return OpenOwned(memory, editable);
    }

    public IOpenXmlPackage Open(Stream source, bool editable)
    {
        if (source is MemoryStream ms && ms.TryGetBuffer(out var buf))
        {
            var owned = new byte[buf.Count];
            Array.Copy(buf.Array!, buf.Offset, owned, 0, buf.Count);
            return Open(owned, editable);
        }

        var memory = new MemoryStream();
        source.CopyTo(memory);
        memory.Position = 0;
        return OpenOwned(memory, editable);
    }

    public void Save(IOpenXmlPackage package, Stream destination)
    {
        var bytes = package.ToBytes();
        destination.Write(bytes, 0, bytes.Length);
    }

    private static IOpenXmlPackage OpenOwned(MemoryStream memory, bool editable)
    {
        var format = DetectFormat(memory);
        memory.Position = 0;

        OpenXmlPackage package = format switch
        {
            DocFormat.Word => WordprocessingDocument.Open(memory, editable),
            DocFormat.Excel => SpreadsheetDocument.Open(memory, editable),
            DocFormat.PowerPoint => PresentationDocument.Open(memory, editable),
            _ => throw new NotSupportedException($"Unsupported format: {format}.")
        };

        return new OpenXmlPackageHandle(memory, package, format, editable);
    }

    private static DocFormat DetectFormat(Stream stream)
    {
        stream.Position = 0;
        using var zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
        var contentTypes = zip.GetEntry("[Content_Types].xml")
            ?? throw new NotSupportedException("Not an OOXML package: missing [Content_Types].xml.");

        using var reader = contentTypes.Open();
        var manifest = XDocument.Load(reader);
        var mainTypes = manifest.Root?.Elements()
            .Select(element => (string?)element.Attribute("ContentType"))
            .Where(contentType => contentType?.EndsWith(".main+xml", StringComparison.OrdinalIgnoreCase) == true)
            .ToList() ?? new List<string?>();

        if (mainTypes.Any(type => type!.Contains("wordprocessingml"))) return DocFormat.Word;
        if (mainTypes.Any(type => type!.Contains("spreadsheetml"))) return DocFormat.Excel;
        if (mainTypes.Any(type => type!.Contains("presentationml"))) return DocFormat.PowerPoint;

        throw new NotSupportedException("Unrecognised OOXML content types.");
    }
}
