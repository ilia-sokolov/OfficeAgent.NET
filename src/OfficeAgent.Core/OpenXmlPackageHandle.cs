using DocumentFormat.OpenXml.Packaging;
using OfficeAgent.Abstractions;
using DocFormat = OfficeAgent.Abstractions.DocumentFormat;

namespace OfficeAgent.Core;

/// <summary>Default <see cref="IOpenXmlPackage"/> backed by an owned memory stream.</summary>
internal sealed class OpenXmlPackageHandle : IOpenXmlPackage
{
    private readonly MemoryStream _stream;
    private bool _disposed;

    internal OpenXmlPackageHandle(MemoryStream stream, OpenXmlPackage package, DocFormat format, bool editable)
    {
        _stream = stream;
        Package = package;
        Format = format;
        IsEditable = editable;
    }

    public OpenXmlPackage Package { get; }
    public DocFormat Format { get; }
    public bool IsEditable { get; }

    public string MainPartContentType => Format switch
    {
        DocFormat.Word => "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
        DocFormat.Excel => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
        DocFormat.PowerPoint => "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
        _ => "application/octet-stream"
    };

    public byte[] ToBytes()
    {
        Package.Save();
        _stream.Flush();
        return _stream.ToArray();
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        Package.Dispose();
        _stream.Dispose();
    }
}
