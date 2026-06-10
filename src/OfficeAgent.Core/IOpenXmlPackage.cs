using DocumentFormat.OpenXml.Packaging;
using OfficeAgent.Abstractions;
using DocFormat = OfficeAgent.Abstractions.DocumentFormat;

namespace OfficeAgent.Core;

/// <summary>
/// A handle to an open, in-memory OOXML (OPC) package. Format modules operate on
/// the typed <see cref="Package"/> while the core engine uses the common package abstraction.
/// </summary>
public interface IOpenXmlPackage : IDisposable
{
    DocFormat Format { get; }

    string MainPartContentType { get; }

    OpenXmlPackage Package { get; }

    bool IsEditable { get; }


    byte[] ToBytes();
}
