using System.Diagnostics.CodeAnalysis;
using DocumentFormat.OpenXml.Packaging;
using OfficeAgent.Abstractions;
using DocFormat = OfficeAgent.Abstractions.DocumentFormat;

namespace OfficeAgent.Core;

/// <summary>
/// A handle to an open, in-memory OOXML (OPC) package. Format modules operate on
/// the typed <see cref="Package"/> while the core engine uses the common package abstraction.
/// </summary>
[Experimental(EngineExtensibility.DiagnosticId, UrlFormat = EngineExtensibility.UrlFormat)]
public interface IOpenXmlPackage : IDisposable
{
    /// <summary>Gets the document format the package holds.</summary>
    DocFormat Format { get; }

    /// <summary>Gets the content type of the package's main part, which identifies the format.</summary>
    string MainPartContentType { get; }

    /// <summary>Gets the typed Open XML SDK package a format module edits.</summary>
    OpenXmlPackage Package { get; }

    /// <summary>Gets whether the package was opened for editing.</summary>
    bool IsEditable { get; }

    /// <summary>Saves the package and returns its bytes.</summary>
    byte[] ToBytes();
}
