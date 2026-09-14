using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeAgent.Abstractions;
using DocFormat = OfficeAgent.Abstractions.DocumentFormat;

namespace OfficeAgent.Core;

/// <summary>
/// Opens bytes as an in-memory Open XML package. Avoids a second copy when the
/// caller already owns the bytes and detects the format from <c>[Content_Types].xml</c>
/// in a single zip pass.
/// </summary>
/// <remarks>
/// Every package this engine opens passes through here, which is why the host ingestion
/// ceilings are applied at this one point rather than at each entry point. A limit that
/// lived in the provider would not cover inline base64; one in the agent adapter would
/// not cover the direct API.
/// </remarks>
internal sealed class OpenXmlPackageService
{
    private readonly OpenXmlIngestionLimits _limits;

    public OpenXmlPackageService() : this(OpenXmlIngestionLimits.Default)
    {
    }

    public OpenXmlPackageService(OpenXmlIngestionLimits limits)
    {
        _limits = limits ?? OpenXmlIngestionLimits.Default;
        _limits.Validate();
    }

    public OpenXmlIngestionLimits Limits => _limits;

    public IOpenXmlPackage Open(byte[] bytes, bool editable)
    {
        if (bytes is null) throw new ArgumentNullException(nameof(bytes));
        if (bytes.LongLength > _limits.MaximumCompressedBytes)
            throw new OpenXmlIngestionLimitException(
                nameof(_limits.MaximumCompressedBytes), _limits.MaximumCompressedBytes, bytes.LongLength);

        var memory = new MemoryStream();
        memory.Write(bytes, 0, bytes.Length);
        memory.Position = 0;
        return OpenOwned(memory, editable);
    }

    public IOpenXmlPackage Open(Stream source, bool editable)
    {
        if (source is null) throw new ArgumentNullException(nameof(source));

        if (source is MemoryStream ms && ms.TryGetBuffer(out var buf))
        {
            if (buf.Count > _limits.MaximumCompressedBytes)
                throw new OpenXmlIngestionLimitException(
                    nameof(_limits.MaximumCompressedBytes), _limits.MaximumCompressedBytes, buf.Count);

            var owned = new byte[buf.Count];
            Array.Copy(buf.Array!, buf.Offset, owned, 0, buf.Count);
            return Open(owned, editable);
        }

        var memory = ReadBounded(source);
        memory.Position = 0;
        return OpenOwned(memory, editable);
    }

    private MemoryStream ReadBounded(Stream source)
    {
        var memory = new MemoryStream();
        CopyBounded(source, memory, _limits.MaximumCompressedBytes);
        return memory;
    }

    /// <summary>
    /// Reads a stream into a byte array under a ceiling. A non-seekable source has no
    /// length to check in advance, so the only honest bound is to stop reading: an endless
    /// stream is refused at the ceiling rather than drained into memory first.
    /// </summary>
    /// <remarks>
    /// Shared with the flow orchestrator, which reads a caller's handle before this
    /// service ever sees the bytes. A ceiling applied only at package open would leave
    /// that earlier read unbounded.
    /// </remarks>
    internal static byte[] ReadBounded(Stream source, long limit)
    {
        using var memory = new MemoryStream();
        CopyBounded(source, memory, limit);
        return memory.ToArray();
    }

    internal static async Task<byte[]> ReadBoundedAsync(
        Stream source, long limit, CancellationToken cancellationToken)
    {
        using var memory = new MemoryStream();
        var buffer = new byte[81920];
        long total = 0;

        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();
            int read = await source.ReadAsync(buffer, 0, buffer.Length, cancellationToken).ConfigureAwait(false);
            if (read <= 0) break;

            total += read;
            if (total > limit) throw Exceeded(limit);

            memory.Write(buffer, 0, read);
        }

        return memory.ToArray();
    }

    private static void CopyBounded(Stream source, MemoryStream destination, long limit)
    {
        // A seekable source knows its own size, so the refusal can name it exactly
        // instead of reporting that the limit was simply passed.
        if (source.CanSeek && source.Length > limit)
        {
            destination.Dispose();
            throw new OpenXmlIngestionLimitException("MaximumCompressedBytes", limit, source.Length);
        }

        var buffer = new byte[81920];
        long total = 0;

        while (true)
        {
            int read = source.Read(buffer, 0, buffer.Length);
            if (read <= 0) break;

            total += read;
            if (total > limit)
            {
                destination.Dispose();
                throw Exceeded(limit);
            }

            destination.Write(buffer, 0, read);
        }
    }

    private static OpenXmlIngestionLimitException Exceeded(long limit) =>
        new("MaximumCompressedBytes", limit, observed: -1,
            detail: "The source stream was still producing bytes when the limit was reached.");

    /// <summary>
    /// The failures a hostile or damaged package produces. Each is translated to one
    /// stable refusal rather than surfacing the SDK's own type and message, which name
    /// internal part URIs and content types the caller has no use for.
    /// </summary>
    internal static bool IsPackageFault(Exception ex) =>
        ex is InvalidDataException
           or FileFormatException
           or System.Xml.XmlException
           or OpenXmlPackageException;

    /// <summary>
    /// Runs work that touches package parts, translating a damaged-package fault into the
    /// one stable refusal. Parts load lazily, so a corrupt relationship or style part can
    /// fail during an inspect rather than at the open, and the caller should still be told
    /// the same thing about the same document.
    /// </summary>
    internal static T Guarded<T>(Func<T> work)
    {
        try
        {
            return work();
        }
        catch (Exception ex) when (IsPackageFault(ex))
        {
            throw new OpenXmlPackageRejectedException(
                "The package could not be read: one of its parts is corrupt or not readable OOXML.", ex);
        }
    }

    /// <summary>Forces the format's main part to parse, so lazy faults are not deferred.</summary>
    private static void ForceMainPart(OpenXmlPackage package, DocFormat format)
    {
        switch (format)
        {
            case DocFormat.Word:
                _ = ((WordprocessingDocument)package).MainDocumentPart?.Document?.Body;
                break;
            case DocFormat.Excel:
                _ = ((SpreadsheetDocument)package).WorkbookPart?.Workbook?.Sheets;
                break;
            case DocFormat.PowerPoint:
                _ = ((PresentationDocument)package).PresentationPart?.Presentation?.SlideIdList;
                break;
        }
    }

    public void Save(IOpenXmlPackage package, Stream destination)
    {
        var bytes = package.ToBytes();
        destination.Write(bytes, 0, bytes.Length);
    }

    private IOpenXmlPackage OpenOwned(MemoryStream memory, bool editable)
    {
        var format = Inspect(memory);
        memory.Position = 0;

        OpenXmlPackage package;
        try
        {
            package = format switch
            {
                DocFormat.Word => WordprocessingDocument.Open(memory, editable),
                DocFormat.Excel => SpreadsheetDocument.Open(memory, editable),
                DocFormat.PowerPoint => PresentationDocument.Open(memory, editable),
                _ => throw new NotSupportedException($"Unsupported format: {format}.")
            };

            // The SDK loads parts lazily, so a corrupt main part would otherwise surface
            // as a raw XmlException from somewhere inside a later inspect or apply, long
            // past any place that could describe it. Touching it here keeps the refusal
            // at the boundary where the package is admitted.
            ForceMainPart(package, format);
        }
        catch (Exception ex) when (IsPackageFault(ex))
        {
            throw new OpenXmlPackageRejectedException(
                "The package could not be opened: it is corrupt, truncated, or not a readable OOXML file.", ex);
        }

        return new OpenXmlPackageHandle(memory, package, format, editable);
    }

    /// <summary>
    /// One zip pass that both bounds the package and determines its format. Entry
    /// metadata is not trusted for size: <see cref="ZipArchiveEntry.Length"/> is written
    /// by whoever produced the archive, so declared sizes are checked and the parts that
    /// matter are measured while being read.
    /// </summary>
    private DocFormat Inspect(Stream stream)
    {
        stream.Position = 0;
        long compressed = stream.Length;

        ZipArchive zip;
        try
        {
            zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
        }
        catch (InvalidDataException ex)
        {
            throw new OpenXmlPackageRejectedException(
                "The package is not a readable archive: it is corrupt or truncated.", ex);
        }

        using (zip)
        {
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            long declaredExpanded = 0;
            int count = 0;

            foreach (var entry in zip.Entries)
            {
                count++;
                if (count > _limits.MaximumParts)
                    throw new OpenXmlIngestionLimitException(
                        nameof(_limits.MaximumParts), _limits.MaximumParts, observed: -1,
                        detail: "Entry enumeration stopped at the limit.");

                RejectUnsafeName(entry.FullName);

                if (!names.Add(entry.FullName))
                    throw new OpenXmlPackageRejectedException(
                        $"The package contains more than one entry named '{entry.FullName}'. " +
                        "A duplicated part name makes the package ambiguous and it is refused.");

                if (entry.Length > _limits.MaximumPartBytes)
                    throw new OpenXmlIngestionLimitException(
                        nameof(_limits.MaximumPartBytes), _limits.MaximumPartBytes, entry.Length,
                        detail: $"Part '{entry.FullName}' declares more than the per-part limit.");

                declaredExpanded += entry.Length;
                if (declaredExpanded > _limits.MaximumExpandedBytes)
                    throw new OpenXmlIngestionLimitException(
                        nameof(_limits.MaximumExpandedBytes), _limits.MaximumExpandedBytes, declaredExpanded);
            }

            // A decompression bomb declares its true expanded size honestly and relies on
            // the reader to expand it anyway. The ratio check catches that shape before any
            // part is read, without assuming the declared sizes are truthful: they are only
            // ever used to refuse, never to admit.
            if (compressed > 0 && declaredExpanded / Math.Max(compressed, 1) > _limits.MaximumExpansionRatio)
                throw new OpenXmlIngestionLimitException(
                    nameof(_limits.MaximumExpansionRatio), _limits.MaximumExpansionRatio,
                    declaredExpanded / Math.Max(compressed, 1),
                    detail: $"The package declares {declaredExpanded} expanded bytes from {compressed} compressed.");

            var contentTypes = zip.GetEntry("[Content_Types].xml")
                ?? throw new OpenXmlPackageRejectedException(
                    "Not an OOXML package: the required '[Content_Types].xml' part is missing.");

            return DetectFormat(contentTypes);
        }
    }

    /// <summary>
    /// Refuses an entry name that escapes the package root or names an absolute location.
    /// OPC part names are relative; anything else is a traversal attempt aimed at whatever
    /// extracts the archive next.
    /// </summary>
    private static void RejectUnsafeName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new OpenXmlPackageRejectedException("The package contains an entry with no name.");

        if (name.IndexOf('\0') >= 0)
            throw new OpenXmlPackageRejectedException("The package contains an entry name with an embedded null.");

        if (name.StartsWith("/", StringComparison.Ordinal) ||
            name.StartsWith("\\", StringComparison.Ordinal) ||
            (name.Length > 1 && name[1] == ':'))
            throw new OpenXmlPackageRejectedException(
                $"The package contains an absolute entry name '{name}'. Part names must be relative.");

        var segments = name.Split('/', '\\');
        if (segments.Any(segment => segment == ".."))
            throw new OpenXmlPackageRejectedException(
                $"The package contains a traversing entry name '{name}'. Part names must stay inside the package.");
    }

    /// <summary>
    /// Reads the content-types manifest under explicit reader limits. DTD processing is
    /// prohibited, so no external entity or entity-expansion payload is resolved, and the
    /// part is bounded by character count and nesting depth.
    /// </summary>
    private DocFormat DetectFormat(ZipArchiveEntry contentTypes)
    {
        using var raw = contentTypes.Open();
        using var counting = new CountingStream(raw, _limits.MaximumXmlCharacters, "[Content_Types].xml");

        var settings = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            MaxCharactersInDocument = _limits.MaximumXmlCharacters,
            MaxCharactersFromEntities = 0,
            IgnoreWhitespace = true,
            CloseInput = false
        };

        XDocument manifest;
        try
        {
            using var reader = XmlReader.Create(counting, settings);
            manifest = XDocument.Load(reader);
        }
        catch (OpenXmlIngestionLimitException)
        {
            throw;
        }
        catch (XmlException ex)
        {
            throw new OpenXmlPackageRejectedException(
                "The package's '[Content_Types].xml' is not XML this engine will process. " +
                "Document type definitions and external entities are refused.", ex);
        }

        int depth = Depth(manifest.Root);
        if (depth > _limits.MaximumXmlDepth)
            throw new OpenXmlIngestionLimitException(
                nameof(_limits.MaximumXmlDepth), _limits.MaximumXmlDepth, depth,
                detail: "Nesting in '[Content_Types].xml' is deeper than the limit.");

        var mainTypes = manifest.Root?.Elements()
            .Select(element => (string?)element.Attribute("ContentType"))
            .Where(contentType => contentType?.EndsWith(".main+xml", StringComparison.OrdinalIgnoreCase) == true)
            .ToList() ?? new List<string?>();

        if (mainTypes.Any(type => type!.Contains("wordprocessingml"))) return DocFormat.Word;
        if (mainTypes.Any(type => type!.Contains("spreadsheetml"))) return DocFormat.Excel;
        if (mainTypes.Any(type => type!.Contains("presentationml"))) return DocFormat.PowerPoint;

        throw new OpenXmlPackageRejectedException(
            "The package declares no recognised Word, Excel, or PowerPoint main part.");
    }

    private static int Depth(XElement? element)
    {
        if (element is null) return 0;
        int deepest = 0;
        foreach (var child in element.Elements())
            deepest = Math.Max(deepest, Depth(child));
        return deepest + 1;
    }

    /// <summary>
    /// Stops a part from expanding past its ceiling while it is being read. The archive's
    /// declared length is checked first, but a declared length can lie; this measures.
    /// </summary>
    private sealed class CountingStream : Stream
    {
        private readonly Stream _inner;
        private readonly long _limit;
        private readonly string _part;
        private long _read;

        public CountingStream(Stream inner, long limit, string part)
        {
            _inner = inner;
            _limit = limit;
            _part = part;
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            int read = _inner.Read(buffer, offset, count);
            _read += read;
            if (_read > _limit)
                throw new OpenXmlIngestionLimitException(
                    "MaximumXmlCharacters", _limit, _read,
                    detail: $"Part '{_part}' expanded past the limit while being read.");
            return read;
        }

        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position
        {
            get => _read;
            set => throw new NotSupportedException();
        }
        public override void Flush() => _inner.Flush();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        protected override void Dispose(bool disposing)
        {
            if (disposing) _inner.Dispose();
            base.Dispose(disposing);
        }
    }
}
