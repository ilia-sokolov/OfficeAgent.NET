using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace OfficeAgent.Word;

/// <summary>
/// Surfaces inline images (drawings with a <c>pic:pic</c> reference) as nodes
/// addressable by <c>image#N</c>, allowing them to be removed via
/// <see cref="RemoveImageOp"/>. Floating shapes and SmartArt are not enumerated.
/// </summary>
internal sealed class ImageNodeProvider : IWordNodeProvider
{
    public string Kind => "image";

    public IEnumerable<NodeInfo> Enumerate(WordObjectMap map)
    {
        int index = 0;
        foreach (var drawing in EnumerateDrawings(map.Package))
        {
            var blip = drawing.Descendants<A.Blip>().FirstOrDefault();
            var name = drawing.Descendants<PIC.NonVisualDrawingProperties>().FirstOrDefault()?.Name?.Value
                       ?? drawing.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>().FirstOrDefault()?.Description?.Value
                       ?? "(image)";
            yield return new NodeInfo
            {
                Kind = Kind,
                Path = $"image#{index}",
                Summary = $"Image {index}: {name}{(blip?.Embed?.Value is string rid ? $" (rId={rid})" : "")}",
                Anchor = new NodeAnchor { Id = $"img:{index}", Kind = Kind, Path = $"image#{index}" }
            };
            index++;
        }
    }

    public ResolvedNode? Resolve(NodeAnchor anchor, WordObjectMap map)
    {
        var located = ResolveWithHost(anchor, map);
        return located is null
            ? null
            : new ResolvedNode { Kind = Kind, Elements = new OpenXmlElement[] { located.Value.Drawing } };
    }

    /// <summary>
    /// Resolves an <c>image#N</c> anchor to the live drawing together with the
    /// <see cref="OpenXmlPart"/> that hosts it. The hosting part owns the blip's
    /// relationship to the underlying <see cref="ImagePart"/>, so callers that need to
    /// release that resource (see <see cref="RemoveImageHandler"/>) must address the
    /// part the drawing actually lives in - body, header, footer, footnotes, or endnotes.
    /// </summary>
    internal (Drawing Drawing, OpenXmlPart Host)? ResolveWithHost(NodeAnchor anchor, WordObjectMap map)
    {
        if (!anchor.Path.StartsWith("image#", StringComparison.Ordinal)) return null;
        if (!int.TryParse(anchor.Path.Substring("image#".Length), out var target)) return null;

        int index = 0;
        foreach (var located in EnumerateDrawingsWithHost(map.Package))
        {
            if (index == target) return located;
            index++;
        }
        return null;
    }

    internal static IEnumerable<Drawing> EnumerateDrawings(IOpenXmlPackage package) =>
        EnumerateDrawingsWithHost(package).Select(located => located.Drawing);

    /// <summary>
    /// Enumerates inline drawings paired with the part that hosts them, in the same
    /// order (body → headers → footers → footnotes → endnotes) the <c>image#N</c> paths
    /// are assigned, so an index resolved here matches the one surfaced by inspection.
    /// </summary>
    internal static IEnumerable<(Drawing Drawing, OpenXmlPart Host)> EnumerateDrawingsWithHost(IOpenXmlPackage package)
    {
        var main = WordModel.Main(package);

        foreach (var located in DrawingsIn(main, main.Document?.Body)) yield return located;
        foreach (var header in main.HeaderParts)
            foreach (var located in DrawingsIn(header, header.Header)) yield return located;
        foreach (var footer in main.FooterParts)
            foreach (var located in DrawingsIn(footer, footer.Footer)) yield return located;
        if (main.FootnotesPart is { } footnotes)
            foreach (var located in DrawingsIn(footnotes, footnotes.Footnotes)) yield return located;
        if (main.EndnotesPart is { } endnotes)
            foreach (var located in DrawingsIn(endnotes, endnotes.Endnotes)) yield return located;
    }

    private static IEnumerable<(Drawing Drawing, OpenXmlPart Host)> DrawingsIn(OpenXmlPart host, OpenXmlElement? root)
    {
        if (root is null) yield break;
        foreach (var drawing in root.Descendants<Drawing>())
            yield return (drawing, host);
    }
}
