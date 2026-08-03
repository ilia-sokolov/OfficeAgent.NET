using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeAgent.PowerPoint;

/// <summary>
/// Shared PowerPoint helpers: slide enumeration in presentation order, the text hosts
/// inside a slide (shapes, table cells, and the slide's notes), paragraph addressing,
/// and the snapshot token. Inspect, find, and apply all walk the deck through here so
/// an anchor issued by one is understood by the others.
/// </summary>
internal static class PowerPointModel
{
    public static readonly PresentationmlDialect Dialect = new();
    public static readonly TextBodyEngine Text = new(Dialect);

    public static PresentationDocument Doc(IOpenXmlPackage package) =>
        (PresentationDocument)package.Package;

    public static PresentationPart Main(IOpenXmlPackage package) =>
        Doc(package).PresentationPart
        ?? throw new InvalidOperationException("Presentation has no presentation part.");

    /// <summary>
    /// Enumerates slides in the order the deck presents them, paired with the slide id
    /// PowerPoint stores in <c>p:sldIdLst</c>. That id is durable across reordering and
    /// insertion, which a positional slide number is not, so it is what anchors carry.
    /// </summary>
    public static IEnumerable<SlideRef> Slides(IOpenXmlPackage package)
    {
        var main = Main(package);
        var list = main.Presentation?.SlideIdList;
        if (list is null) yield break;

        var number = 1;
        foreach (var entry in list.Elements<SlideId>())
        {
            var relationshipId = entry.RelationshipId?.Value;
            if (string.IsNullOrEmpty(relationshipId)) continue;
            if (main.GetPartById(relationshipId!) is not SlidePart part) continue;

            yield return new SlideRef(entry.Id?.Value ?? 0, number++, part);
        }
    }

    /// <summary>Resolves one slide by the id an anchor carries.</summary>
    public static SlideRef? Slide(IOpenXmlPackage package, uint slideId)
    {
        foreach (var slide in Slides(package))
            if (slide.SlideId == slideId) return slide;
        return null;
    }

    /// <summary>
    /// Enumerates every addressable text body on a slide: each shape's body, each table
    /// cell's body, and - when present - the slide's notes. The host key is what makes a
    /// paragraph id unique, so it encodes where the body lives rather than a bare index.
    /// </summary>
    public static IEnumerable<TextHost> TextHosts(SlideRef slide)
    {
        foreach (var host in TextHostsIn(slide.Part.Slide, notes: false))
            yield return host;

        if (slide.Part.NotesSlidePart?.NotesSlide is { } notes)
            foreach (var host in TextHostsIn(notes, notes: true))
                yield return host;
    }

    private static IEnumerable<TextHost> TextHostsIn(OpenXmlElement root, bool notes)
    {
        // Shapes carry their own text body; a table's text lives per cell, and a cell is
        // addressed by its position so an anchor survives edits to neighbouring cells.
        foreach (var shape in root.Descendants<Shape>())
        {
            var shapeId = ShapeIdOf(shape);
            if (shapeId is null || shape.TextBody is null) continue;
            yield return new TextHost(
                notes ? $"notes/shape{shapeId}" : $"shape{shapeId}",
                shape.TextBody,
                notes);
        }

        foreach (var frame in root.Descendants<GraphicFrame>())
        {
            var frameId = ShapeIdOf(frame);
            var table = frame.Graphic?.GraphicData?.GetFirstChild<A.Table>();
            if (frameId is null || table is null) continue;

            var rowIndex = 0;
            foreach (var row in table.Elements<A.TableRow>())
            {
                var columnIndex = 0;
                foreach (var cell in row.Elements<A.TableCell>())
                {
                    if (cell.TextBody is { } body)
                        yield return new TextHost(
                            (notes ? "notes/" : "") + $"shape{frameId}/r{rowIndex}c{columnIndex}",
                            body,
                            notes);
                    columnIndex++;
                }
                rowIndex++;
            }
        }
    }

    /// <summary>
    /// Enumerates every paragraph in the deck with the id anchors use.
    /// </summary>
    /// <remarks>
    /// The id reads <c>slide{slideId}/shape{shapeId}/p{index}</c> - or
    /// <c>…/r{row}c{col}/p{index}</c> inside a table, and <c>notes/…</c> for a notes body.
    /// DrawingML has no durable per-paragraph identifier the way WordprocessingML's
    /// <c>w14:paraId</c> is, so the trailing index is positional within its own text body.
    /// It is stable against every edit this module performs, because none of them add or
    /// remove paragraphs inside a body.
    /// </remarks>
    public static IEnumerable<ParagraphRef> Paragraphs(IOpenXmlPackage package)
    {
        foreach (var slide in Slides(package))
            foreach (var paragraph in Paragraphs(slide))
                yield return paragraph;
    }

    /// <summary>Enumerates the paragraphs of one slide.</summary>
    public static IEnumerable<ParagraphRef> Paragraphs(SlideRef slide)
    {
        foreach (var host in TextHosts(slide))
        {
            var index = 0;
            foreach (var paragraph in host.Body.Elements<A.Paragraph>())
            {
                yield return new ParagraphRef(
                    $"slide{slide.SlideId}/{host.Key}/p{index}",
                    paragraph,
                    slide,
                    host);
                index++;
            }
        }
    }

    /// <summary>Resolves a plan anchor's paragraph, or <see langword="null"/> when it no longer exists.</summary>
    public static ParagraphRef? ResolveParagraph(IOpenXmlPackage package, string paraId)
    {
        foreach (var paragraph in Paragraphs(package))
            if (string.Equals(paragraph.ParaId, paraId, StringComparison.Ordinal))
                return paragraph;
        return null;
    }

    /// <summary>Resolves a plan anchor's paragraph, applying the stabilization alias map first.</summary>
    public static ParagraphRef? ResolveParagraph(ApplyContext context, string anchorParaId) =>
        ResolveParagraph(context.Package, context.ResolveAlias(anchorParaId));

    /// <summary>Reads the logical text of a paragraph, joined across its runs.</summary>
    public static string TextOf(A.Paragraph paragraph) => Text.GetLogicalText(paragraph);

    /// <summary>The shape id PowerPoint assigns in <c>p:cNvPr</c>, unique within a slide.</summary>
    public static uint? ShapeIdOf(OpenXmlElement shape) => shape switch
    {
        Shape s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value,
        GraphicFrame f => f.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Id?.Value,
        Picture p => p.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value,
        _ => null
    };

    /// <summary>The name an author gave a shape, used to make inspection readable.</summary>
    public static string ShapeNameOf(OpenXmlElement shape) => shape switch
    {
        Shape s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? string.Empty,
        GraphicFrame f => f.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value ?? string.Empty,
        Picture p => p.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value ?? string.Empty,
        _ => string.Empty
    };

    /// <summary>
    /// Allocates a shape id that no shape on the slide is using. PowerPoint requires
    /// uniqueness within a slide and silently misbehaves when it is violated.
    /// </summary>
    public static uint NextShapeId(SlidePart part)
    {
        uint highest = 1;
        foreach (var element in part.Slide.Descendants<NonVisualDrawingProperties>())
            if (element.Id?.Value is { } id && id > highest) highest = id;
        return highest + 1;
    }

    /// <summary>
    /// Snapshot etag over every slide's XML plus its notes, so drift anywhere the module
    /// can address is detected - not merely drift on the first slide.
    /// </summary>
    public static SnapshotToken Snapshot(IOpenXmlPackage package)
    {
        var builder = new StringBuilder();
        foreach (var slide in Slides(package))
        {
            builder.Append(slide.SlideId).Append('\n');
            builder.Append(slide.Part.Slide.OuterXml).Append('\n');
            if (slide.Part.NotesSlidePart?.NotesSlide is { } notes)
                builder.Append(notes.OuterXml).Append('\n');
        }

        byte[] hash;
        using (var sha = SHA256.Create())
            hash = sha.ComputeHash(Encoding.UTF8.GetBytes(builder.ToString()));
        return new SnapshotToken(BitConverter.ToString(hash).Replace("-", ""));
    }

    /// <summary>Renders a short excerpt around a match, for find results and previews.</summary>
    public static string Snippet(string text, int start, int length, int pad = 32)
    {
        var from = Math.Max(0, start - pad);
        var to = Math.Min(text.Length, start + length + pad);
        var slice = text.Substring(from, to - from);
        return (from > 0 ? "…" : "") + slice + (to < text.Length ? "…" : "");
    }

    public static StringComparison Comparison(bool caseSensitive) =>
        caseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
}

/// <summary>One slide, with the durable id anchors carry and its presentation order.</summary>
internal sealed class SlideRef
{
    public SlideRef(uint slideId, int number, SlidePart part)
    {
        SlideId = slideId;
        Number = number;
        Part = part;
    }

    /// <summary>The id from <c>p:sldIdLst</c>; durable across reordering.</summary>
    public uint SlideId { get; }

    /// <summary>The 1-based position in the deck, for human-readable output only.</summary>
    public int Number { get; }

    public SlidePart Part { get; }
}

/// <summary>One addressable text body and the key that identifies it within its slide.</summary>
internal sealed class TextHost
{
    public TextHost(string key, OpenXmlElement body, bool isNotes)
    {
        Key = key;
        Body = body;
        IsNotes = isNotes;
    }

    /// <summary>For example <c>shape5</c>, <c>shape7/r0c1</c>, or <c>notes/shape3</c>.</summary>
    public string Key { get; }

    public OpenXmlElement Body { get; }

    /// <summary>Whether this body lives on the slide's notes rather than the slide itself.</summary>
    public bool IsNotes { get; }
}

/// <summary>One paragraph with the anchor id that addresses it and where it lives.</summary>
internal sealed class ParagraphRef
{
    public ParagraphRef(string paraId, A.Paragraph paragraph, SlideRef slide, TextHost host)
    {
        ParaId = paraId;
        Paragraph = paragraph;
        Slide = slide;
        Host = host;
    }

    public string ParaId { get; }
    public A.Paragraph Paragraph { get; }
    public SlideRef Slide { get; }
    public TextHost Host { get; }

    /// <summary>The agent-facing location: <c>slide</c> or <c>notes</c>.</summary>
    public string Location => Host.IsNotes ? "notes" : "slide";
}
