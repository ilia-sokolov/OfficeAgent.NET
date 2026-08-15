using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeAgent.PowerPoint;

/// <summary>
/// Surfaces every shape on a slide as an addressable node.
/// </summary>
/// <remarks>
/// The table and image providers address the two kinds a plan manipulates by content;
/// this one addresses <em>any</em> shape by identity, which is what moving, resizing, and
/// deleting need. A shape therefore appears under two paths - <c>image#…</c> and
/// <c>shape#…</c> for a picture, say - because the verbs care about different things: one
/// about the picture, the other about the box it sits in.
/// </remarks>
internal sealed class ShapeNodeProvider : IPowerPointNodeProvider
{
    /// <inheritdoc />
    public string Kind => "shape";

    /// <inheritdoc />
    public IEnumerable<NodeInfo> Enumerate(PowerPointObjectMap map)
    {
        foreach (var slide in PowerPointModel.Slides(map.Package))
        {
            var tree = slide.Part.Slide.CommonSlideData?.ShapeTree;
            if (tree is null) continue;

            foreach (var element in tree.ChildElements)
            {
                var shapeId = PowerPointModel.ShapeIdOf(element);
                if (shapeId is null) continue;

                var path = $"shape#{slide.SlideId}/{shapeId}";
                yield return new NodeInfo
                {
                    Kind = Kind,
                    Path = path,
                    Summary = Describe(slide, element),
                    Anchor = new NodeAnchor { Id = path, Kind = Kind, Path = path }
                };
            }
        }
    }

    /// <inheritdoc />
    public ResolvedNode? Resolve(NodeAnchor anchor, PowerPointObjectMap map)
    {
        var located = Locate(anchor.Path, map.Package);
        return located is null
            ? null
            : new ResolvedNode
            {
                Kind = Kind,
                Elements = new[] { located.Element },
                Value = anchor.Path
            };
    }

    /// <summary>Finds a shape by its <c>shape#{slideId}/{shapeId}</c> path.</summary>
    public static LocatedShape? Locate(string path, IOpenXmlPackage package)
    {
        if (!TryParse(path, out var slideId, out var shapeId)) return null;

        var slide = PowerPointModel.Slide(package, slideId);
        var tree = slide?.Part.Slide.CommonSlideData?.ShapeTree;
        if (slide is null || tree is null) return null;

        foreach (var element in tree.ChildElements)
            if (PowerPointModel.ShapeIdOf(element) == shapeId)
                return new LocatedShape(slide, element);

        return null;
    }

    internal static bool TryParse(string path, out uint slideId, out uint shapeId)
    {
        slideId = 0;
        shapeId = 0;
        if (string.IsNullOrEmpty(path)) return false;

        var value = path.StartsWith("shape#", StringComparison.OrdinalIgnoreCase)
            ? path.Substring("shape#".Length)
            : path;

        var slash = value.IndexOf('/');
        if (slash <= 0) return false;

        return uint.TryParse(value.Substring(0, slash), out slideId)
            && uint.TryParse(value.Substring(slash + 1), out shapeId);
    }

    /// <summary>
    /// Whether the shape is a layout placeholder. Deleting one is refused: the layout
    /// re-offers it as an empty prompt, so the slide looks unchanged while its content is
    /// gone - the worst kind of edit, one that appears not to have happened.
    /// </summary>
    public static bool IsPlaceholder(OpenXmlElement element) =>
        element is Shape s &&
        s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape is not null;

    private static string Describe(SlideRef slide, OpenXmlElement element)
    {
        var kind = element switch
        {
            Shape when IsPlaceholder(element) => "placeholder",
            Shape => "text box",
            GraphicFrame => "table",
            Picture => "picture",
            _ => element.LocalName
        };

        var name = PowerPointModel.ShapeNameOf(element);
        var text = string.Concat(element.Descendants<A.Text>().Select(t => t.Text));
        if (text.Length > 40) text = text.Substring(0, 40) + "…";

        var label = name.Length > 0 ? $"{kind} '{name}'" : kind;
        return text.Length > 0
            ? $"slide {slide.Number}: {label} — \"{text}\""
            : $"slide {slide.Number}: {label}";
    }
}

/// <summary>One shape and the slide it lives on.</summary>
internal sealed class LocatedShape
{
    public LocatedShape(SlideRef slide, OpenXmlElement element)
    {
        Slide = slide;
        Element = element;
    }

    public SlideRef Slide { get; }
    public OpenXmlElement Element { get; }

    /// <summary>
    /// Moves and resizes the shape, creating the transform when it has none.
    /// </summary>
    /// <remarks>
    /// A graphic frame keeps its transform in <c>p:xfrm</c> and everything else in
    /// <c>a:xfrm</c> under <c>p:spPr</c> - different elements with the same content model,
    /// so the two are handled together here rather than at every call site. A placeholder
    /// legitimately has no transform at all, inheriting one from its layout; writing one
    /// is what pins it, which is exactly what moving or resizing it means.
    /// </remarks>
    public void Arrange(long? x, long? y, long? cx, long? cy)
    {
        if (Element is GraphicFrame frame)
        {
            var transform = frame.Transform ??= new Transform();
            Apply(
                transform.Offset ??= new A.Offset { X = 0, Y = 0 },
                transform.Extents ??= new A.Extents { Cx = 0, Cy = 0 });

            // A table renders at the width of its own column grid, not the frame's, so a
            // frame widened on its own leaves the table the size it was and the caller
            // sees nothing happen. PowerPoint rescales the columns when a user drags the
            // handle; doing the same here is what makes the resize mean anything.
            if (cx is { } frameWidth) RescaleColumns(frame, frameWidth);
            return;
        }

        var properties = Element switch
        {
            Shape s => s.ShapeProperties ??= new ShapeProperties(),
            Picture p => p.ShapeProperties ??= new ShapeProperties(),
            _ => throw new InvalidOperationException(
                $"Shape kind '{Element.LocalName}' has no transform to change.")
        };

        var xfrm = properties.Transform2D ??= new A.Transform2D();
        Apply(
            xfrm.Offset ??= new A.Offset { X = 0, Y = 0 },
            xfrm.Extents ??= new A.Extents { Cx = 0, Cy = 0 });

        void Apply(A.Offset offset, A.Extents extents)
        {
            if (x is { } left) offset.X = left;
            if (y is { } top) offset.Y = top;
            if (cx is { } width) extents.Cx = width;
            if (cy is { } height) extents.Cy = height;
        }
    }

    /// <summary>
    /// Spreads a new frame width across a table's columns, keeping their relative
    /// proportions. Rounding is absorbed by the last column so the grid still totals the
    /// frame width exactly - a grid that is a few EMUs out draws a visible seam.
    /// </summary>
    private static void RescaleColumns(GraphicFrame frame, long frameWidth)
    {
        var columns = frame.Descendants<A.GridColumn>().ToList();
        if (columns.Count == 0 || frameWidth <= 0) return;

        var current = columns.Sum(c => c.Width?.Value ?? 0L);
        long used = 0;

        for (var i = 0; i < columns.Count - 1; i++)
        {
            var width = current > 0
                ? (long)(frameWidth * ((double)(columns[i].Width?.Value ?? 0L) / current))
                : frameWidth / columns.Count;

            columns[i].Width = width;
            used += width;
        }

        columns[columns.Count - 1].Width = frameWidth - used;
    }

    /// <summary>The shape's current box in EMUs, or null where it inherits one.</summary>
    public (long X, long Y, long Cx, long Cy)? Box()
    {
        var offset = Element switch
        {
            GraphicFrame f => f.Transform?.Offset,
            Shape s => s.ShapeProperties?.Transform2D?.Offset,
            Picture p => p.ShapeProperties?.Transform2D?.Offset,
            _ => null
        };
        var extents = Element switch
        {
            GraphicFrame f => f.Transform?.Extents,
            Shape s => s.ShapeProperties?.Transform2D?.Extents,
            Picture p => p.ShapeProperties?.Transform2D?.Extents,
            _ => null
        };

        return offset is null || extents is null
            ? null
            : (offset.X?.Value ?? 0, offset.Y?.Value ?? 0, extents.Cx?.Value ?? 0, extents.Cy?.Value ?? 0);
    }
}
