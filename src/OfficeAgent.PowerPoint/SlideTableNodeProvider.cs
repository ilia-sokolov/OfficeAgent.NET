using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using OfficeAgent.Abstractions;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeAgent.PowerPoint;

/// <summary>
/// Surfaces each table on a slide as an addressable node.
/// </summary>
/// <remarks>
/// A PowerPoint table is an <c>a:tbl</c> wrapped in a <c>p:graphicFrame</c>, and the frame
/// is the thing that carries the shape id. Paths therefore read
/// <c>table#{slideId}/{shapeId}</c> rather than Word's ordinal <c>table#N</c>: an ordinal
/// would silently retarget when a table is added to an earlier slide, whereas the shape id
/// is durable for the life of the shape.
/// </remarks>
internal sealed class SlideTableNodeProvider : IPowerPointNodeProvider
{
    /// <inheritdoc />
    public string Kind => "table";

    /// <inheritdoc />
    public IEnumerable<NodeInfo> Enumerate(PowerPointObjectMap map)
    {
        foreach (var table in Tables(map))
        {
            var rows = table.Table.Elements<A.TableRow>().Count();
            var columns = table.Table.TableGrid?.Elements<A.GridColumn>().Count() ?? 0;

            yield return new NodeInfo
            {
                Kind = Kind,
                Path = table.Path,
                Summary = $"slide {table.Slide.Number}: table {rows}×{columns}",
                Anchor = new NodeAnchor { Id = table.Path, Kind = Kind, Path = table.Path }
            };
        }
    }

    /// <inheritdoc />
    public ResolvedNode? Resolve(NodeAnchor anchor, PowerPointObjectMap map)
    {
        var located = Locate(anchor.Path, map);
        if (located is null) return null;

        return new ResolvedNode
        {
            Kind = Kind,
            Elements = new OpenXmlElement[] { located.Table, located.Frame },
            Value = located.Path
        };
    }

    /// <summary>Finds one table by its path, or null when it is gone.</summary>
    internal static TableRef? Locate(string path, PowerPointObjectMap map)
    {
        foreach (var table in Tables(map))
            if (string.Equals(table.Path, path, StringComparison.OrdinalIgnoreCase))
                return table;
        return null;
    }

    /// <summary>Enumerates every table in the deck, in slide order.</summary>
    internal static IEnumerable<TableRef> Tables(PowerPointObjectMap map)
    {
        foreach (var slide in PowerPointModel.Slides(map.Package))
            foreach (var frame in slide.Part.Slide.Descendants<GraphicFrame>())
            {
                var shapeId = PowerPointModel.ShapeIdOf(frame);
                var table = frame.Graphic?.GraphicData?.GetFirstChild<A.Table>();
                if (shapeId is null || table is null) continue;

                yield return new TableRef($"table#{slide.SlideId}/{shapeId}", table, frame, slide);
            }
    }
}

/// <summary>One table, its frame, and the slide it sits on.</summary>
internal sealed class TableRef
{
    public TableRef(string path, A.Table table, GraphicFrame frame, SlideRef slide)
    {
        Path = path;
        Table = table;
        Frame = frame;
        Slide = slide;
    }

    public string Path { get; }
    public A.Table Table { get; }
    public GraphicFrame Frame { get; }
    public SlideRef Slide { get; }
}
