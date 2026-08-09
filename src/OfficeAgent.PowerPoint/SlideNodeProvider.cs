using OfficeAgent.Abstractions;

namespace OfficeAgent.PowerPoint;

/// <summary>
/// Surfaces each slide as an addressable node, so inspection lists the deck's structure
/// even at a fidelity that omits paragraph text, and so later verbs have a slide to
/// target without inventing an anchor vocabulary.
/// </summary>
internal sealed class SlideNodeProvider : IPowerPointNodeProvider
{
    /// <inheritdoc />
    public string Kind => "slide";

    /// <inheritdoc />
    public IEnumerable<NodeInfo> Enumerate(PowerPointObjectMap map)
    {
        foreach (var slide in PowerPointModel.Slides(map.Package))
        {
            var shapes = 0;
            foreach (var host in PowerPointModel.TextHosts(slide))
                if (!host.IsNotes) shapes++;

            var hasNotes = slide.Part.NotesSlidePart?.NotesSlide is not null;

            yield return new NodeInfo
            {
                Kind = Kind,
                Path = $"slide#{slide.SlideId}",
                Summary = $"slide {slide.Number}: {shapes} text body/bodies{(hasNotes ? ", has notes" : "")}",
                Anchor = new NodeAnchor
                {
                    Id = $"slide#{slide.SlideId}",
                    Kind = Kind,
                    Path = $"slide#{slide.SlideId}"
                }
            };
        }
    }

    /// <inheritdoc />
    public ResolvedNode? Resolve(NodeAnchor anchor, PowerPointObjectMap map)
    {
        if (!TryParseSlideId(anchor.Path, out var slideId)) return null;

        var slide = PowerPointModel.Slide(map.Package, slideId);
        if (slide is null) return null;

        return new ResolvedNode
        {
            Kind = Kind,
            Elements = new[] { (DocumentFormat.OpenXml.OpenXmlElement)slide.Part.Slide },
            Value = slide.Number.ToString()
        };
    }

    /// <summary>
    /// Parses a slide id from any of the forms the deck vocabulary uses: the node path
    /// <c>slide#256</c>, the paragraph-id prefix <c>slide256</c>, or a bare <c>256</c>.
    /// All three appear in ids the engine itself hands out, so accepting only one of them
    /// would fail an agent that copied the prefix from a paragraph id.
    /// </summary>
    internal static bool TryParseSlideId(string path, out uint slideId)
    {
        slideId = 0;
        if (string.IsNullOrEmpty(path)) return false;

        var value = path;
        if (value.StartsWith("slide#", StringComparison.OrdinalIgnoreCase))
            value = value.Substring("slide#".Length);
        else if (value.StartsWith("slide", StringComparison.OrdinalIgnoreCase))
            value = value.Substring("slide".Length);

        return uint.TryParse(value, out slideId);
    }
}
