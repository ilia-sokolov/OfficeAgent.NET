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

    /// <summary>Parses the <c>slide#{id}</c> path form, tolerating a bare id.</summary>
    internal static bool TryParseSlideId(string path, out uint slideId)
    {
        slideId = 0;
        if (string.IsNullOrEmpty(path)) return false;

        var value = path.StartsWith("slide#", StringComparison.OrdinalIgnoreCase)
            ? path.Substring("slide#".Length)
            : path;
        return uint.TryParse(value, out slideId);
    }
}
