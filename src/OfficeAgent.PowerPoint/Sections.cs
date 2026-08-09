using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using OfficeAgent.Core;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace OfficeAgent.PowerPoint;

/// <summary>
/// Reading and maintaining the named slide groups PowerPoint shows in the thumbnail pane.
/// </summary>
/// <remarks>
/// Sections live in an extension on the presentation - <c>p14:sectionLst</c> under
/// <c>p:extLst</c> - and partition the deck in presentation order. Once a deck has any
/// section, every slide belongs to exactly one; a slide listed in no section, or listed in
/// a section but absent from <c>p:sldIdLst</c>, is what makes PowerPoint offer to repair
/// the file. <see cref="Reconcile"/> re-establishes that invariant and is called after
/// every operation that changes the slide list.
/// </remarks>
internal static class Sections
{
    /// <summary>The extension uri PowerPoint stores the section list under.</summary>
    private const string SectionListUri = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}";

    /// <summary>The section list, or null when the deck has no sections.</summary>
    public static P14.SectionList? List(IOpenXmlPackage package) =>
        PowerPointModel.Main(package).Presentation?
            .GetFirstChild<PresentationExtensionList>()?
            .Elements<PresentationExtension>()
            .FirstOrDefault(e => string.Equals(e.Uri?.Value, SectionListUri, StringComparison.OrdinalIgnoreCase))?
            .GetFirstChild<P14.SectionList>();

    /// <summary>The section list, created if the deck has none yet.</summary>
    public static P14.SectionList Require(IOpenXmlPackage package)
    {
        if (List(package) is { } existing) return existing;

        var presentation = PowerPointModel.Main(package).Presentation
            ?? throw new InvalidOperationException("Presentation part has no presentation.");

        // p:extLst is the last child of p:presentation; appending keeps the sequence valid.
        var extensions = presentation.GetFirstChild<PresentationExtensionList>();
        if (extensions is null)
        {
            extensions = new PresentationExtensionList();
            presentation.Append(extensions);
        }

        var list = new P14.SectionList();
        extensions.Append(new PresentationExtension(list) { Uri = SectionListUri });
        return list;
    }

    /// <summary>Sections in deck order, paired with the slides each currently owns.</summary>
    public static IEnumerable<SectionRef> All(IOpenXmlPackage package)
    {
        var list = List(package);
        if (list is null) yield break;

        var number = 1;
        foreach (var section in list.Elements<P14.Section>())
            yield return new SectionRef(section, number++);
    }

    /// <summary>Resolves one section by the id an anchor carries.</summary>
    public static SectionRef? Find(IOpenXmlPackage package, string sectionId) =>
        All(package).FirstOrDefault(s =>
            string.Equals(s.Id, sectionId, StringComparison.OrdinalIgnoreCase));

    /// <summary>Parses the <c>section#{id}</c> path form, tolerating a bare id.</summary>
    public static bool TryParseId(string path, out string sectionId)
    {
        sectionId = string.Empty;
        if (string.IsNullOrEmpty(path)) return false;

        sectionId = path.StartsWith("section#", StringComparison.OrdinalIgnoreCase)
            ? path.Substring("section#".Length)
            : path;
        return sectionId.Length > 0;
    }

    /// <summary>The id PowerPoint expects: a braced, upper-case GUID.</summary>
    public static string NewId() => "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}";

    /// <summary>
    /// Rebuilds every section's slide list from the deck's own order, so the grouping
    /// survives a slide being added, removed, moved, or copied.
    /// </summary>
    /// <remarks>
    /// A slide keeps whichever section already claimed it. One that no section claims -
    /// a slide just inserted or duplicated - joins the section of the slide before it,
    /// which is what PowerPoint does when you add a slide inside a section. Ids for slides
    /// that no longer exist are dropped. A deck with no sections is left alone entirely.
    /// </remarks>
    public static void Reconcile(IOpenXmlPackage package)
    {
        var list = List(package);
        if (list is null) return;

        var sections = list.Elements<P14.Section>().ToList();
        if (sections.Count == 0) return;

        // A section is identified by the slide it starts at, not by everything it lists.
        // That is what keeps a section contiguous: membership follows position, so a slide
        // moved into the middle of another section joins it, exactly as in PowerPoint.
        // Tracking every prior claim instead would let one section own slides either side
        // of another - a shape PresentationML does not allow and PowerPoint will not open.
        var starters = new Dictionary<uint, P14.Section>();
        foreach (var section in sections)
            if (section.SectionSlideIdList?.Elements<P14.SectionSlideIdListEntry>()
                    .FirstOrDefault()?.Id?.Value is { } first)
                starters[first] = section;

        var rebuilt = sections.ToDictionary(s => s, _ => new List<uint>());
        var current = sections[0];

        foreach (var slide in PowerPointModel.Slides(package))
        {
            if (starters.TryGetValue(slide.SlideId, out var starting)) current = starting;
            rebuilt[current].Add(slide.SlideId);
        }

        foreach (var section in sections)
        {
            var entries = section.SectionSlideIdList ??= new P14.SectionSlideIdList();
            entries.RemoveAllChildren();
            foreach (var id in rebuilt[section])
                entries.Append(new P14.SectionSlideIdListEntry { Id = id });
        }

        // PowerPoint presents sections in the order their slides appear, whatever order
        // the list stores them in. Storing them that way too keeps the file canonical and
        // stops inspect_document listing them differently from the thumbnail pane - after
        // a slide is moved across the deck, the two would otherwise disagree.
        var order = PowerPointModel.Slides(package).Select(s => s.SlideId).ToList();
        var position = 0;
        var ranked = new List<(P14.Section Section, int Rank)>();
        foreach (var section in sections)
        {
            var first = rebuilt[section].FirstOrDefault();
            // A section left with no slides keeps its place relative to its neighbours
            // rather than jumping to the front.
            if (rebuilt[section].Count > 0) position = order.IndexOf(first);
            ranked.Add((section, position));
        }

        foreach (var (section, _) in ranked.OrderBy(r => r.Rank))
        {
            section.Remove();
            list.Append(section);
        }
    }
}

/// <summary>One section, its stable id, and its position in the deck.</summary>
internal sealed class SectionRef
{
    public SectionRef(P14.Section section, int number)
    {
        Section = section;
        Number = number;
    }

    public P14.Section Section { get; }

    /// <summary>The 1-based position among sections, for readable output only.</summary>
    public int Number { get; }

    public string Id => Section.Id?.Value ?? string.Empty;

    public string Name => Section.Name?.Value ?? string.Empty;

    public IReadOnlyList<uint> SlideIds =>
        Section.SectionSlideIdList?.Elements<P14.SectionSlideIdListEntry>()
            .Select(e => e.Id?.Value ?? 0U)
            .Where(id => id != 0)
            .ToList()
        ?? (IReadOnlyList<uint>)Array.Empty<uint>();
}
