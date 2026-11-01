using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;

namespace OfficeAgent.Word;

/// <summary>
/// Resolves the style a plan names to a style the document defines.
/// </summary>
/// <remarks>
/// A plan may name a style by its id (<c>Heading2</c>) or by the name Word shows
/// (<c>Heading 2</c>, or a company's <c>Body Text Blue</c>). A name that matches nothing is
/// refused: Word renders a paragraph whose style id it does not know as Normal, without a
/// word, which is how a company template quietly loses its look.
/// </remarks>
internal static class WordStyleNames
{
    /// <summary>One style: its id, the name Word shows, and its type.</summary>
    internal sealed record Entry(string Id, string Name, string Type);

    /// <summary>The styles the document defines, in <c>styles.xml</c> order.</summary>
    public static List<Entry> Catalog(IOpenXmlPackage package)
    {
        var entries = new List<Entry>();
        var styles = WordModel.Doc(package).MainDocumentPart?.StyleDefinitionsPart?.Styles;
        if (styles is null) return entries;

        foreach (var style in styles.Elements<Style>())
        {
            if (style.StyleId?.Value is not { Length: > 0 } id) continue;
            entries.Add(new Entry(id, style.StyleName?.Val?.Value ?? id, TypeOf(style)));
        }

        return entries;
    }

    /// <summary>
    /// The style id a plan's value stands for, or null with the reason in
    /// <paramref name="error"/>. An exact id wins, then the name Word shows, then an id in
    /// another case; each ignores styles of another type when one of the right type matches.
    /// </summary>
    public static string? Resolve(IReadOnlyList<Entry> catalog, string value, string type, out string? error)
    {
        error = null;

        foreach (var match in new Func<Entry, bool>[]
                 {
                     e => string.Equals(e.Id, value, StringComparison.Ordinal),
                     e => string.Equals(e.Name, value, StringComparison.OrdinalIgnoreCase),
                     e => string.Equals(e.Id, value, StringComparison.OrdinalIgnoreCase)
                 })
        {
            var found = catalog.Where(match).ToList();
            if (found.Count == 0) continue;

            if (found.FirstOrDefault(e => e.Type == type) is { } right) return right.Id;

            error = $"'{value}' is a {found[0].Type} style; this needs a {type} style. " +
                Known(catalog, type);
            return null;
        }

        // Word adds a built-in style's definition the first time it is applied; so do we,
        // for the few whose standard definition this module carries.
        if (BuiltIn(value) is { } builtIn && TypeOf(builtIn) == type)
            return builtIn.StyleId!.Value!;

        error = $"The document has no {type} style named '{value}'. " + Known(catalog, type) +
            " Define it first with defineStyle, or use one of these.";
        return null;
    }

    /// <summary>
    /// The id to write for a value already checked by plan validation. Applied in order, a
    /// style an earlier operation defined is in the document by now; a built-in the document
    /// lacks is added; a value that still resolves to nothing is written as given rather
    /// than failing a validated plan.
    /// </summary>
    public static string IdFor(IOpenXmlPackage package, string value, string type)
    {
        var catalog = Catalog(package);
        var id = Resolve(catalog, value, type, out _);
        if (id is null) return value;

        if (!catalog.Any(e => string.Equals(e.Id, id, StringComparison.Ordinal)) && BuiltIn(id) is { } builtIn)
            AddBuiltIn(package, builtIn);
        return id;
    }

    /// <summary>
    /// The standard definition of a built-in style this module carries, by id or by the name
    /// Word shows; null for any other style.
    /// </summary>
    internal static Style? BuiltIn(string value)
    {
        foreach (var style in BuiltIns())
            if (string.Equals(style.StyleId?.Value, value, StringComparison.OrdinalIgnoreCase)
                || string.Equals(style.StyleName?.Val?.Value, value, StringComparison.OrdinalIgnoreCase))
                return style;
        return null;
    }

    private static IEnumerable<Style> BuiltIns()
    {
        yield return WordModule.Normal();
        yield return WordModule.Heading(1, "heading 1", 32, "2F5496");
        yield return WordModule.Heading(2, "heading 2", 26, "2F5496");
        yield return WordModule.Heading(3, "heading 3", 24, "1F3763");
        yield return WordModule.ListParagraph();
        yield return WordModule.TableGrid();
    }

    /// <summary>Adds a built-in definition, and Normal first when it is based on Normal.</summary>
    private static void AddBuiltIn(IOpenXmlPackage package, Style style)
    {
        var main = WordModel.Doc(package).MainDocumentPart
            ?? throw new InvalidOperationException("The document has no main part.");
        var part = main.StyleDefinitionsPart ?? main.AddNewPart<StyleDefinitionsPart>();
        part.Styles ??= new Styles();

        if (style.GetFirstChild<BasedOn>()?.Val?.Value is { } parent
            && !part.Styles.Elements<Style>().Any(s => string.Equals(s.StyleId?.Value, parent, StringComparison.Ordinal))
            && BuiltIn(parent) is { } parentStyle)
            AddBuiltIn(package, parentStyle);

        part.Styles.AppendChild(style);
    }

    private static string Known(IReadOnlyList<Entry> catalog, string type)
    {
        var names = catalog.Where(e => e.Type == type).Select(e => $"'{e.Name}'").Distinct().ToList();
        if (names.Count == 0) return $"The document defines no {type} styles.";

        const int shown = 40;
        var list = string.Join(", ", names.Take(shown));
        if (names.Count > shown) list += $" and {names.Count - shown} more";
        return $"Its {type} styles are: {list}; inspect_document lists them under styles.";
    }

    private static string TypeOf(Style style)
    {
        var type = style.Type?.Value;
        if (type == StyleValues.Character) return "character";
        if (type == StyleValues.Table) return "table";
        if (type == StyleValues.Numbering) return "numbering";
        // A style with no type is a paragraph style, as Word reads it.
        return "paragraph";
    }
}
