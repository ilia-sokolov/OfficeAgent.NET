using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.Word;

/// <summary>
/// Real list numbering: the <c>numbering.xml</c> part, the definitions in it, and the
/// <c>w:numPr</c> that points a paragraph at one.
/// </summary>
/// <remarks>
/// <para>
/// The point of this over writing "1." at the front of the text is that Word owns the
/// numbers. Insert a clause in the middle of a contract and everything below it renumbers;
/// type a literal "4.2" and it does not, which is how a contract ends up with two clause
/// 7s. The same goes for a manual's steps.
/// </para>
/// <para>
/// Both <c>w:numbering</c> and <c>w:lvl</c> are strict sequences, and Word offers to repair
/// a file that gets either wrong, so every element here is built in schema order rather than
/// appended.
/// </para>
/// </remarks>
internal static class WordNumbering
{
    /// <summary>The list looks a plan may ask for.</summary>
    public const string Names = "bullet, decimal, clause, none";

    /// <summary>How far each level is indented, in twips. A quarter inch per level.</summary>
    private const int IndentPerLevel = 360;

    /// <summary>The deepest level WordprocessingML allows, zero-based.</summary>
    public const int MaxLevel = 8;

    public static bool IsStyle(string? value) =>
        value is not null && Vocabulary.Contains(value.Trim().ToLowerInvariant());

    public static bool IsNone(string? value) =>
        string.Equals(value?.Trim(), "none", StringComparison.OrdinalIgnoreCase);

    private static readonly HashSet<string> Vocabulary =
        new(StringComparer.OrdinalIgnoreCase) { "bullet", "decimal", "clause", "none" };

    /// <summary>
    /// Points a paragraph at a numbering definition, creating the definition and the part
    /// when they do not exist yet. Passing <c>none</c> takes the numbering off again.
    /// </summary>
    public static void Apply(
        MainDocumentPart main, Paragraph paragraph, string style, int level, int listId)
    {
        var properties = paragraph.ParagraphProperties ??= new ParagraphProperties();

        if (IsNone(style))
        {
            properties.GetFirstChild<NumberingProperties>()?.Remove();
            return;
        }

        var numberId = NumberIdFor(main, style, listId);
        properties.GetFirstChild<NumberingProperties>()?.Remove();

        var numbering = new NumberingProperties(
            new NumberingLevelReference { Val = level < 0 ? 0 : (level > MaxLevel ? MaxLevel : level) },
            new NumberingId { Val = numberId });

        // Placed through the one shared order rather than by hand. w:numPr sits after
        // w:pStyle *and* after w:pageBreakBefore, which is easy to get backwards and shows
        // up only as a repair prompt on the one paragraph that has both.
        FormatHandler.PlaceParagraphProperty(properties, numbering);
    }

    /// <summary>
    /// The <c>w:numId</c> for a look, creating the abstract definition and the instance on
    /// first use and reusing them afterwards.
    /// </summary>
    /// <remarks>
    /// One instance per (look, <paramref name="listId"/>) pair. Paragraphs sharing both share
    /// one running sequence - which is what a contract's clauses want. A manual whose second
    /// chapter restarts at 1 asks for a different <paramref name="listId"/>, and gets its own
    /// instance pointing at the same shared definition.
    /// </remarks>
    public static int NumberIdFor(MainDocumentPart main, string style, int listId)
    {
        var part = main.NumberingDefinitionsPart;
        if (part is null)
        {
            part = main.AddNewPart<NumberingDefinitionsPart>();
            part.Numbering = new Numbering();
        }

        var numbering = part.Numbering ??= new Numbering();
        var name = $"officeagent-{style.ToLowerInvariant()}";
        var instanceName = $"{name}-{listId}";

        // The instance is recognised by the marker written into its abstract definition's
        // w:name, so a document edited over several plans reuses what it already has.
        foreach (var existing in numbering.Elements<NumberingInstance>())
        {
            var abstractId = existing.AbstractNumId?.Val?.Value;
            if (abstractId is null) continue;

            var definition = numbering.Elements<AbstractNum>()
                .FirstOrDefault(a => a.AbstractNumberId?.Value == abstractId);

            if (definition?.GetFirstChild<AbstractNumDefinitionName>()?.Val?.Value == instanceName)
                return existing.NumberID!.Value;
        }

        var newAbstractId = NextAbstractId(numbering);
        var newNumberId = NextNumberId(numbering);

        InsertAbstract(numbering, BuildAbstract(style, newAbstractId, instanceName));
        InsertInstance(numbering, new NumberingInstance(
            new AbstractNumId { Val = newAbstractId })
        {
            NumberID = newNumberId
        });

        return newNumberId;
    }

    /// <summary>
    /// <c>w:abstractNum</c> elements precede every <c>w:num</c> in <c>w:numbering</c>.
    /// </summary>
    private static void InsertAbstract(Numbering numbering, AbstractNum definition)
    {
        var firstInstance = numbering.Elements<NumberingInstance>().FirstOrDefault();
        if (firstInstance is null) numbering.AppendChild(definition);
        else numbering.InsertBefore(definition, firstInstance);
    }

    private static void InsertInstance(Numbering numbering, NumberingInstance instance) =>
        numbering.AppendChild(instance);

    private static int NextAbstractId(Numbering numbering)
    {
        var highest = -1;
        foreach (var definition in numbering.Elements<AbstractNum>())
            if (definition.AbstractNumberId?.Value is { } id && id > highest) highest = id;
        return highest + 1;
    }

    private static int NextNumberId(Numbering numbering)
    {
        var highest = 0;
        foreach (var instance in numbering.Elements<NumberingInstance>())
            if (instance.NumberID?.Value is { } id && id > highest) highest = id;
        return highest + 1;
    }

    /// <summary>
    /// Builds the nine levels of a definition. The looks differ only in what each level
    /// prints and where it restarts.
    /// </summary>
    private static AbstractNum BuildAbstract(string style, int abstractId, string name)
    {
        var definition = new AbstractNum { AbstractNumberId = abstractId };

        // Children of w:abstractNum are a sequence: nsid, multiLevelType, tmpl, name, then
        // the levels.
        definition.Append(new Nsid { Val = Hash(name) });
        definition.Append(new MultiLevelType
        {
            Val = style.Equals("clause", StringComparison.OrdinalIgnoreCase)
                ? MultiLevelValues.Multilevel
                : MultiLevelValues.HybridMultilevel
        });
        definition.Append(new AbstractNumDefinitionName { Val = name });

        for (var level = 0; level <= MaxLevel; level++)
            definition.Append(BuildLevel(style, level));

        return definition;
    }

    private static Level BuildLevel(string style, int level)
    {
        var (format, text) = Shape(style, level);

        // Children of w:lvl are a sequence: start, numFmt, lvlText, lvlJc, pPr, rPr.
        var definition = new Level { LevelIndex = level };
        definition.Append(new StartNumberingValue { Val = 1 });
        definition.Append(new NumberingFormat { Val = format });
        definition.Append(new LevelText { Val = text });
        definition.Append(new LevelJustification { Val = LevelJustificationValues.Left });

        var indent = IndentPerLevel * (level + 1);
        definition.Append(new PreviousParagraphProperties(
            new Indentation
            {
                Left = indent.ToString(),
                Hanging = IndentPerLevel.ToString()
            }));

        // A bullet is a glyph from a symbol face, not a letter - without the font the
        // reader gets whatever their default face has at that code point.
        if (format == NumberFormatValues.Bullet)
            definition.Append(new NumberingSymbolRunProperties(
                new RunFonts { Ascii = "Symbol", HighAnsi = "Symbol", Hint = FontTypeHintValues.Default }));

        return definition;
    }

    /// <summary>
    /// What each level prints. <c>%1</c> is the counter for level 0, <c>%2</c> for level 1,
    /// and so on - which is how <c>1.2.3</c> is built out of three of them.
    /// </summary>
    private static (NumberFormatValues Format, string Text) Shape(string style, int level) =>
        style.ToLowerInvariant() switch
        {
            // A contract: every level carries its parents, so a clause is addressable by
            // number in the text of another one.
            "clause" => (NumberFormatValues.Decimal, Numbered(level)),

            // A list: the usual 1. / a. / i. rotation, each level standing alone.
            "decimal" => ((level % 3) switch
            {
                0 => NumberFormatValues.Decimal,
                1 => NumberFormatValues.LowerLetter,
                _ => NumberFormatValues.LowerRoman
            }, $"%{level + 1}."),

            _ => (NumberFormatValues.Bullet, (level % 3) switch
            {
                0 => "",   // •
                1 => "o",
                _ => ""    // ▪
            })
        };

    private static string Numbered(int level) =>
        string.Join(".", Enumerable.Range(1, level + 1).Select(i => $"%{i}")) +
        (level == 0 ? "." : string.Empty);

    /// <summary>
    /// A stable eight-hex-digit id derived from the name, so regenerating a document twice
    /// produces the same bytes rather than a fresh random nsid each run.
    /// </summary>
    private static string Hash(string name)
    {
        unchecked
        {
            var hash = 2166136261u;
            foreach (var c in name)
            {
                hash ^= c;
                hash *= 16777619u;
            }
            // Word treats an nsid of zero as unset.
            return (hash == 0 ? 1u : hash).ToString("X8");
        }
    }
}
