using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using OfficeAgent.Abstractions;
using A = DocumentFormat.OpenXml.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.Core;

/// <summary>
/// A paragraph's text with its character formatting written as numbered tags, so a model can
/// rewrite the whole paragraph - translate it, typically - and keep the bold, italic, colour
/// and link formatting on the words that carried it.
/// </summary>
/// <remarks>
/// <para>
/// Runs with identical properties form one group. The group covering the most characters is
/// the paragraph's base formatting and is written untagged; every other group gets a number,
/// in order of first appearance: <c>Payment is due within &lt;1&gt;30 days&lt;/1&gt;.</c>.
/// A rewrite keeps the tags around the words that should carry that formatting, wherever the
/// translation puts them, and may use a tag more than once or not at all.
/// </para>
/// <para>
/// Tags are flat: <c>&lt;n&gt;</c> opens a group and <c>&lt;/n&gt;</c> closes it, and only a
/// number between angle brackets is a tag, so any other <c>&lt;</c> is text. A paragraph whose
/// text crosses a hyperlink, content control or tracked change in Word has no marked text,
/// because a rewrite cannot move text between those containers.
/// </para>
/// </remarks>
internal static class MarkedText
{
    private static readonly Regex Tag = new(@"<(/?)(\d{1,3})>", RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));

    /// <summary>
    /// Run properties PowerPoint rewrites as the user types or the spell checker runs. They
    /// change nothing a reader sees, so they must not split one formatting group into two.
    /// </summary>
    private static readonly string[] VolatileDrawingAttributes = { "dirty", "err", "smtClean", "smtId", "bmk" };

    /// <summary>One formatting group: every run with these properties, and the first of them as its pattern.</summary>
    internal sealed class Group
    {
        internal Group(string key, OpenXmlElement pattern)
        {
            Key = key;
            Pattern = pattern;
        }

        internal string Key { get; }

        internal OpenXmlElement Pattern { get; }

        internal int Characters { get; set; }

        /// <summary>The tag number, or 0 for the base group, which is written untagged.</summary>
        internal int Number { get; set; }
    }

    /// <summary>A paragraph's formatting groups, or why it has none a rewrite can use.</summary>
    internal sealed class Layout
    {
        internal Layout(IReadOnlyList<(OpenXmlElement Run, string Text, Group Group)> runs, IReadOnlyList<Group> groups, string? refusal)
        {
            Runs = runs;
            Groups = groups;
            Refusal = refusal;
        }

        internal IReadOnlyList<(OpenXmlElement Run, string Text, Group Group)> Runs { get; }

        internal IReadOnlyList<Group> Groups { get; }

        /// <summary>Why the paragraph cannot take a marked rewrite, or <see langword="null"/> when it can.</summary>
        internal string? Refusal { get; }

        internal Group? Base => Groups.FirstOrDefault(group => group.Number == 0);
    }

    /// <summary>Reads a paragraph's runs into formatting groups.</summary>
    internal static Layout Read(OpenXmlElement paragraph, ITextDialect dialect)
    {
        var runs = new List<(OpenXmlElement, string, Group)>();
        var groups = new List<Group>();
        var byKey = new Dictionary<string, Group>(StringComparer.Ordinal);
        OpenXmlElement? container = null;
        var crossesContainers = false;

        foreach (var run in dialect.GetRuns(paragraph))
        {
            if (!dialect.IsTextRun(run)) continue;
            var text = dialect.GetRunText(run);
            if (text.Length == 0) continue;

            container ??= run.Parent;
            if (!ReferenceEquals(container, run.Parent)) crossesContainers = true;

            var key = FormattingKey(run);
            if (!byKey.TryGetValue(key, out var group))
            {
                group = new Group(key, run);
                byKey[key] = group;
                groups.Add(group);
            }

            group.Characters += text.Length;
            runs.Add((run, text, group));
        }

        // The base is the formatting most of the text has; on a tie, the one that comes first.
        var baseGroup = groups.OrderByDescending(group => group.Characters).FirstOrDefault();
        var number = 1;
        foreach (var group in groups)
            group.Number = ReferenceEquals(group, baseGroup) ? 0 : number++;

        var refusal = crossesContainers
            ? "its text crosses a hyperlink, content control or tracked change, and a rewrite cannot move " +
              "text between those. Rewrite each part with its own changeText and \"with\", or settle pending " +
              "revisions first"
            : null;
        return new Layout(runs, groups, refusal);
    }

    /// <summary>
    /// The paragraph's marked text, or <see langword="null"/> when tags would add nothing - one
    /// formatting throughout - or the paragraph cannot take a marked rewrite.
    /// </summary>
    internal static string? Describe(OpenXmlElement paragraph, ITextDialect dialect)
    {
        var layout = Read(paragraph, dialect);
        if (layout.Groups.Count < 2 || layout.Refusal is not null) return null;

        var builder = new StringBuilder();
        Group? open = null;
        foreach (var (_, text, group) in layout.Runs)
        {
            if (!ReferenceEquals(open, group))
            {
                if (open is { Number: > 0 }) builder.Append("</").Append(open.Number).Append('>');
                if (group.Number > 0) builder.Append('<').Append(group.Number).Append('>');
                open = group;
            }

            builder.Append(text);
        }

        if (open is { Number: > 0 }) builder.Append("</").Append(open.Number).Append('>');
        return builder.ToString();
    }

    /// <summary>One piece of a parsed rewrite: its text and the group number it takes, 0 for the base.</summary>
    internal readonly struct Piece
    {
        internal Piece(int number, string text)
        {
            Number = number;
            Text = text;
        }

        internal int Number { get; }

        internal string Text { get; }
    }

    /// <summary>
    /// Splits marked text into pieces, or returns why it cannot. Tags are checked against
    /// <paramref name="highestTag"/>, the largest number the paragraph's marked text uses.
    /// </summary>
    internal static IReadOnlyList<Piece> Parse(string marked, int highestTag, out string? error)
    {
        var pieces = new List<Piece>();
        error = null;
        var open = 0;
        var cursor = 0;

        foreach (Match match in Tag.Matches(marked))
        {
            if (match.Index > cursor)
                pieces.Add(new Piece(open, marked.Substring(cursor, match.Index - cursor)));
            cursor = match.Index + match.Length;

            var closing = match.Groups[1].Length > 0;
            var number = int.Parse(match.Groups[2].Value, System.Globalization.CultureInfo.InvariantCulture);

            if (number < 1 || number > highestTag)
            {
                error = highestTag == 0
                    ? $"withMarked uses tag <{number}>, but this paragraph has one formatting throughout, so it takes no tags. Send the text without them."
                    : $"withMarked uses tag <{number}>, which this paragraph does not have. Its tags run from 1 to {highestTag}; copy them from markedText in inspect_document.";
                return Array.Empty<Piece>();
            }

            if (!closing)
            {
                if (open != 0)
                {
                    error = $"withMarked opens <{number}> inside <{open}>. Tags do not nest: close <{open}> first.";
                    return Array.Empty<Piece>();
                }

                open = number;
            }
            else
            {
                if (open != number)
                {
                    error = open == 0
                        ? $"withMarked closes </{number}> without opening it."
                        : $"withMarked closes </{number}> while <{open}> is open.";
                    return Array.Empty<Piece>();
                }

                open = 0;
            }
        }

        if (open != 0)
        {
            error = $"withMarked opens <{open}> and never closes it.";
            return Array.Empty<Piece>();
        }

        if (cursor < marked.Length)
            pieces.Add(new Piece(0, marked.Substring(cursor)));

        return pieces.Where(piece => piece.Text.Length > 0).ToList();
    }

    /// <summary>
    /// Checks a <see cref="ChangeTextOp.WithMarked"/> rewrite against the paragraph it targets
    /// and parses it, or returns the error that refuses it. Shared by every format that takes one.
    /// </summary>
    internal static ValidationError? Check(
        ChangeTextOp op,
        TextSpanAnchor anchor,
        string paragraphText,
        Layout layout,
        out IReadOnlyList<Piece> pieces)
    {
        pieces = Array.Empty<Piece>();
        var marked = op.WithMarked!;

        if (op.With.Length > 0)
            return Refuse("changeText takes either \"with\" or \"withMarked\", not both. Keep withMarked to preserve the paragraph's formatting.");

        if (!string.Equals(anchor.Expect, paragraphText, StringComparison.Ordinal) || anchor.Occurrence != 0)
            return Refuse("withMarked rewrites a whole paragraph, so \"expect\" must be the paragraph's complete text " +
                          "with occurrence 0. To change part of a paragraph, use \"with\".");

        if (layout.Refusal is not null)
            return Refuse($"withMarked cannot rewrite paragraph '{anchor.ParaId}': {layout.Refusal}.");

        pieces = Parse(marked, Math.Max(0, layout.Groups.Count - 1), out var error);
        return error is null ? null : Refuse(error);

        ValidationError Refuse(string message) =>
            new(ValidationErrorCodes.InvalidOperation, message, anchor);
    }

    /// <summary>The text a marked rewrite produces, with its tags removed.</summary>
    internal static string Plain(IReadOnlyList<Piece> pieces) => string.Concat(pieces.Select(piece => piece.Text));

    /// <summary>
    /// Builds one new run per piece, each carrying the properties of its group's pattern run
    /// and nothing else of it: a pattern run that also held a tab, a break or a picture lends
    /// its formatting, not its content.
    /// </summary>
    internal static IReadOnlyList<OpenXmlElement> BuildRuns(Layout layout, IReadOnlyList<Piece> pieces)
    {
        var byNumber = layout.Groups.ToDictionary(group => group.Number);
        return pieces.Select(piece => NewRun(byNumber[piece.Number].Pattern, piece.Text)).ToList();
    }

    private static OpenXmlElement NewRun(OpenXmlElement pattern, string text) => pattern switch
    {
        W.Run run => new W.Run(
            run.RunProperties is { } properties
                ? new OpenXmlElement[] { properties.CloneNode(deep: true), new W.Text(text) { Space = SpaceProcessingModeValues.Preserve } }
                : new OpenXmlElement[] { new W.Text(text) { Space = SpaceProcessingModeValues.Preserve } }),
        A.Run run => new A.Run(
            run.RunProperties is { } properties
                ? new OpenXmlElement[] { properties.CloneNode(deep: true), new A.Text(text) }
                : new OpenXmlElement[] { new A.Text(text) }),
        _ => throw new InvalidOperationException($"No marked-text support for run type {pattern.GetType().Name}.")
    };

    private static string FormattingKey(OpenXmlElement run)
    {
        switch (run)
        {
            case W.Run word:
                return word.RunProperties?.OuterXml ?? string.Empty;
            case A.Run drawing when drawing.RunProperties is { } properties:
                var copy = (A.RunProperties)properties.CloneNode(deep: true);
                foreach (var name in VolatileDrawingAttributes)
                    copy.RemoveAttribute(name, string.Empty);
                return copy.OuterXml;
            default:
                return string.Empty;
        }
    }
}
