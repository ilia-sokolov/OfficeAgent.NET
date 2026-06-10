using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.Word;

/// <summary>
/// Shared Word helpers used across the format module and its handlers: text-host
/// enumeration (body, headers, footers, footnotes, endnotes), paragraph addressing
/// (stable ids), the run-spanning text engine, and the snapshot token. Inspect,
/// stabilize, and apply use the same id algorithm so anchors round-trip.
/// </summary>
internal static class WordModel
{
    public static readonly WordmlDialect Dialect = new();
    public static readonly TextBodyEngine Text = new(Dialect);

    public static WordprocessingDocument Doc(IOpenXmlPackage package) =>
        (WordprocessingDocument)package.Package;

    public static MainDocumentPart Main(IOpenXmlPackage package) =>
        Doc(package).MainDocumentPart
        ?? throw new InvalidOperationException("Word document has no main part.");

    public static Body Body(IOpenXmlPackage package) =>
        Main(package).Document?.Body
        ?? throw new InvalidOperationException("Word document has no body.");

    /// <summary>
    /// Enumerates the editable text hosts of the document with a stable per-host key:
    /// the body (""), each header (<c>hdr0</c>…), each footer (<c>ftr0</c>…), the
    /// footnotes (<c>fn</c>) and endnotes (<c>en</c>) parts.
    /// </summary>
    public static IEnumerable<(OpenXmlElement Root, string HostKey)> TextHosts(IOpenXmlPackage package)
    {
        var main = Main(package);

        if (main.Document?.Body is { } body)
            yield return (body, "");

        int h = 0;
        foreach (var header in main.HeaderParts)
            if (header.Header is { } el)
                yield return (el, $"hdr{h++}");

        int f = 0;
        foreach (var footer in main.FooterParts)
            if (footer.Footer is { } el)
                yield return (el, $"ftr{f++}");

        if (main.FootnotesPart?.Footnotes is { } footnotes)
            yield return (footnotes, "fn");

        if (main.EndnotesPart?.Endnotes is { } endnotes)
            yield return (endnotes, "en");
    }

    public static IEnumerable<(Paragraph Paragraph, string ParaId, string HostKey)> Paragraphs(IOpenXmlPackage package)
    {
        foreach (var (root, hostKey) in TextHosts(package))
        {
            int index = 0;
            foreach (var paragraph in root.Descendants<Paragraph>())
            {
                yield return (paragraph, ComputeParaId(paragraph, hostKey, index), hostKey);
                index++;
            }
        }
    }

    public static string ComputeParaId(Paragraph paragraph, string hostKey, int index)
    {
        var w14 = paragraph.ParagraphId?.Value;
        if (!string.IsNullOrEmpty(w14)) return $"w14:{w14}";
        return hostKey.Length == 0 ? $"auto-{index:D4}" : $"auto-{hostKey}-{index:D4}";
    }

    public static Paragraph? ResolveParagraph(IOpenXmlPackage package, string paraId)
    {
        foreach (var (paragraph, id, _) in Paragraphs(package))
            if (string.Equals(id, paraId, StringComparison.Ordinal))
                return paragraph;
        return null;
    }

    /// <summary>Resolves a plan anchor's paragraph, applying the stabilization alias map first.</summary>
    public static Paragraph? ResolveParagraph(ApplyContext context, string anchorParaId) =>
        ResolveParagraph(context.Package, context.ResolveAlias(anchorParaId));

    /// <summary>
    /// Assigns <c>w14:paraId</c> to every paragraph (across all hosts) that lacks one and
    /// returns the positional-to-stable id alias map. Idempotent.
    /// </summary>
    public static IReadOnlyDictionary<string, string> Stabilize(IOpenXmlPackage package)
    {
        var aliases = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var (root, hostKey) in TextHosts(package))
        {
            int index = 0;
            foreach (var paragraph in root.Descendants<Paragraph>())
            {
                var positional = hostKey.Length == 0 ? $"auto-{index:D4}" : $"auto-{hostKey}-{index:D4}";
                if (string.IsNullOrEmpty(paragraph.ParagraphId?.Value))
                    paragraph.ParagraphId = NewParaId();
                aliases[positional] = $"w14:{paragraph.ParagraphId!.Value}";
                index++;
            }
        }
        return aliases;
    }

    private static string NewParaId()
    {
        var bytes = new byte[4];
        using (var rng = RandomNumberGenerator.Create())
            rng.GetBytes(bytes);
        return ToHex(bytes);
    }

    private static string ToHex(byte[] bytes) =>
        BitConverter.ToString(bytes).Replace("-", "");

    public static string? StyleOf(Paragraph paragraph) =>
        paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;

    public static StringComparison Comparison(bool caseSensitive) =>
        caseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

    public static string Snippet(string text, int start, int length, int pad = 32)
    {
        int from = Math.Max(0, start - pad);
        int to = Math.Min(text.Length, start + length + pad);
        var slice = text.Substring(from, to - from);
        return (from > 0 ? "…" : "") + slice + (to < text.Length ? "…" : "");
    }

    /// <summary>
    /// Snapshot etag over every text host so drift in a header or footer is detected,
    /// not just drift in the body.
    /// </summary>
    public static SnapshotToken Snapshot(IOpenXmlPackage package)
    {
        var sb = new StringBuilder();
        foreach (var (root, hostKey) in TextHosts(package))
        {
            sb.Append(hostKey).Append('\n');
            sb.Append(root.OuterXml).Append('\n');
        }
        byte[] hash;
        using (var sha = SHA256.Create())
            hash = sha.ComputeHash(Encoding.UTF8.GetBytes(sb.ToString()));
        return new SnapshotToken(ToHex(hash));
    }
}

/// <summary>
/// Document-scoped, collision-aware revision id allocator. Seeds from the highest
/// id already present across every text host so revisions written by this engine
/// never clash with revisions Word or other tools wrote earlier.
/// </summary>
internal sealed class WordRevisionIdAllocator
{
    private int _next;

    public WordRevisionIdAllocator(IOpenXmlPackage package)
    {
        int max = 0;
        foreach (var (root, _) in WordModel.TextHosts(package))
        {
            foreach (var ins in root.Descendants<InsertedRun>())
                if (int.TryParse(ins.Id?.Value, out var id) && id > max) max = id;
            foreach (var del in root.Descendants<DeletedRun>())
                if (int.TryParse(del.Id?.Value, out var id) && id > max) max = id;
        }
        _next = max + 1;
    }

    public int Next() => _next++;
}
