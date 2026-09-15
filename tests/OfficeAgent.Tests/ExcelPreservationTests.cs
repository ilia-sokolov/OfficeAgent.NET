using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Excel;

namespace OfficeAgent.Tests;

/// <summary>
/// What survives an Excel cell edit, byte for byte and semantically.
/// </summary>
/// <remarks>
/// V09-13's benchmark preservation check surfaced this. Word's modules leave parts they do
/// not touch byte-identical, which V09-03B established and asserts. Excel does not, for parts
/// it reads: writing one cell well outside a table's range re-serialises the tables part.
/// <para>
/// The difference is presentational, not semantic. It matters because "unchanged" is a claim
/// this project makes carefully, and it is worth the distinction being pinned by a test rather
/// than discovered again later. If the table part ever changes in meaning, this fails.
/// </para>
/// </remarks>
public sealed class ExcelPreservationTests
{
    private const uint DataSheetId = 7U;

    [Fact]
    public void A_cell_edit_outside_a_table_leaves_styles_byte_identical()
    {
        var (before, after) = EditOutsideTheTable();

        Assert.Equal(Part(before, "xl/styles.xml"), Part(after, "xl/styles.xml"));
    }

    [Fact]
    public void A_cell_edit_outside_a_table_leaves_the_table_semantically_unchanged()
    {
        var (before, after) = EditOutsideTheTable();

        var one = Part(before, "xl/tables/table1.xml");
        var two = Part(after, "xl/tables/table1.xml");
        Assert.NotNull(one);
        Assert.NotNull(two);

        // Not byte-identical: the SDK rewrites the part it loaded, moving the namespace
        // declaration to the front of the attribute list. Same length, same content.
        Assert.False(one!.SequenceEqual(two!), "the table part is now byte-identical; " +
            "this test's premise has changed and its remarks need revisiting");

        // What actually has to hold: the table still describes the same table. Compared
        // canonically, because attribute order is exactly what the SDK changed and
        // XNode.DeepEquals is sensitive to it.
        Assert.Equal(
            Canonical(XDocument.Parse(Encoding.UTF8.GetString(one!)).Root!),
            Canonical(XDocument.Parse(Encoding.UTF8.GetString(two!)).Root!));
    }

    /// <summary>An element as name, sorted attributes and children, so ordering cannot differ.</summary>
    private static string Canonical(XElement element)
    {
        var attributes = element.Attributes()
            .Where(attribute => !attribute.IsNamespaceDeclaration)
            .OrderBy(attribute => attribute.Name.ToString(), StringComparer.Ordinal)
            .Select(attribute => $"{attribute.Name}={attribute.Value}");

        var children = element.Elements().Select(Canonical);
        var text = element.Nodes().OfType<XText>().Select(node => node.Value.Trim())
            .Where(value => value.Length > 0);

        return $"{element.Name}[{string.Join(",", attributes)}]" +
               $"({string.Join("", children)}){string.Join("", text)}";
    }

    private static (byte[] Before, byte[] After) EditOutsideTheTable()
    {
        var root = Path.Combine(
            Path.GetDirectoryName(typeof(ExcelPreservationTests).Assembly.Location)!,
            "..", "..", "..", "Corpus", "v0.8.0");
        var before = File.ReadAllBytes(Path.GetFullPath(Path.Combine(root, "workbook-styles.xlsx")));

        // The table occupies A1:B2; H1 is nowhere near it.
        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new SetCellOp
                {
                    Target = new CellAnchor { SheetId = DataSheetId, Address = "H1" },
                    Value = "11"
                }
            }
        };

        var client = new OfficeAgentClient(new ExcelModule());
        using var applied = client.Apply(
            new StreamHandle(new MemoryStream(before, writable: false)), plan, ApplyOptions.Commit);
        Assert.True(applied.Committed, string.Join("; ", applied.Report.Errors.Select(e => e.Code)));
        return (before, applied.ToBytes());
    }

    private static byte[]? Part(byte[] package, string name)
    {
        using var archive = new ZipArchive(new MemoryStream(package, writable: false), ZipArchiveMode.Read);
        var entry = archive.GetEntry(name);
        if (entry is null) return null;
        using var stream = entry.Open();
        using var copy = new MemoryStream();
        stream.CopyTo(copy);
        return copy.ToArray();
    }
}
