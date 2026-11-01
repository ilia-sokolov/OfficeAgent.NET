using System.IO.Compression;
using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using OfficeAgent.Abstractions;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Excel;
using OfficeAgent.PowerPoint;
using OfficeAgent.Word;
using DocFormat = OfficeAgent.Abstractions.DocumentFormat;
using Vt = DocumentFormat.OpenXml.VariantTypes;

namespace OfficeAgent.Tests;

/// <summary>
/// Sensitivity labels and encrypted files. A label is reported and kept; an encrypted file is
/// refused by name. The hosts decide what labels mean, so the engine's job is to say what is
/// there and to leave it alone.
/// </summary>
public class ProtectedContentTests
{
    private const string LabelId = "11111111-2222-3333-4444-555555555555";

    private static readonly string CorpusRoot = Path.Combine(
        AppContext.BaseDirectory, "..", "..", "..", "Corpus", "encrypted");

    private static byte[] Encrypted(string name) => File.ReadAllBytes(Path.Combine(CorpusRoot, name));

    // ── encrypted files ───────────────────────────────────────────────────

    [Theory]
    [InlineData("encrypted.xlsx")]
    [InlineData("encrypted.pptx")]
    public void A_file_Office_encrypted_is_refused_by_name(string fixture)
    {
        var client = new OfficeAgentClient(new ExcelModule(), new PowerPointModule(), new WordModule());

        var error = Assert.Throws<EncryptedDocumentException>(() =>
            client.Inspect(new StreamHandle(new MemoryStream(Encrypted(fixture)), fixture)));

        Assert.Contains("sensitivity label or a password", error.Message);
    }

    [Fact]
    public void An_encrypted_document_is_refused_whichever_format_its_name_claims()
    {
        // The encrypted container is the same for every format, and the check runs before the
        // format is known.
        var client = new OfficeAgentClient(new WordModule());

        Assert.Throws<EncryptedDocumentException>(() =>
            client.Inspect(new StreamHandle(new MemoryStream(Encrypted("encrypted.xlsx")), "contract.docx")));
    }

    [Fact]
    public async Task A_tool_reports_the_stable_code_and_keeps_the_inline_path_honest()
    {
        using var workspace = new CompositeToolsTests.CompositeWorkspace();
        var tools = new OfficeAgentTools(workspace.Client);
        workspace.Stage("protected.docx", Encrypted("encrypted.xlsx"));

        const string plan = """[ { "op": "changeText", "target": { "find": "x" }, "with": "y" } ]""";
        var edited = await tools.EditDocument("workspace", "protected.docx", plan);
        Assert.Contains(EncryptedDocumentException.Code, edited);

        // The inline tools call unreadable bytes a bad copy; re-sending this one cannot help.
        var inline = await tools.EditDocumentContent(Convert.ToBase64String(Encrypted("encrypted.xlsx")), plan);
        Assert.Contains(EncryptedDocumentException.Code, inline);
    }

    [Fact]
    public void Comparing_an_encrypted_document_names_the_protection()
    {
        // Comparison reads the package itself, so it must not call this an oversized input.
        var client = new OfficeAgentClient(new WordModule());

        Assert.Throws<EncryptedDocumentException>(() =>
            client.CompareDocuments(Encrypted("encrypted.xlsx"), DocxFactory.Contract()));
        Assert.Throws<EncryptedDocumentException>(() =>
            client.CompareDocuments(DocxFactory.Contract(), Encrypted("encrypted.xlsx")));
    }

    [Fact]
    public async Task Merging_an_encrypted_document_names_the_protection()
    {
        var store = new MemoryDocumentProvider("docs");
        var left = store.Add("a.docx", DocxFactory.Contract());
        var right = store.Add("b.docx", Encrypted("encrypted.xlsx"));
        var client = new OfficeAgentClient(new DocumentProviderRegistry(new[] { store }), new WordModule());

        await Assert.ThrowsAsync<EncryptedDocumentException>(() =>
            client.PreviewMergeAsync(new() { Sources = new[] { right, left } }));
    }

    [Fact]
    public void A_table_past_the_header_list_is_followed_to_the_stream()
    {
        // A file above about 7 MB keeps part of its allocation table in a further sector chain.
        var client = new OfficeAgentClient(new WordModule());

        Assert.Throws<EncryptedDocumentException>(() =>
            client.Inspect(new StreamHandle(new MemoryStream(CompoundFile(encrypted: true)), "big.docx")));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    public void A_compound_file_without_the_stream_is_only_unreadable(bool encrypted, bool cyclic)
    {
        // A legacy .doc is a compound file too, and a chain that loops must end.
        var client = new OfficeAgentClient(new WordModule());

        Assert.Throws<OpenXmlPackageRejectedException>(() =>
            client.Inspect(new StreamHandle(new MemoryStream(CompoundFile(encrypted, cyclic)), "legacy.docx")));
    }

    [Fact]
    public void A_truncated_compound_file_is_only_unreadable()
    {
        var client = new OfficeAgentClient(new WordModule());
        var header = CompoundFile(encrypted: true).Take(600).ToArray();

        Assert.Throws<OpenXmlPackageRejectedException>(() =>
            client.Inspect(new StreamHandle(new MemoryStream(header), "cut.docx")));
    }

    /// <summary>
    /// A compound file whose directory runs through a second sector, recorded in an allocation
    /// sector reachable only through the extended list that large files use.
    /// </summary>
    private static byte[] CompoundFile(bool encrypted, bool cyclic = false)
    {
        const int sectorSize = 512;
        const int perFatSector = sectorSize / 4;
        const uint free = 0xFFFFFFFF, end = 0xFFFFFFFE;
        const int fatSectors = 110;
        const int extended = fatSectors;
        const int firstDirectory = 109 * perFatSector;
        const int total = firstDirectory + 2;

        var file = new byte[sectorSize * (total + 1)];
        void U32(int offset, uint value) => BitConverter.GetBytes(value).CopyTo(file, offset);
        int Sector(int n) => (n + 1) * sectorSize;

        new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 }.CopyTo(file, 0);
        BitConverter.GetBytes((ushort)9).CopyTo(file, 30);
        U32(44, fatSectors);
        U32(48, firstDirectory);
        U32(68, extended);
        U32(72, 1);
        for (var i = 0; i < 109; i++) U32(76 + (i * 4), (uint)i);

        // The extended list names the 110th allocation sector and ends.
        U32(Sector(extended), 109);
        for (var i = 1; i < (sectorSize / 4) - 1; i++) U32(Sector(extended) + (i * 4), free);
        U32(Sector(extended) + sectorSize - 4, end);

        U32(Sector(109), firstDirectory + 1);
        U32(Sector(109) + 4, cyclic ? (uint)firstDirectory : end);

        if (encrypted)
        {
            var entry = Sector(firstDirectory + 1);
            Encoding.Unicode.GetBytes("EncryptedPackage").CopyTo(file, entry);
            BitConverter.GetBytes((ushort)34).CopyTo(file, entry + 64);
            file[entry + 66] = 2;
        }

        return file;
    }

    // ── labels ────────────────────────────────────────────────────────────

    [Fact]
    public void A_label_in_the_custom_properties_is_reported_in_every_format()
    {
        foreach (var (client, bytes) in Formats())
        {
            var labelled = WithProperties(bytes, enabled: "true");

            var node = Assert.Single(client.Inspect(new StreamHandle(new MemoryStream(labelled)))
                .Nodes, n => n.Kind == "sensitivityLabel");

            Assert.Equal("Confidential", node.Name);
            Assert.Equal($"sensitivityLabel#{LabelId}", node.Path);
            Assert.Contains(LabelId, node.Summary);
            Assert.Contains("applied by a person", node.Summary);
        }
    }

    [Fact]
    public void A_label_in_LabelInfo_is_reported_without_a_name()
    {
        foreach (var (client, bytes) in Formats())
        {
            var labelled = WithLabelInfo(bytes, "enabled=\"1\" method=\"Standard\" removed=\"0\"");

            var node = Assert.Single(client.Inspect(new StreamHandle(new MemoryStream(labelled)))
                .Nodes, n => n.Kind == "sensitivityLabel");

            Assert.Contains("unnamed label", node.Summary);
            Assert.Contains("applied by default", node.Summary);
        }
    }

    [Fact]
    public void A_label_recorded_in_both_places_is_one_label()
    {
        var (client, bytes) = Formats().First();
        var labelled = WithLabelInfo(WithProperties(bytes, enabled: "true"), "enabled=\"1\" method=\"Privileged\"");

        var node = Assert.Single(client.Inspect(new StreamHandle(new MemoryStream(labelled)))
            .Nodes, n => n.Kind == "sensitivityLabel");
        Assert.Equal("Confidential", node.Name);
    }

    [Theory]
    [InlineData(Fidelity.Outline)]
    [InlineData(Fidelity.Structure)]
    [InlineData(Fidelity.Content)]
    public void A_label_is_reported_at_every_fidelity(Fidelity fidelity)
    {
        // A policy check must not pass because it asked for less detail.
        foreach (var (client, bytes) in Formats())
        {
            var labelled = WithProperties(bytes, enabled: "true");

            var nodes = client.Inspect(
                new StreamHandle(new MemoryStream(labelled)), new InspectOptions { Fidelity = fidelity }).Nodes;

            Assert.Single(nodes, n => n.Kind == "sensitivityLabel");
        }
    }

    [Fact]
    public void Metadata_that_is_not_readable_does_not_block_the_document()
    {
        // The part is unrelated to any edit, and it opened fine before labels were read.
        var (client, bytes) = Formats().First();
        var damaged = Edit(bytes, package =>
        {
            var part = ((WordprocessingDocument)package).AddCustomFilePropertiesPart();
            using var stream = part.GetStream(FileMode.Create);
            var content = Encoding.UTF8.GetBytes("<!DOCTYPE x><not-closed>");
            stream.Write(content, 0, content.Length);
        });

        var inspection = client.Inspect(new StreamHandle(new MemoryStream(damaged), "c.docx"));

        Assert.DoesNotContain(inspection.Nodes, n => n.Kind == "sensitivityLabel");
    }

    [Fact]
    public void A_removed_label_and_an_unlabelled_document_have_no_node()
    {
        var (client, bytes) = Formats().First();

        foreach (var document in new[]
                 {
                     bytes,
                     WithProperties(bytes, enabled: "false"),
                     WithLabelInfo(bytes, "enabled=\"1\" method=\"Standard\" removed=\"1\"")
                 })
            Assert.DoesNotContain(
                client.Inspect(new StreamHandle(new MemoryStream(document))).Nodes,
                n => n.Kind == "sensitivityLabel");
    }

    [Fact]
    public void Editing_a_word_document_leaves_its_label_byte_identical()
    {
        var client = new OfficeAgentClient(new WordModule());
        var labelled = WithLabelInfo(WithProperties(DocxFactory.Contract(), "true"), "enabled=\"1\" method=\"Standard\"");
        var hit = client.Find(new StreamHandle(new MemoryStream(labelled), "c.docx"), new FindQuery { Pattern = "Acme Corp" })[0];

        AssertLabelKept(labelled, Commit(client, labelled, "c.docx", DocFormat.Word, new ChangeTextOp
        {
            Target = hit.Anchor, With = "Globex", Mode = ChangeMode.Direct
        }));
    }

    [Fact]
    public void Editing_a_workbook_leaves_its_label_byte_identical()
    {
        var client = new OfficeAgentClient(new ExcelModule());
        var labelled = WithLabelInfo(WithProperties(new ExcelModule().CreateBlank(), "true"), "enabled=\"1\" method=\"Standard\"");

        AssertLabelKept(labelled, Commit(client, labelled, "c.xlsx", DocFormat.Excel, new SetCellOp
        {
            Target = new CellAnchor { SheetId = 1, Address = "A1" },
            Value = "Kept",
            ValueKind = SpreadsheetCellValueKind.String
        }));
    }

    [Fact]
    public void Editing_a_deck_leaves_its_label_byte_identical()
    {
        var client = new OfficeAgentClient(new PowerPointModule());
        var labelled = WithLabelInfo(WithProperties(PptxFactory.Deck(), "true"), "enabled=\"1\" method=\"Standard\"");
        var hit = client.Find(new StreamHandle(new MemoryStream(labelled), "c.pptx"), new FindQuery { Pattern = "Acme Corp" })[0];

        AssertLabelKept(labelled, Commit(client, labelled, "c.pptx", DocFormat.PowerPoint, new ChangeTextOp
        {
            Target = hit.Anchor, With = "Globex", Mode = ChangeMode.Direct
        }));
    }

    [Fact]
    public void No_operation_can_target_a_label()
    {
        var client = new OfficeAgentClient(new WordModule());
        var labelled = WithProperties(DocxFactory.Contract(), "true");

        foreach (var target in new[]
                 {
                     new NodeAnchor { Kind = "docProperty", Path = $"custom/MSIP_Label_{LabelId}_Name" },
                     new NodeAnchor { Kind = "sensitivityLabel", Path = $"sensitivityLabel#{LabelId}" }
                 })
        {
            var report = client.Preview(new StreamHandle(new MemoryStream(labelled), "c.docx"), new DocumentPlan
            {
                Format = DocFormat.Word,
                Operations = new PlanOperation[] { new SetPropertyOp { Target = target, Value = "Public" } }
            });

            Assert.NotEmpty(report.Errors);
        }
    }

    // ── helpers ───────────────────────────────────────────────────────────

    private static IEnumerable<(OfficeAgentClient Client, byte[] Bytes)> Formats()
    {
        yield return (new OfficeAgentClient(new WordModule()), DocxFactory.Contract());
        yield return (new OfficeAgentClient(new ExcelModule()), new ExcelModule().CreateBlank());
        yield return (new OfficeAgentClient(new PowerPointModule()), PptxFactory.Deck());
    }

    private static byte[] Commit(
        OfficeAgentClient client, byte[] bytes, string name, DocFormat format, PlanOperation operation)
    {
        using var result = client.Commit(new StreamHandle(new MemoryStream(bytes), name), new DocumentPlan
        {
            Format = format,
            Operations = new[] { operation }
        });
        Assert.True(result.Committed, string.Join("; ", result.Report.Errors.Select(e => e.Message)));
        return result.ToBytes();
    }

    private static void AssertLabelKept(byte[] before, byte[] after)
    {
        Assert.NotEqual(before, after);
        Assert.Equal(Entry(before, "docProps/custom.xml"), Entry(after, "docProps/custom.xml"));
        Assert.Equal(LabelInfo(before), LabelInfo(after));
    }

    private static string Entry(byte[] package, string name)
    {
        using var zip = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        using var reader = new StreamReader(zip.GetEntry(name)!.Open());
        return reader.ReadToEnd();
    }

    /// <summary>The label part wherever the fixture's SDK put it; Office names it docMetadata/LabelInfo.xml.</summary>
    private static string LabelInfo(byte[] package)
    {
        using var zip = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        foreach (var entry in zip.Entries)
        {
            using var reader = new StreamReader(entry.Open());
            var text = reader.ReadToEnd();
            if (text.Contains("clbl:labelList", StringComparison.Ordinal)) return text;
        }

        throw new InvalidOperationException(
            "The package has no label part: " + string.Join(", ", zip.Entries.Select(e => e.FullName)));
    }

    private static byte[] Edit(byte[] bytes, Action<OpenXmlPackage> change)
    {
        var buffer = new MemoryStream();
        buffer.Write(bytes, 0, bytes.Length);
        using (var package = Open(buffer)) change(package);
        return buffer.ToArray();
    }

    private static OpenXmlPackage Open(Stream stream)
    {
        using var zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
        var names = zip.Entries.Select(e => e.FullName).ToList();
        stream.Position = 0;
        if (names.Any(n => n.StartsWith("word/", StringComparison.Ordinal))) return WordprocessingDocument.Open(stream, true);
        if (names.Any(n => n.StartsWith("xl/", StringComparison.Ordinal))) return SpreadsheetDocument.Open(stream, true);
        return PresentationDocument.Open(stream, true);
    }

    private static byte[] WithProperties(byte[] bytes, string enabled) => Edit(bytes, package =>
    {
        var part = package switch
        {
            WordprocessingDocument word => word.AddCustomFilePropertiesPart(),
            SpreadsheetDocument book => book.AddCustomFilePropertiesPart(),
            _ => ((PresentationDocument)package).AddCustomFilePropertiesPart()
        };

        var properties = new Properties();
        var id = 2;
        foreach (var (field, value) in new[] { ("Enabled", enabled), ("Name", "Confidential"), ("Method", "Privileged") })
            properties.Append(new CustomDocumentProperty(new Vt.VTLPWSTR(value))
            {
                FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}",
                PropertyId = id++,
                Name = $"MSIP_Label_{LabelId}_{field}"
            });
        part.Properties = properties;
    });

    private static byte[] WithLabelInfo(byte[] bytes, string attributes) => Edit(bytes, package =>
    {
        var part = package.AddExtendedPart(
            "http://schemas.microsoft.com/office/2020/02/relationships/classificationlabels",
            "application/vnd.ms-office.classificationlabels+xml",
            "xml");

        using var stream = part.GetStream(FileMode.Create);
        var xml = "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                  "<clbl:labelList xmlns:clbl=\"http://schemas.microsoft.com/office/2020/mipLabelMetadata\">" +
                  $"<clbl:label id=\"{{{LabelId.ToUpperInvariant()}}}\" {attributes} siteId=\"{{00000000-0000-0000-0000-000000000000}}\"/>" +
                  "</clbl:labelList>";
        var content = Encoding.UTF8.GetBytes(xml);
        stream.Write(content, 0, content.Length);
    });
}
