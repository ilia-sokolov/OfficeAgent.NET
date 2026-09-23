using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using W = DocumentFormat.OpenXml.Wordprocessing;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;

namespace OfficeAgent.Tests;

/// <summary>
/// The package and meaning layers of the native corpus. Office itself is the third layer, run
/// by <c>scripts/native/office_check.ps1</c> against what the generation test writes.
/// </summary>
public sealed class NativeCorpusTests
{
    private static readonly Lazy<IReadOnlyList<NativeCorpus.Case>> Cases = new(NativeCorpus.Build);

    /// <summary>Every corpus case id, one theory row each.</summary>
    public static IEnumerable<object[]> Ids() => Cases.Value.Select(c => new object[] { c.Id });

    private static NativeCorpus.Case Get(string id) => Cases.Value.Single(c => c.Id == id);

    /// <summary>Case ids name files and results, so they must be unique.</summary>
    [Fact]
    public void Ids_are_unique_and_every_format_is_covered()
    {
        Assert.Equal(Cases.Value.Count, Cases.Value.Select(c => c.Id).Distinct().Count());
        Assert.Contains(Cases.Value, c => c.Format == "word");
        Assert.Contains(Cases.Value, c => c.Format == "powerpoint");
        Assert.Contains(Cases.Value, c => c.Format == "excel");
    }

    /// <summary>
    /// Every verb discovery advertises has at least one corpus case, so a newly advertised
    /// verb cannot ship without native evidence.
    /// </summary>
    [Theory]
    [InlineData("word")]
    [InlineData("powerpoint")]
    [InlineData("excel")]
    public void Every_advertised_operation_has_a_case(string format)
    {
        var module = format switch
        {
            "word" => (OfficeAgent.Core.IFormatModule)new OfficeAgent.Word.WordModule(),
            "powerpoint" => new OfficeAgent.PowerPoint.PowerPointModule(),
            _ => new OfficeAgent.Excel.ExcelModule()
        };
        var advertised = CapabilityDiscoveryTests.Describe(module).Operations;
        var covered = Cases.Value.Where(c => c.Format == format).Select(c => c.Family).ToHashSet(StringComparer.Ordinal);
        Assert.DoesNotContain(advertised, op => !covered.Contains(op));
    }

    /// <summary>Layer 1: the output introduces no schema error the input did not already have.</summary>
    [Theory]
    [MemberData(nameof(Ids))]
    public void Output_adds_no_schema_errors(string id)
    {
        var c = Get(id);
        var added = SchemaErrors(c.Format, c.Output).Except(SchemaErrors(c.Format, c.Input)).ToArray();
        Assert.True(added.Length == 0, string.Join("\n", added));
    }

    /// <summary>Layer 2: only the parts the case may touch changed, and the meaning is as expected.</summary>
    [Theory]
    [MemberData(nameof(Ids))]
    public void Output_changes_only_allowed_parts_and_means_what_was_asked(string id)
    {
        var c = Get(id);
        if (!c.AllowedChangedParts.Contains("*"))
        {
            var unexpected = WordPreservationVerifier.ChangedParts(c.Input, c.Output)
                .Where(part => !c.AllowedChangedParts.Contains(part)).ToArray();
            Assert.True(unexpected.Length == 0, "Unexpected changed parts: " + string.Join(", ", unexpected));
        }

        var text = VisibleText(c.Format, c.Output);
        var tracked = c.Expect.ContainsKey("revisionsAtLeast") || c.Expect.ContainsKey("revisionAuthors");
        foreach (var (key, value) in c.Expect)
        {
            switch (key)
            {
                case "textContains" when !tracked:
                    foreach (var s in (string[])value) Assert.Contains(s, text);
                    break;
                case "textExcludes" when !tracked:
                    foreach (var s in (string[])value) Assert.DoesNotContain(s, text);
                    break;
                case "revisions":
                    Assert.Equal((int)value, Revisions(c.Output).Count);
                    break;
                case "revisionsAtLeast":
                    Assert.True(Revisions(c.Output).Count >= (int)value, $"{Revisions(c.Output).Count} revisions");
                    break;
                case "revisionAuthors":
                    Assert.Equal((string[])value, Revisions(c.Output).Distinct().Order(StringComparer.Ordinal).ToArray());
                    break;
                case "tables":
                    Assert.Equal((int)value, Count(c.Format, c.Output, "tbl"));
                    break;
                case "slides" when c.Format == "powerpoint":
                    using (var deck = PresentationDocument.Open(new MemoryStream(c.Output), false))
                        Assert.Equal((int)value, deck.PresentationPart!.Presentation!.SlideIdList!.Count());
                    break;
                case "charts":
                    Assert.Equal((int)value, Count(c.Format, c.Output, "chart"));
                    break;
            }
        }
    }

    /// <summary>Every published refusal id, one theory row each.</summary>
    public static IEnumerable<object[]> RefusalIds() => NativeCorpus.Refusals().Select(r => new object[] { r.Id });

    /// <summary>
    /// A published unsupported case is refused with its code before anything is written, so
    /// the gap is a diagnostic rather than a damaged file.
    /// </summary>
    [Theory]
    [MemberData(nameof(RefusalIds))]
    public void Unsupported_cases_are_refused_with_their_code(string id)
    {
        var refusal = NativeCorpus.Refusals().Single(r => r.Id == id);
        var module = refusal.Format switch
        {
            "word" => (OfficeAgent.Core.IFormatModule)new OfficeAgent.Word.WordModule(),
            "powerpoint" => new OfficeAgent.PowerPoint.PowerPointModule(),
            _ => new OfficeAgent.Excel.ExcelModule()
        };
        var extension = refusal.Format switch { "word" => ".docx", "powerpoint" => ".pptx", _ => ".xlsx" };
        var client = new OfficeAgent.Core.OfficeAgentClient(module);
        var before = NativeCorpus.Sha256(refusal.Input);
        using var result = client.Commit(new StreamHandle(new MemoryStream(refusal.Input, writable: false), "input" + extension),
            new DocumentPlan { Operations = refusal.Operations });
        Assert.False(result.Committed);
        Assert.Contains(result.Report.Errors, e => e.Code == refusal.Code);
        Assert.Equal(before, NativeCorpus.Sha256(refusal.Input));
    }

    /// <summary>
    /// The same refusals through the filesystem provider: the stored file keeps its exact bytes
    /// and no sibling, temporary or new-version file appears.
    /// </summary>
    [Theory]
    [MemberData(nameof(RefusalIds))]
    public async Task Unsupported_cases_leave_the_stored_file_and_folder_untouched(string id)
    {
        var refusal = NativeCorpus.Refusals().Single(r => r.Id == id);
        var extension = refusal.Format switch { "word" => ".docx", "powerpoint" => ".pptx", _ => ".xlsx" };
        var root = Path.Combine(Path.GetTempPath(), $"officeagent-refusal-{Guid.NewGuid():N}");
        Directory.CreateDirectory(root);
        try
        {
            var path = Path.Combine(root, "original" + extension);
            await File.WriteAllBytesAsync(path, refusal.Input);
            var client = new OfficeAgentClient(
                new DocumentProviderRegistry(new[]
                {
                    new FileSystemDocumentProvider(new FileSystemDocumentProviderOptions { ConnectionId = "workspace", RootPath = root, AllowedExtensions = new[] { ".docx", ".pptx", ".xlsx" } })
                }),
                new OfficeAgent.Word.WordModule(), new OfficeAgent.PowerPoint.PowerPointModule(), new OfficeAgent.Excel.ExcelModule());
            var reference = await client.RegisterAsync("workspace", path);
            var before = Directory.GetFiles(root, "*", SearchOption.AllDirectories).Order(StringComparer.Ordinal).ToArray();

            var result = await client.CommitAsync(reference, new DocumentPlan { Operations = refusal.Operations });

            Assert.False(result.Committed);
            Assert.Contains(result.Report.Errors, e => e.Code == refusal.Code);
            Assert.Equal(NativeCorpus.Sha256(refusal.Input), NativeCorpus.Sha256(await File.ReadAllBytesAsync(path)));
            Assert.Equal(before, Directory.GetFiles(root, "*", SearchOption.AllDirectories).Order(StringComparer.Ordinal).ToArray());
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    /// <summary>
    /// Writes each case's input, output and a manifest for the native harness. Runs only when
    /// <c>OFFICEAGENT_NATIVE_CORPUS_OUTPUT</c> names a directory.
    /// </summary>
    [Fact]
    public void Writes_the_corpus_for_native_verification_when_asked()
    {
        var root = Environment.GetEnvironmentVariable("OFFICEAGENT_NATIVE_CORPUS_OUTPUT");
        if (string.IsNullOrWhiteSpace(root)) return;
        Directory.CreateDirectory(Path.Combine(root, "inputs"));
        Directory.CreateDirectory(Path.Combine(root, "outputs"));
        var manifest = new List<object>();
        foreach (var c in Cases.Value)
        {
            File.WriteAllBytes(Path.Combine(root, "inputs", c.FileName), c.Input);
            File.WriteAllBytes(Path.Combine(root, "outputs", c.FileName), c.Output);
            manifest.Add(new
            {
                id = c.Id, format = c.Format, family = c.Family, description = c.Description, file = c.FileName,
                inputSha256 = NativeCorpus.Sha256(c.Input), outputSha256 = NativeCorpus.Sha256(c.Output),
                visual = c.ExportForVisualReview, expect = c.Expect
            });
        }
        // Negative controls for the harness: a file Office must repair or refuse. A harness that
        // reports these as clean cannot detect repair, and its passes mean nothing.
        Directory.CreateDirectory(Path.Combine(root, "controls"));
        var controls = new (string File, byte[] Bytes, string Part, string Bad)[]
        {
            ("control-corrupt.docx", DocxFactory.Contract(), "word/document.xml", "<w:body><w:p><w:r><w:t>unclosed"),
            ("control-corrupt.pptx", PptxFactory.Deck(), "ppt/slides/slide1.xml", "<p:cSld><p:spTree><p:sp>unclosed"),
            ("control-corrupt.xlsx", new OfficeAgent.Excel.ExcelModule().CreateBlank(), "xl/worksheets/sheet1.xml", "<sheetData><row r=\"1\"><c r=\"A1\">unclosed")
        };
        foreach (var (file, bytes, part, bad) in controls)
            File.WriteAllBytes(Path.Combine(root, "controls", file), Corrupt(bytes, part, bad));
        // LF on every platform: JsonSerializer indents with Environment.NewLine on .NET 8, which
        // once made the published manifest CRLF, and its hash is what native results are bound to.
        File.WriteAllText(Path.Combine(root, "manifest.json"),
            JsonSerializer.Serialize(manifest, new JsonSerializerOptions { WriteIndented = true }).Replace("\r\n", "\n"));
    }

    private static string PublishedRoot() =>
        Path.Combine(RepositoryRoot(), "tests", "OfficeAgent.Tests", "Corpus", "v1.0.0", "native");

    /// <summary>
    /// The published native record applies to exactly the published corpus: every binding
    /// between results, manifest and files holds, and every operation discovery advertises today,
    /// and every case the corpus builds, appears in it. A new verb or case therefore fails here
    /// until native evidence for it is recorded.
    /// </summary>
    [Fact]
    public void The_published_native_record_matches_its_files_and_covers_every_advertised_operation()
    {
        var root = PublishedRoot();
        var problems = NativeRecordProblems(root);
        Assert.True(problems.Count == 0, string.Join("\n", problems));

        using var manifest = JsonDocument.Parse(File.ReadAllBytes(Path.Combine(root, "manifest.json")));
        var families = new Dictionary<string, HashSet<string>>(StringComparer.Ordinal);
        foreach (var entry in manifest.RootElement.EnumerateArray())
        {
            var format = entry.GetProperty("format").GetString()!;
            if (!families.TryGetValue(format, out var set)) families[format] = set = new HashSet<string>(StringComparer.Ordinal);
            set.Add(entry.GetProperty("family").GetString()!);
        }

        // Every case the corpus builds today was natively verified, so a new case needs a new run.
        var published = manifest.RootElement.EnumerateArray().Select(e => e.GetProperty("id").GetString()).ToHashSet(StringComparer.Ordinal);
        Assert.DoesNotContain(Cases.Value, c => !published.Contains(c.Id));

        foreach (var (format, module) in new (string, OfficeAgent.Core.IFormatModule)[]
                 {
                     ("word", new OfficeAgent.Word.WordModule()),
                     ("powerpoint", new OfficeAgent.PowerPoint.PowerPointModule()),
                     ("excel", new OfficeAgent.Excel.ExcelModule())
                 })
            Assert.DoesNotContain(CapabilityDiscoveryTests.Describe(module).Operations, op => !families[format].Contains(op));
    }

    /// <summary>
    /// The binding guard fails closed. Each tamper, applied to a temporary copy of the published
    /// corpus, is one a regeneration could leave behind while keeping the previous passing
    /// results: none of them may pass.
    /// </summary>
    /// <remarks>
    /// Each case names the problem it must produce. Asserting only that some problem appears let
    /// every case pass for an unrelated reason: the tampered JSON was once rewritten with CRLF and
    /// rejected for that, so removing the binding under test went unnoticed.
    /// </remarks>
    [Theory]
    [InlineData("output-regenerated-with-its-hash", "Office opened")]
    [InlineData("results-manifest-hash", "results were recorded against manifest")]
    [InlineData("result-case-id", "no native result for")]
    [InlineData("result-output-hash", "Office opened")]
    [InlineData("result-case-duplicated", "results list")]
    [InlineData("check-failed-under-a-passing-case", "check '")]
    [InlineData("control-missing", "controls recorded")]
    [InlineData("result-for-another-operation", "result family differs")]
    [InlineData("control-checked-by-the-wrong-application", "checked by the wrong application")]
    public void The_native_record_guard_rejects_a_stale_or_altered_record(string tamper, string expectedProblem)
    {
        var copy = Path.Combine(Path.GetTempPath(), $"officeagent-native-record-{Guid.NewGuid():N}");
        CopyDirectory(PublishedRoot(), copy);
        try
        {
            Assert.Empty(NativeRecordProblems(copy));
            var resultsPath = Path.Combine(copy, "results.json");
            var manifestPath = Path.Combine(copy, "manifest.json");
            var results = System.Text.Json.Nodes.JsonNode.Parse(File.ReadAllText(resultsPath))!;
            var cases = results["cases"]!.AsArray();

            switch (tamper)
            {
                case "output-regenerated-with-its-hash":
                {
                    // A regeneration: one output changes and the manifest follows it; the old
                    // native results are kept.
                    var manifest = System.Text.Json.Nodes.JsonNode.Parse(File.ReadAllText(manifestPath))!.AsArray();
                    var file = manifest[0]!["file"]!.GetValue<string>();
                    var output = Path.Combine(copy, "outputs", file);
                    var bytes = File.ReadAllBytes(output);
                    bytes[^1] ^= 0xFF;
                    File.WriteAllBytes(output, bytes);
                    manifest[0]!["outputSha256"] = NativeCorpus.Sha256(bytes);
                    File.WriteAllText(manifestPath, manifest.ToJsonString(new JsonSerializerOptions { WriteIndented = true }).Replace("\r\n", "\n"));
                    break;
                }
                case "results-manifest-hash":
                    results["environment"]!["manifestSha256"] = new string('0', 64);
                    break;
                case "result-case-id":
                    cases[0]!["id"] = "no-such-case";
                    break;
                case "result-output-hash":
                    cases[0]!["sha256"] = new string('0', 64);
                    break;
                case "result-case-duplicated":
                    cases[1] = cases[0]!.DeepClone();
                    break;
                case "check-failed-under-a-passing-case":
                    cases[0]!["checks"]!.AsArray()[0]!["pass"] = false;
                    break;
                case "control-missing":
                    results["controls"]!.AsArray().RemoveAt(0);
                    break;
                case "result-for-another-operation":
                    cases[0]!["family"] = "someOtherOperation";
                    break;
                case "control-checked-by-the-wrong-application":
                {
                    var control = results["controls"]!.AsArray()[0]!;
                    control["format"] = control["format"]!.GetValue<string>() == "excel" ? "word" : "excel";
                    break;
                }
            }
            if (tamper != "output-regenerated-with-its-hash")
                File.WriteAllText(resultsPath, results.ToJsonString(new JsonSerializerOptions { WriteIndented = true }).Replace("\r\n", "\n"));

            var problems = NativeRecordProblems(copy);
            Assert.Contains(problems, problem => problem.Contains(expectedProblem, StringComparison.Ordinal));
        }
        finally
        {
            Directory.Delete(copy, recursive: true);
        }
    }

    /// <summary>
    /// Every way the native results could fail to describe the corpus beside them. Empty means
    /// the results were produced by a run over exactly these files.
    /// </summary>
    internal static IReadOnlyList<string> NativeRecordProblems(string root)
    {
        var problems = new List<string>();
        var manifestBytes = File.ReadAllBytes(Path.Combine(root, "manifest.json"));
        var resultsBytes = File.ReadAllBytes(Path.Combine(root, "results.json"));
        if (manifestBytes.Contains((byte)'\r')) problems.Add("manifest.json contains carriage returns; the corpus writes LF");
        if (resultsBytes.Contains((byte)'\r')) problems.Add("results.json contains carriage returns; the corpus writes LF");
        using var manifest = JsonDocument.Parse(manifestBytes);
        using var results = JsonDocument.Parse(resultsBytes);

        // 1. The results were produced against these exact manifest bytes.
        var recordedManifest = results.RootElement.GetProperty("environment").GetProperty("manifestSha256").GetString();
        if (recordedManifest != NativeCorpus.Sha256(manifestBytes))
            problems.Add($"results were recorded against manifest {recordedManifest}, not this manifest {NativeCorpus.Sha256(manifestBytes)}");

        // 2-4. Unique ids on both sides, and the same set.
        var entries = manifest.RootElement.EnumerateArray().ToList();
        var manifestIds = entries.Select(e => e.GetProperty("id").GetString()!).ToList();
        foreach (var duplicate in manifestIds.GroupBy(id => id).Where(g => g.Count() > 1))
            problems.Add($"manifest lists {duplicate.Key} {duplicate.Count()} times");
        var cases = results.RootElement.GetProperty("cases").EnumerateArray().ToList();
        var resultIds = cases.Select(c => c.GetProperty("id").GetString()!).ToList();
        foreach (var duplicate in resultIds.GroupBy(id => id).Where(g => g.Count() > 1))
            problems.Add($"results list {duplicate.Key} {duplicate.Count()} times");
        foreach (var missing in manifestIds.Except(resultIds))
            problems.Add($"no native result for {missing}");
        foreach (var extra in resultIds.Except(manifestIds))
            problems.Add($"a native result for {extra}, which the manifest does not list");

        var byId = cases.GroupBy(c => c.GetProperty("id").GetString()!).ToDictionary(g => g.Key, g => g.First());
        foreach (var entry in entries)
        {
            var id = entry.GetProperty("id").GetString()!;
            var file = entry.GetProperty("file").GetString()!;
            var outputSha = entry.GetProperty("outputSha256").GetString()!;

            // The manifest names the files beside it.
            var input = Path.Combine(root, "inputs", file);
            var output = Path.Combine(root, "outputs", file);
            if (!File.Exists(input) || entry.GetProperty("inputSha256").GetString() != NativeCorpus.Sha256(File.ReadAllBytes(input)))
                problems.Add($"{id}: input does not match the manifest");
            var actualOutput = File.Exists(output) ? NativeCorpus.Sha256(File.ReadAllBytes(output)) : "(missing)";
            if (outputSha != actualOutput)
                problems.Add($"{id}: output {actualOutput} does not match the manifest {outputSha}");

            if (!byId.TryGetValue(id, out var result)) continue;

            // 5. Same case.
            foreach (var field in new[] { "format", "family" })
                if (result.GetProperty(field).GetString() != entry.GetProperty(field).GetString())
                    problems.Add($"{id}: result {field} differs from the manifest");

            // 6-7. Office opened exactly this output.
            if (result.GetProperty("manifestSha256").GetString() != outputSha)
                problems.Add($"{id}: result was run against manifest output {result.GetProperty("manifestSha256").GetString()}");
            if (result.GetProperty("sha256").GetString() != outputSha || result.GetProperty("sha256").GetString() != actualOutput)
                problems.Add($"{id}: Office opened {result.GetProperty("sha256").GetString()}, not the published output");

            // 8. Every check passed, not only the case flag.
            if (!result.GetProperty("pass").GetBoolean()) problems.Add($"{id}: case failed");
            if (!result.GetProperty("openedCleanly").GetBoolean()) problems.Add($"{id}: not opened cleanly");
            var checks = result.GetProperty("checks").EnumerateArray().ToList();
            if (checks.Count == 0) problems.Add($"{id}: no native checks recorded");
            foreach (var check in checks.Where(check => !check.GetProperty("pass").GetBoolean()))
                problems.Add($"{id}: check '{check.GetProperty("name").GetString()}' failed");
        }

        // 9. The controls are exactly the corrupt files beside the corpus, each caught by the
        // application its extension names.
        var extensionFormat = new Dictionary<string, string> { [".docx"] = "word", [".pptx"] = "powerpoint", [".xlsx"] = "excel" };
        var controlFiles = Directory.GetFiles(Path.Combine(root, "controls")).Select(Path.GetFileName).OrderBy(n => n, StringComparer.Ordinal).ToList();
        var controls = results.RootElement.GetProperty("controls").EnumerateArray().ToList();
        var recordedControls = controls.Select(c => c.GetProperty("file").GetString()!).ToList();
        if (!recordedControls.OrderBy(n => n, StringComparer.Ordinal).SequenceEqual(controlFiles))
            problems.Add($"controls recorded [{string.Join(", ", recordedControls)}] are not the control files [{string.Join(", ", controlFiles)}]");
        foreach (var control in controls)
        {
            var file = control.GetProperty("file").GetString()!;
            if (!control.GetProperty("detected").GetBoolean()) problems.Add($"control {file} was not caught");
            if (!control.GetProperty("appOpenedValidCase").GetBoolean()) problems.Add($"control {file}: its application opened no valid case in the run");
            if (!extensionFormat.TryGetValue(Path.GetExtension(file), out var format) || control.GetProperty("format").GetString() != format)
                problems.Add($"control {file} was checked by the wrong application");
        }
        return problems;
    }

    private static void CopyDirectory(string source, string destination)
    {
        foreach (var directory in Directory.GetDirectories(source, "*", SearchOption.AllDirectories))
            Directory.CreateDirectory(directory.Replace(source, destination));
        Directory.CreateDirectory(destination);
        foreach (var file in Directory.GetFiles(source, "*", SearchOption.AllDirectories))
            File.Copy(file, file.Replace(source, destination));
    }

    private static string RepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "OfficeAgent.NET.sln")))
            directory = directory.Parent;
        return directory?.FullName ?? throw new InvalidOperationException("Repository root not found.");
    }

    private static byte[] Corrupt(byte[] package, string part, string body)
    {
        using var stream = new MemoryStream();
        stream.Write(package);
        using (var zip = new System.IO.Compression.ZipArchive(stream, System.IO.Compression.ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = zip.GetEntry(part) ?? throw new InvalidOperationException(part + " is not in the package.");
            string xml;
            using (var reader = new StreamReader(entry.Open())) xml = reader.ReadToEnd();
            entry.Delete();
            var root = xml.IndexOf('>', xml.IndexOf("?>", StringComparison.Ordinal) + 2) + 1;
            using var writer = new StreamWriter(zip.CreateEntry(part).Open());
            writer.Write(xml[..root] + body);
        }
        return stream.ToArray();
    }

    private static string[] SchemaErrors(string format, byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using OpenXmlPackage package = format switch
        {
            "word" => WordprocessingDocument.Open(stream, false),
            "powerpoint" => PresentationDocument.Open(stream, false),
            _ => SpreadsheetDocument.Open(stream, false)
        };
        return new OpenXmlValidator(FileFormatVersions.Office2019).Validate(package)
            .Select(e => $"{e.Part?.Uri} {e.Path?.XPath} {e.Id}: {e.Description}")
            .Distinct().ToArray();
    }

    private static string VisibleText(string format, byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        switch (format)
        {
            case "word":
                using (var document = WordprocessingDocument.Open(stream, false))
                    return string.Join("\n", document.MainDocumentPart!.Document!.Body!.Descendants<W.Paragraph>()
                        .Select(p => string.Concat(p.Descendants<W.Text>().Select(t => t.Text))));
            case "powerpoint":
                using (var deck = PresentationDocument.Open(stream, false))
                    return string.Join("\n", deck.PresentationPart!.SlideParts
                        .SelectMany(s => s.Slide!.Descendants<A.Paragraph>())
                        .Select(p => string.Concat(p.Descendants<A.Text>().Select(t => t.Text))));
            default:
                using (var book = SpreadsheetDocument.Open(stream, false))
                    return string.Join("|", book.WorkbookPart!.Parts.Select(p => p.OpenXmlPart.RootElement?.InnerText ?? "")
                        .Concat(book.WorkbookPart.WorksheetParts.SelectMany(w => w.Worksheet!.InnerText is { } t ? new[] { t } : Array.Empty<string>())));
        }
    }

    /// <summary>The author of every body revision mark Word counts as a revision.</summary>
    private static List<string> Revisions(byte[] bytes)
    {
        using var document = WordprocessingDocument.Open(new MemoryStream(bytes, writable: false), false);
        return document.MainDocumentPart!.Document!.Body!.Descendants()
            .Where(e => e.NamespaceUri == WordNamespace &&
                (RevisionMarks.Contains(e.LocalName) && e.Parent?.LocalName != "rPr" || e.LocalName is "rPrChange" or "pPrChange"))
            .Select(e => e.GetAttribute("author", WordNamespace).Value ?? "")
            .ToList();
    }

    // w:ins and w:del wrap runs or mark paragraph marks and rows; rPrChange and pPrChange are
    // formatting revisions. A paragraph-mark w:ins inside rPr counts once, with its run.
    private const string WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly HashSet<string> RevisionMarks = new(StringComparer.Ordinal) { "ins", "del", "moveFrom", "moveTo" };

    private static int Count(string format, byte[] bytes, string localName)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using OpenXmlPackage package = format switch
        {
            "word" => WordprocessingDocument.Open(stream, false),
            "powerpoint" => PresentationDocument.Open(stream, false),
            _ => SpreadsheetDocument.Open(stream, false)
        };
        return format switch
        {
            "word" when localName == "tbl" => ((WordprocessingDocument)package).MainDocumentPart!.Document!.Body!.Elements<W.Table>().Count(),
            "powerpoint" when localName == "tbl" => ((PresentationDocument)package).PresentationPart!.SlideParts.Sum(s => s.Slide!.Descendants<A.Table>().Count()),
            "excel" when localName == "tbl" => ((SpreadsheetDocument)package).WorkbookPart!.WorksheetParts.Sum(w => w.TableDefinitionParts.Count()),
            _ when localName == "chart" => package.GetAllParts().OfType<ChartPart>().Count(),
            _ => throw new ArgumentOutOfRangeException(nameof(localName))
        };
    }
}
