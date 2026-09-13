using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Serialization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

public sealed class V08CorpusTests
{
    private static readonly string RepositoryRoot = FindRepositoryRoot();
    private static readonly string CorpusRoot = Path.Combine(
        RepositoryRoot, "tests", "OfficeAgent.Tests", "Corpus", "v0.8.0");

    [Fact]
    [Trait("Category", "Fixture")]
    public void Generate_versioned_corpus_only_when_an_output_root_is_explicit()
    {
        var outputRoot = Environment.GetEnvironmentVariable("OFFICEAGENT_V08_CORPUS_OUTPUT");
        if (string.IsNullOrWhiteSpace(outputRoot))
            return;

        Directory.CreateDirectory(outputRoot);
        foreach (var fixture in V08CorpusFixtureBuilder.BuildAll(RepositoryRoot))
            File.WriteAllBytes(Path.Combine(outputRoot, fixture.Key), fixture.Value);
    }

    [Fact]
    public void Generated_fixture_recipes_produce_valid_packages_with_the_committed_shape()
    {
        var manifest = ReadManifest();
        foreach (var generated in V08CorpusFixtureBuilder.BuildAll(RepositoryRoot))
        {
            var path = Path.Combine(CorpusRoot, generated.Key);
            Assert.True(File.Exists(path), $"Missing generated corpus fixture: {generated.Key}");
            var committed = File.ReadAllBytes(path);
            if (Path.GetExtension(path).Equals(".json", StringComparison.OrdinalIgnoreCase))
                Assert.Equal(Sha256(generated.Value), Sha256(committed));
            else
            {
                var format = Path.GetExtension(path).TrimStart('.').ToLowerInvariant();
                CorpusAssertions.AssertPackageValid(format, generated.Value);
                CorpusAssertions.AssertPackageShapeEqual(generated.Value, committed);
                foreach (var item in manifest.Cases.Where(item =>
                    Path.GetFileName(item.Fixture).Equals(generated.Key, StringComparison.Ordinal)))
                    CorpusAssertions.AssertSemantics(item, generated.Value);
            }
        }
    }

    [Fact]
    public void Manifest_ids_hashes_origins_and_files_are_valid()
    {
        var manifest = ReadManifest();
        Assert.Equal("1.0", manifest.SchemaVersion);
        Assert.Equal("0.8.0", manifest.ReleaseBaseline);
        Assert.Equal("f36e410a19b3c1d24bfa87d4485dce5dad78b8bb", manifest.SourceCommit);
        Assert.NotEmpty(manifest.Cases);
        Assert.Equal(manifest.Cases.Count,
            manifest.Cases.Select(item => item.Id).Distinct(StringComparer.Ordinal).Count());

        foreach (var item in manifest.Cases)
        {
            Assert.False(string.IsNullOrWhiteSpace(item.Id));
            Assert.False(string.IsNullOrWhiteSpace(item.Workflow));
            Assert.Contains(item.Disposition, new[] { "positive", "refusal", "preservation" });
            Assert.False(string.IsNullOrWhiteSpace(item.Origin.Kind));
            Assert.Contains(item.Origin.Kind, new[]
            {
                "generated-fictional",
                "generated-contract",
                "generated-configuration"
            });
            Assert.False(string.IsNullOrWhiteSpace(item.Origin.Recipe));
            Assert.Equal("MIT repository fixture", item.Origin.License);
            Assert.False(string.IsNullOrWhiteSpace(item.Format));
            Assert.False(string.IsNullOrWhiteSpace(item.AuthoringApplication.Name));
            Assert.False(string.IsNullOrWhiteSpace(item.AuthoringApplication.Version));
            Assert.NotEmpty(item.OperationSequence);
            Assert.NotEmpty(item.ExpectedStructure);
            Assert.NotNull(item.ExpectedText);
            Assert.NotNull(item.ProtectedPackageParts);
            Assert.NotNull(item.ExpectedRefusalCodes);

            var path = Path.GetFullPath(Path.Combine(RepositoryRoot, item.Fixture));
            Assert.True(path.StartsWith(RepositoryRoot + Path.DirectorySeparatorChar,
                StringComparison.OrdinalIgnoreCase), $"Fixture escapes the repository: {item.Fixture}");
            Assert.True(File.Exists(path), $"Missing corpus fixture: {item.Fixture}");
            Assert.Equal(item.Sha256, Sha256(File.ReadAllBytes(path)));
        }

        var generatedFiles = V08CorpusFixtureBuilder.BuildAll(RepositoryRoot).Keys
            .OrderBy(name => name, StringComparer.Ordinal)
            .ToArray();
        var manifestedFiles = manifest.Cases
            .Select(item => Path.GetFileName(item.Fixture))
            .Distinct(StringComparer.Ordinal)
            .OrderBy(name => name, StringComparer.Ordinal)
            .ToArray();
        Assert.Equal(generatedFiles, manifestedFiles);
    }

    [Fact]
    public void Manifest_covers_every_advertised_v08_workflow_with_a_positive_and_a_boundary_case()
    {
        var required = new[]
        {
            "word-creation-editing",
            "word-review",
            "powerpoint-creation-editing",
            "excel-inspection-editing",
            "template-generation",
            "word-comparison",
            "agent-application-integration",
            "document-access",
            "word-assembly",
            "optional-rendering"
        };
        var cases = ReadManifest().Cases;

        foreach (var workflow in required)
        {
            var workflowCases = cases.Where(item => item.Workflow == workflow).ToArray();
            Assert.Contains(workflowCases, item => item.Disposition == "positive");
            Assert.Contains(workflowCases,
                item => item.Disposition is "refusal" or "preservation");
        }
    }

    [Fact]
    public void Ooxml_fixtures_are_schema_valid_and_match_declared_semantics()
    {
        foreach (var item in ReadManifest().Cases
            .GroupBy(entry => entry.Fixture, StringComparer.Ordinal)
            .Select(group => group.First())
            .Where(item => item.Format is "docx" or "pptx" or "xlsx"))
        {
            var bytes = File.ReadAllBytes(Path.Combine(RepositoryRoot, item.Fixture));
            CorpusAssertions.AssertPackageValid(item.Format, bytes);
            CorpusAssertions.AssertSemantics(item, bytes);
        }
    }

    [Fact]
    public void Reusable_preservation_assertion_detects_changes_and_accepts_unchanged_parts()
    {
        var input = File.ReadAllBytes(Path.Combine(CorpusRoot, "complex-contract.docx"));
        var client = new OfficeAgentClient(new WordModule(new FixedTimeProvider(
            new DateTimeOffset(2026, 9, 12, 11, 0, 0, TimeSpan.Zero))));
        var paragraph = client.Inspect(input).Paragraphs.First(info =>
            info.Location == "body" && info.Text.Contains("twelve months", StringComparison.Ordinal));
        using var result = client.Commit(
            new StreamHandle(new MemoryStream(input, writable: false)),
            new DocumentPlan
            {
                Revision = new RevisionMetadata
                {
                    Author = "Corpus Verification",
                    TimestampUtc = new DateTimeOffset(2026, 9, 12, 11, 0, 0, TimeSpan.Zero)
                },
                Operations = new PlanOperation[]
                {
                    new ChangeTextOp
                    {
                        Target = new TextSpanAnchor
                        {
                            ParaId = paragraph.ParaId,
                            Expect = "twelve months"
                        },
                        With = "eighteen months",
                        Mode = ChangeMode.Tracked
                    }
                }
            });

        Assert.True(result.Committed, string.Join("; ", result.Report.Errors.Select(error => error.Message)));
        CorpusAssertions.AssertPartsPreserved(input, result.ToBytes(), new[]
        {
            "word/comments.xml",
            "word/styles.xml",
            "media/image.png"
        });
        Assert.ThrowsAny<Xunit.Sdk.XunitException>(() =>
            CorpusAssertions.AssertPartsPreserved(input, result.ToBytes(), new[] { "word/document.xml" }));
    }

    [Fact]
    public void Reusable_rejection_assertion_proves_no_output_or_input_mutation()
    {
        var input = File.ReadAllBytes(Path.Combine(CorpusRoot, "complex-contract.docx"));
        var client = new OfficeAgentClient(new WordModule());
        CorpusAssertions.AssertRejectedWithoutMutation(input, bytes => client.Commit(
            new StreamHandle(new MemoryStream(bytes, writable: false)),
            new DocumentPlan
            {
                Operations = new PlanOperation[]
                {
                    new ChangeTextOp
                    {
                        Target = new TextSpanAnchor
                        {
                            ParaId = "does-not-exist",
                            Expect = "missing"
                        },
                        With = "replacement"
                    }
                }
            }), ValidationErrorCodes.AnchorNotFound);
    }

    [Fact]
    public void V08_json_contract_fixtures_deserialize_and_keep_expected_boundaries()
    {
        var options = new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true,
            Converters = { new JsonStringEnumConverter() }
        };
        var plan = JsonSerializer.Deserialize<DocumentPlan>(
            File.ReadAllText(Path.Combine(CorpusRoot, "word-edit-plan.json")), options);
        Assert.NotNull(plan);
        Assert.Equal("0.2", plan.ContractVersion);
        Assert.Equal(OfficeAgent.Abstractions.DocumentFormat.Word, plan.Format);
        Assert.IsType<ChangeTextOp>(Assert.Single(plan.Operations));

        using var configuration = JsonDocument.Parse(File.ReadAllBytes(Path.Combine(CorpusRoot, "mcp-config.json")));
        Assert.True(configuration.RootElement.GetProperty("OfficeAgent").GetProperty("AllowCreation").GetBoolean());
        using var error = JsonDocument.Parse(File.ReadAllBytes(Path.Combine(CorpusRoot, "validation-error.json")));
        Assert.Equal("stale-snapshot", error.RootElement.GetProperty("errors")[0].GetProperty("code").GetString());
        using var receipt = JsonDocument.Parse(File.ReadAllBytes(Path.Combine(CorpusRoot, "apply-receipt.json")));
        Assert.Equal("fixture-user", receipt.RootElement.GetProperty("actor").GetProperty("subject").GetString());
    }

    private static CorpusManifest ReadManifest()
    {
        var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
        return JsonSerializer.Deserialize<CorpusManifest>(
            File.ReadAllText(Path.Combine(CorpusRoot, "manifest.json")), options)
            ?? throw new InvalidDataException("The v0.8 corpus manifest is empty.");
    }

    private static string FindRepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "OfficeAgent.NET.sln")))
            directory = directory.Parent;
        return directory?.FullName
            ?? throw new DirectoryNotFoundException("Could not locate the OfficeAgent.NET repository root.");
    }

    private static string Sha256(byte[] bytes) =>
        Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    private sealed class FixedTimeProvider : TimeProvider
    {
        private readonly DateTimeOffset _now;

        public FixedTimeProvider(DateTimeOffset now) => _now = now;

        public override DateTimeOffset GetUtcNow() => _now;
    }
}

internal static class CorpusAssertions
{
    public static void AssertPackageValid(string format, byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        var validator = new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019);
        var errors = format switch
        {
            "docx" => Validate(WordprocessingDocument.Open(stream, false), validator),
            "pptx" => Validate(PresentationDocument.Open(stream, false), validator),
            "xlsx" => Validate(SpreadsheetDocument.Open(stream, false), validator),
            _ => throw new ArgumentOutOfRangeException(nameof(format), format, "Unsupported corpus format.")
        };
        Assert.True(errors.Count == 0, string.Join("; ", errors));
    }

    public static void AssertSemantics(CorpusCase item, byte[] bytes)
    {
        var text = item.Format switch
        {
            "docx" => WordText(bytes),
            "pptx" => PresentationText(bytes),
            "xlsx" => SpreadsheetText(bytes),
            _ => string.Empty
        };
        foreach (var expected in item.ExpectedText)
            Assert.Contains(expected, text, StringComparison.Ordinal);

        using var stream = new MemoryStream(bytes, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var parts = archive.Entries.Select(entry => entry.FullName).ToHashSet(StringComparer.Ordinal);
        foreach (var expected in item.ExpectedStructure.Where(value => value.StartsWith("part:", StringComparison.Ordinal)))
            Assert.Contains(expected[5..], parts);
    }

    public static void AssertPartsPreserved(byte[] before, byte[] after, IEnumerable<string> parts)
    {
        var beforeHashes = PackagePartHashes(before);
        var afterHashes = PackagePartHashes(after);
        foreach (var part in parts)
        {
            Assert.True(beforeHashes.TryGetValue(part, out var beforeHash), $"Missing input part: {part}");
            Assert.True(afterHashes.TryGetValue(part, out var afterHash), $"Missing output part: {part}");
            Assert.True(beforeHash == afterHash,
                $"Protected package part changed: {part}. Before {beforeHash}; after {afterHash}.");
        }
    }

    public static void AssertPackageShapeEqual(byte[] expected, byte[] actual)
    {
        var expectedHashes = PackagePartHashes(expected);
        var actualHashes = PackagePartHashes(actual);
        var expectedParts = expectedHashes.Keys
            .Where(key => !key.EndsWith(".rels", StringComparison.Ordinal))
            .OrderBy(key => key, StringComparer.Ordinal)
            .ToArray();
        var actualParts = actualHashes.Keys
            .Where(key => !key.EndsWith(".rels", StringComparison.Ordinal))
            .OrderBy(key => key, StringComparer.Ordinal)
            .ToArray();
        Assert.Equal(expectedParts, actualParts);
    }

    public static void AssertRejectedWithoutMutation(
        byte[] input,
        Func<byte[], ApplyResult> apply,
        string expectedCode)
    {
        var before = Sha256(input);
        using var result = apply(input);
        Assert.False(result.Committed);
        Assert.Contains(result.Report.Errors, error => error.Code == expectedCode);
        Assert.Null(result.Output);
        Assert.Equal(before, Sha256(input));
    }

    private static IReadOnlyList<string> Validate(OpenXmlPackage package, OpenXmlValidator validator)
    {
        using (package)
            return validator.Validate(package)
                .Select(error => $"{error.Path?.XPath}: {error.Description}")
                .ToArray();
    }

    private static string WordText(byte[] bytes)
    {
        using var document = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        var main = document.MainDocumentPart!;
        return string.Join("|", new[]
        {
            main.Document.InnerText,
            main.WordprocessingCommentsPart?.Comments?.InnerText ?? string.Empty,
            main.FootnotesPart?.Footnotes?.InnerText ?? string.Empty,
            main.EndnotesPart?.Endnotes?.InnerText ?? string.Empty
        });
    }

    private static string PresentationText(byte[] bytes)
    {
        using var document = PresentationDocument.Open(new MemoryStream(bytes), false);
        return string.Join("|", document.PresentationPart!.Parts
            .Select(part => part.OpenXmlPart.RootElement?.InnerText ?? string.Empty));
    }

    private static string SpreadsheetText(byte[] bytes)
    {
        using var document = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        return string.Join("|", document.WorkbookPart!.Parts
            .Select(part => part.OpenXmlPart.RootElement?.InnerText ?? string.Empty));
    }

    private static IReadOnlyDictionary<string, string> PackagePartHashes(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        return archive.Entries.ToDictionary(
            entry => entry.FullName,
            entry =>
            {
                using var content = entry.Open();
                return Convert.ToHexString(SHA256.HashData(content)).ToLowerInvariant();
            },
            StringComparer.Ordinal);
    }

    private static string Sha256(byte[] bytes) =>
        Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
}

internal sealed class CorpusManifest
{
    public string SchemaVersion { get; init; } = string.Empty;
    public string ReleaseBaseline { get; init; } = string.Empty;
    public string SourceCommit { get; init; } = string.Empty;
    public IReadOnlyList<CorpusCase> Cases { get; init; } = Array.Empty<CorpusCase>();
}

internal sealed class CorpusCase
{
    public string Id { get; init; } = string.Empty;
    public string Workflow { get; init; } = string.Empty;
    public string Disposition { get; init; } = string.Empty;
    public string Fixture { get; init; } = string.Empty;
    public CorpusOrigin Origin { get; init; } = new();
    public string Sha256 { get; init; } = string.Empty;
    public string Format { get; init; } = string.Empty;
    public CorpusApplication AuthoringApplication { get; init; } = new();
    public IReadOnlyList<string> OperationSequence { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> ExpectedText { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> ExpectedStructure { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> ProtectedPackageParts { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> ExpectedRefusalCodes { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> AllowedTransformations { get; init; } = Array.Empty<string>();
}

internal sealed class CorpusOrigin
{
    public string Kind { get; init; } = string.Empty;
    public string Recipe { get; init; } = string.Empty;
    public string License { get; init; } = string.Empty;
}

internal sealed class CorpusApplication
{
    public string Name { get; init; } = string.Empty;
    public string Version { get; init; } = string.Empty;
}
