using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Word;
using Xunit.Sdk;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using W16Cid = DocumentFormat.OpenXml.Office2019.Word.Cid;

namespace OfficeAgent.Tests;

public sealed class WordPreservationEvidenceTests
{
    private const string TestedBaseCommit = "a687bc15a1d7af87790333984be4721055bdb90e";
    private static readonly DateTimeOffset EvidenceInstant =
        new(2026, 9, 14, 8, 0, 0, TimeSpan.Zero);
    private static readonly TimeProvider EvidenceClock = new FixedTimeProvider(EvidenceInstant);
    private static readonly string[] RequiredCategories =
    {
        "mixed-runs",
        "existing-insertions-deletions",
        "classic-comments",
        "commentsExtended",
        "commentsIds",
        "people",
        "paragraph-identifiers",
        "nested-controls",
        "numbering",
        "fields",
        "sections"
    };
    private static readonly string RepositoryRoot = FindRepositoryRoot();
    private static readonly string CorpusRoot = Path.Combine(
        RepositoryRoot, "tests", "OfficeAgent.Tests", "Corpus", "v0.9.0", "word-preservation");
    private static readonly string[] BodyEditProtectedParts =
    {
        "[Content_Types].xml",
        "_rels/.rels",
        "word/_rels/document.xml.rels",
        "word/comments.xml",
        "word/commentsExtended.xml",
        "word/commentsIds.xml",
        "word/numbering.xml",
        "word/people.xml"
    };
    private static readonly string[] CommentResolveProtectedParts =
    {
        "[Content_Types].xml",
        "_rels/.rels",
        "word/_rels/document.xml.rels",
        "word/comments.xml",
        "word/commentsIds.xml",
        "word/document.xml",
        "word/numbering.xml",
        "word/people.xml"
    };

    [Fact]
    [Trait("Category", "Fixture")]
    public void Generate_word_preservation_fixture_only_when_an_output_root_is_explicit()
    {
        var outputRoot = Environment.GetEnvironmentVariable("OFFICEAGENT_WORD_PRESERVATION_OUTPUT");
        if (string.IsNullOrWhiteSpace(outputRoot))
            return;

        Directory.CreateDirectory(outputRoot);
        var fixture = WordPreservationFixtureBuilder.Build();
        File.WriteAllBytes(Path.Combine(outputRoot, "rich-word-features.docx"), fixture);
        using (var bodyEdit = ApplyBodyEdit(fixture))
            File.WriteAllBytes(Path.Combine(outputRoot, "tracked-mixed-run-replacement.docx"), bodyEdit.ToBytes());
        using (var commentResolve = ApplyCommentResolve(fixture))
            File.WriteAllBytes(Path.Combine(outputRoot, "resolved-comment.docx"), commentResolve.ToBytes());
        File.WriteAllText(
            Path.Combine(outputRoot, "report.json"),
            GenerateReportJson(),
            new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
    }

    [Fact]
    public void Generated_fixture_is_valid_and_contains_every_inventory_category()
    {
        var generated = WordPreservationFixtureBuilder.Build();
        AssertValid(generated);
        AssertFixtureSemantics(generated);

        var committedPath = Path.Combine(CorpusRoot, "rich-word-features.docx");
        Assert.True(File.Exists(committedPath), $"Missing preservation fixture: {committedPath}");
        var committed = File.ReadAllBytes(committedPath);
        Assert.Equal(Sha256(generated), Sha256(committed));

        using var bodyEdit = ApplyBodyEdit(generated);
        Assert.Equal(
            Sha256(bodyEdit.ToBytes()),
            Sha256(File.ReadAllBytes(Path.Combine(CorpusRoot, "tracked-mixed-run-replacement.docx"))));
        using var commentResolve = ApplyCommentResolve(generated);
        Assert.Equal(
            Sha256(commentResolve.ToBytes()),
            Sha256(File.ReadAllBytes(Path.Combine(CorpusRoot, "resolved-comment.docx"))));
    }

    [Fact]
    public void Tracked_mixed_run_edit_changes_only_document_xml_and_preserves_rich_features()
    {
        var input = Fixture();
        using var result = ApplyBodyEdit(input);
        Assert.True(result.Committed, Errors(result.Report));
        var output = result.ToBytes();

        WordPreservationVerifier.Verify(
            input,
            output,
            new[] { "word/document.xml" },
            BodyEditProtectedParts,
            bytes => AssertBodyEditSemantics(input, bytes));
    }

    [Fact]
    public void Resolve_comment_changes_only_comments_extended_and_keeps_comment_identity()
    {
        var input = Fixture();
        using var result = ApplyCommentResolve(input);
        Assert.True(result.Committed, Errors(result.Report));
        var output = result.ToBytes();

        WordPreservationVerifier.Verify(
            input,
            output,
            new[] { "word/commentsExtended.xml" },
            CommentResolveProtectedParts,
            AssertResolvedCommentSemantics);
    }

    [Fact]
    public void Verifier_rejects_protected_part_relationship_revision_and_comment_identity_corruption()
    {
        var input = Fixture();
        using var bodyResult = ApplyBodyEdit(input);
        var bodyOutput = bodyResult.ToBytes();

        var damagedComment = WordPreservationVerifier.ReplacePart(
            bodyOutput,
            "word/comments.xml",
            bytes => Encoding.UTF8.GetBytes(Encoding.UTF8.GetString(bytes).Replace(
                "Preserve this review thread.", "Damaged review thread.", StringComparison.Ordinal)));
        Assert.Throws<PreservationVerificationException>(() => WordPreservationVerifier.Verify(
            input, damagedComment, new[] { "word/document.xml" }, BodyEditProtectedParts,
            bytes => AssertBodyEditSemantics(input, bytes)));

        var damagedRelationship = WordPreservationVerifier.ReplacePart(
            bodyOutput,
            "word/_rels/document.xml.rels",
            bytes => Encoding.UTF8.GetBytes(Encoding.UTF8.GetString(bytes).Replace(
                "numbering.xml", "missing-numbering.xml", StringComparison.Ordinal)));
        Assert.Throws<PreservationVerificationException>(() => WordPreservationVerifier.Verify(
            input, damagedRelationship, new[] { "word/document.xml" }, BodyEditProtectedParts,
            bytes => AssertBodyEditSemantics(input, bytes)));

        var damagedRevision = WordPreservationVerifier.ReplacePart(
            bodyOutput,
            "word/document.xml",
            bytes => Encoding.UTF8.GetBytes(Encoding.UTF8.GetString(bytes).Replace(
                "Prior Author", "Wrong Author", StringComparison.Ordinal)));
        Assert.Throws<PreservationVerificationException>(() => WordPreservationVerifier.Verify(
            input, damagedRevision, new[] { "word/document.xml" }, BodyEditProtectedParts,
            bytes => AssertBodyEditSemantics(input, bytes)));

        var damagedIdentity = WordPreservationVerifier.ReplacePart(
            input,
            "word/commentsIds.xml",
            bytes => Encoding.UTF8.GetBytes(Encoding.UTF8.GetString(bytes).Replace(
                "20000001", "20000002", StringComparison.Ordinal)));
        Assert.Throws<PreservationVerificationException>(() => WordPreservationVerifier.Verify(
            input,
            damagedIdentity,
            new[] { "word/commentsIds.xml" },
            Array.Empty<string>(),
            AssertCommentIdentity));
    }

    [Fact]
    public async Task Unsupported_word_mutations_refuse_without_direct_or_provider_state_changes()
    {
        var input = Fixture();
        var operations = new PlanOperation[]
        {
            new SetCellOp
            {
                Target = new CellAnchor { SheetId = 1U, Address = "A1" },
                Value = "not a Word cell"
            },
            new InsertSlideOp { Slide = new SlideData { Title = "not a Word slide" } }
        };

        foreach (var operation in operations)
        {
            var client = Client();
            CorpusAssertions.AssertRejectedWithoutMutation(
                input,
                bytes => client.Commit(
                    new StreamHandle(new MemoryStream(bytes, writable: false), "rich-word-features.docx"),
                    new DocumentPlan { Operations = new[] { operation } }),
                ValidationErrorCodes.UnsupportedOperation);

            var provider = new MemoryDocumentProvider("evidence");
            var reference = provider.Add("rich-word-features.docx", input);
            var providerClient = new OfficeAgentClient(
                new DocumentProviderRegistry(new[] { provider }),
                new WordModule(EvidenceClock));
            var beforeVersion = reference.Version;
            var beforeHash = Sha256(provider.Read(reference.ItemId));

            var result = await providerClient.CommitAsync(
                reference,
                new DocumentPlan { Operations = new[] { operation } });

            Assert.False(result.Committed);
            Assert.Contains(result.Report.Errors, error => error.Code == ValidationErrorCodes.UnsupportedOperation);
            Assert.Null(result.Document);
            Assert.Equal(1, provider.Count);
            Assert.Equal(beforeVersion, provider.Describe(reference.ItemId).Version);
            Assert.Equal(beforeHash, Sha256(provider.Read(reference.ItemId)));
        }
    }

    [Fact]
    public void Committed_manifest_and_report_match_fixed_inputs_and_fresh_evidence()
    {
        var manifestPath = Path.Combine(CorpusRoot, "manifest.json");
        Assert.True(File.Exists(manifestPath), $"Missing preservation manifest: {manifestPath}");
        using var manifest = JsonDocument.Parse(File.ReadAllBytes(manifestPath));
        var root = manifest.RootElement;
        Assert.Equal("1.0", root.GetProperty("schemaVersion").GetString());
        Assert.Equal(TestedBaseCommit, root.GetProperty("testedBaseCommit").GetString());
        Assert.Equal(Sha256(Fixture()), root.GetProperty("fixture").GetProperty("sha256").GetString());
        foreach (var artifact in root.GetProperty("artifacts").EnumerateArray()
            .Where(item => item.TryGetProperty("sha256", out _)))
        {
            var path = Path.GetFullPath(Path.Combine(
                RepositoryRoot,
                artifact.GetProperty("path").GetString()!));
            Assert.StartsWith(
                RepositoryRoot + Path.DirectorySeparatorChar,
                path,
                StringComparison.OrdinalIgnoreCase);
            Assert.Equal(artifact.GetProperty("sha256").GetString(), Sha256(File.ReadAllBytes(path)));
        }

        var categories = root.GetProperty("cases").EnumerateArray()
            .SelectMany(item => item.GetProperty("categories").EnumerateArray())
            .Select(item => item.GetString())
            .ToHashSet(StringComparer.Ordinal);
        Assert.Subset(new HashSet<string>(RequiredCategories, StringComparer.Ordinal), categories);
        Assert.All(root.GetProperty("cases").EnumerateArray(), item =>
        {
            Assert.Contains(item.GetProperty("disposition").GetString(),
                new[] { "supported-edit", "preservation-only", "refusal" });
            Assert.NotEmpty(item.GetProperty("operation").GetString() ?? string.Empty);
            Assert.True(item.GetProperty("allowedChangedParts").ValueKind == JsonValueKind.Array);
            Assert.True(item.GetProperty("protectedParts").ValueKind == JsonValueKind.Array);
        });

        var committedReport = File.ReadAllText(Path.Combine(CorpusRoot, "report.json"));
        Assert.Equal(NormalizeNewlines(GenerateReportJson()), NormalizeNewlines(committedReport));
    }

    private static ApplyResult ApplyBodyEdit(byte[] input) => Client().Commit(
        new StreamHandle(new MemoryStream(input, writable: false), "rich-word-features.docx"),
        new DocumentPlan
        {
            Revision = new RevisionMetadata
            {
                Author = "Preservation Agent",
                TimestampUtc = EvidenceInstant
            },
            Operations = new PlanOperation[]
            {
                new ChangeTextOp
                {
                    Target = new TextSpanAnchor
                    {
                        ParaId = "w14:10000001",
                        Expect = "target"
                    },
                    With = "objective",
                    Mode = ChangeMode.Tracked
                }
            }
        });

    private static ApplyResult ApplyCommentResolve(byte[] input) => Client().Commit(
        new StreamHandle(new MemoryStream(input, writable: false), "rich-word-features.docx"),
        new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new CommentOp
                {
                    Target = new NodeAnchor { Kind = "comment", Path = "comment#1" },
                    Action = CommentAction.Resolve
                }
            }
        });

    private static OfficeAgentClient Client() =>
        new(new WordModule(EvidenceClock));

    private static byte[] Fixture() =>
        File.ReadAllBytes(Path.Combine(CorpusRoot, "rich-word-features.docx"));

    private static void AssertBodyEditSemantics(byte[] before, byte[] after)
    {
        AssertValid(after);
        using var beforeDocument = WordprocessingDocument.Open(new MemoryStream(before), false);
        using var afterDocument = WordprocessingDocument.Open(new MemoryStream(after), false);
        var beforeMain = beforeDocument.MainDocumentPart!;
        var afterMain = afterDocument.MainDocumentPart!;
        var beforeBody = beforeMain.Document.Body!;
        var afterBody = afterMain.Document.Body!;

        var beforeIds = beforeBody.Descendants<Paragraph>()
            .Select(paragraph => paragraph.ParagraphId?.Value)
            .Where(id => id is not null)
            .OrderBy(id => id, StringComparer.Ordinal)
            .ToArray();
        var afterIds = afterBody.Descendants<Paragraph>()
            .Select(paragraph => paragraph.ParagraphId?.Value)
            .Where(id => id is not null)
            .OrderBy(id => id, StringComparer.Ordinal)
            .ToArray();
        Assert.Equal(beforeIds, afterIds);

        foreach (var paraId in new[] { "10000002", "10000003", "10000005", "10000006", "10000007" })
        {
            var beforeParagraph = beforeBody.Descendants<Paragraph>().Single(p => p.ParagraphId?.Value == paraId);
            var afterParagraph = afterBody.Descendants<Paragraph>().Single(p => p.ParagraphId?.Value == paraId);
            Assert.Equal(beforeParagraph.OuterXml, afterParagraph.OuterXml);
        }

        Assert.Equal(
            beforeBody.Descendants<SdtBlock>().Single().OuterXml,
            afterBody.Descendants<SdtBlock>().Single().OuterXml);

        var oldInsertion = afterBody.Descendants<InsertedRun>().Single(run => run.Id?.Value == "100");
        var oldDeletion = afterBody.Descendants<DeletedRun>().Single(run => run.Id?.Value == "101");
        Assert.Equal("Prior Author", oldInsertion.Author?.Value);
        Assert.Equal("Prior Author", oldDeletion.Author?.Value);
        Assert.Equal(new DateTime(2026, 8, 1, 9, 0, 0, DateTimeKind.Utc), oldInsertion.Date?.Value);
        Assert.Equal(new DateTime(2026, 8, 1, 9, 1, 0, DateTimeKind.Utc), oldDeletion.Date?.Value);
        Assert.Equal("approved", oldInsertion.InnerText);
        Assert.Equal("draft", string.Concat(oldDeletion.Descendants<DeletedText>().Select(text => text.Text)));

        var newInsertions = afterBody.Descendants<InsertedRun>()
            .Where(run => run.Author?.Value == "Preservation Agent")
            .ToArray();
        var newDeletions = afterBody.Descendants<DeletedRun>()
            .Where(run => run.Author?.Value == "Preservation Agent")
            .ToArray();
        Assert.NotEmpty(newInsertions);
        Assert.NotEmpty(newDeletions);
        Assert.Equal("objective", string.Concat(newInsertions.Select(run => run.InnerText)));
        Assert.Equal("target", string.Concat(newDeletions.SelectMany(run =>
            run.Descendants<DeletedText>()).Select(text => text.Text)));
        Assert.All(newInsertions, run => Assert.Equal(EvidenceInstant.UtcDateTime, run.Date?.Value));
        Assert.All(newDeletions, run => Assert.Equal(EvidenceInstant.UtcDateTime, run.Date?.Value));

        AssertCommentIdentity(after, expectedResolved: false);
    }

    private static void AssertResolvedCommentSemantics(byte[] bytes)
    {
        AssertValid(bytes);
        AssertCommentIdentity(bytes, expectedResolved: true);
    }

    private static void AssertCommentIdentity(byte[] bytes) =>
        AssertCommentIdentity(bytes, expectedResolved: false);

    private static void AssertCommentIdentity(byte[] bytes, bool expectedResolved)
    {
        using var document = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        var main = document.MainDocumentPart!;
        var comment = main.WordprocessingCommentsPart!.Comments!.Elements<Comment>().Single();
        var paraId = comment.Elements<Paragraph>().Single().ParagraphId?.Value;
        Assert.Equal("1", comment.Id?.Value);
        Assert.Equal("Reviewer One", comment.Author?.Value);
        Assert.Equal("20000001", paraId);

        var commentEx = main.WordprocessingCommentsExPart!.CommentsEx!
            .Elements<W15.CommentEx>().Single();
        Assert.Equal(paraId, commentEx.ParaId?.Value);
        Assert.Equal(expectedResolved, commentEx.Done?.Value ?? false);

        var commentId = main.WordprocessingCommentsIdsPart!.CommentsIds!
            .Elements<W16Cid.CommentId>().Single();
        Assert.Equal(paraId, commentId.ParaId?.Value);
        Assert.Equal("D0000001", commentId.DurableId?.Value);

        var person = main.WordprocessingPeoplePart!.People!.Elements<W15.Person>().Single();
        Assert.Equal(comment.Author?.Value, person.Author?.Value);
    }

    private static string GenerateReportJson()
    {
        var input = Fixture();
        AssertValid(input);
        AssertFixtureSemantics(input);

        using var bodyResult = ApplyBodyEdit(input);
        Assert.True(bodyResult.Committed, Errors(bodyResult.Report));
        var bodyOutput = bodyResult.ToBytes();
        WordPreservationVerifier.Verify(
            input, bodyOutput, new[] { "word/document.xml" }, BodyEditProtectedParts,
            bytes => AssertBodyEditSemantics(input, bytes));

        using var resolveResult = ApplyCommentResolve(input);
        Assert.True(resolveResult.Committed, Errors(resolveResult.Report));
        var resolveOutput = resolveResult.ToBytes();
        WordPreservationVerifier.Verify(
            input, resolveOutput, new[] { "word/commentsExtended.xml" },
            CommentResolveProtectedParts, AssertResolvedCommentSemantics);

        var refusalOperations = new (string Name, PlanOperation Operation)[]
        {
            ("setCell-on-word", new SetCellOp
            {
                Target = new CellAnchor { SheetId = 1U, Address = "A1" },
                Value = "not a Word cell"
            }),
            ("insertSlide-on-word", new InsertSlideOp
            {
                Slide = new SlideData { Title = "not a Word slide" }
            })
        };
        var refusals = refusalOperations.Select(item =>
        {
            using var result = Client().Commit(
                new StreamHandle(new MemoryStream(input, writable: false), "rich-word-features.docx"),
                new DocumentPlan { Operations = new[] { item.Operation } });
            Assert.False(result.Committed);
            var code = Assert.Single(result.Report.Errors).Code;
            Assert.Equal(ValidationErrorCodes.UnsupportedOperation, code);
            Assert.Null(result.Output);
            return new
            {
                id = item.Name,
                status = "PASSED",
                expectedCode = ValidationErrorCodes.UnsupportedOperation,
                observedCode = code,
                inputSha256After = Sha256(input),
                outputCreated = false
            };
        }).ToArray();

        var report = new
        {
            schemaVersion = "1.0",
            evidenceDate = "2026-09-14",
            testedBaseCommit = TestedBaseCommit,
            candidate = "v0.9.0 working tree based on testedBaseCommit",
            fixture = new
            {
                path = "tests/OfficeAgent.Tests/Corpus/v0.9.0/word-preservation/rich-word-features.docx",
                origin = "deterministic generated fictional fixture",
                license = "MIT repository fixture",
                sha256 = Sha256(input)
            },
            structuralValidation = new
            {
                status = "PASSED",
                validator = "DocumentFormat.OpenXml OpenXmlValidator",
                fileFormatVersion = "Office2019",
                packagesChecked = 3
            },
            supportedOperations = new object[]
            {
                new
                {
                    id = "tracked-mixed-run-replacement",
                    status = "PASSED",
                    changedParts = WordPreservationVerifier.ChangedParts(input, bodyOutput),
                    protectedParts = BodyEditProtectedParts,
                    semanticAssertions = new[]
                    {
                        "replacement spans styled runs and is tracked",
                        "prior insertion and deletion identity remains",
                        "classic and extended comment identity remains",
                        "paragraph identifiers, nested controls, numbering, field and sections remain"
                    }
                },
                new
                {
                    id = "resolve-classic-comment",
                    status = "PASSED",
                    changedParts = WordPreservationVerifier.ChangedParts(input, resolveOutput),
                    protectedParts = CommentResolveProtectedParts,
                    semanticAssertions = new[]
                    {
                        "commentsExtended done flag becomes true",
                        "classic comment, commentsIds durable identity and person remain"
                    }
                }
            },
            refusals,
            corruptionControls = new[]
            {
                new { id = "protected-comment-part", status = "PASSED", expected = "verifier failure" },
                new { id = "protected-relationship", status = "PASSED", expected = "verifier failure" },
                new { id = "revision-identity", status = "PASSED", expected = "verifier failure" },
                new { id = "comment-identity", status = "PASSED", expected = "verifier failure" }
            },
            nativeOffice = new
            {
                status = "PARTIAL",
                application = "Microsoft Word",
                version = "16.0.20326.20144",
                observationDate = "2026-09-14",
                source = "Manual native observation. The automated generator preserves this recorded result but does not rerun Word.",
                checks = new object[]
                {
                    new
                    {
                        id = "open-tracked-mixed-run-result",
                        status = "PASSED",
                        observation = "Opened in Compatibility Mode with tracked replacement, prior revisions, comment, controls, numbering, field and two pages visible."
                    },
                    new
                    {
                        id = "accept-revisions",
                        status = "NOT_RUN",
                        observation = "No accept operation was performed."
                    },
                    new
                    {
                        id = "reject-revisions",
                        status = "NOT_RUN",
                        observation = "No reject operation was performed."
                    }
                },
                reason = "Opening was observed separately from automated validation. Accept and reject remain assigned to the V10-03 native matrix."
            },
            limitations = new[]
            {
                "Generated fixtures are engineering evidence, not customer documents.",
                "Passing schema, byte-preservation and semantic checks does not prove visual or universal fidelity.",
                "The report covers only the declared operations and package features."
            }
        };

        return JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true }) + "\n";
    }

    private static string Errors(ChangeReport report) =>
        string.Join("; ", report.Errors.Select(error => $"{error.Code}: {error.Message}"));

    private static string NormalizeNewlines(string value) =>
        value.Replace("\r\n", "\n", StringComparison.Ordinal);

    internal static void AssertFixtureSemantics(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var document = WordprocessingDocument.Open(stream, isEditable: false);
        var main = document.MainDocumentPart ?? throw new XunitException("Missing main document part.");
        var body = main.Document.Body ?? throw new XunitException("Missing Word body.");

        var mixed = body.Elements<Paragraph>().Single(p => p.ParagraphId?.Value == "10000001");
        Assert.Equal("Quarterly target remains within plan.", mixed.InnerText);
        Assert.Equal(4, mixed.Elements<Run>().Count());

        var priorInsertion = body.Descendants<InsertedRun>().Single(r => r.Id?.Value == "100");
        var priorDeletion = body.Descendants<DeletedRun>().Single(r => r.Id?.Value == "101");
        Assert.Equal("Prior Author", priorInsertion.Author?.Value);
        Assert.Equal("Prior Author", priorDeletion.Author?.Value);
        Assert.Equal("approved", priorInsertion.InnerText);
        Assert.Equal("draft", string.Concat(priorDeletion.Descendants<DeletedText>().Select(t => t.Text)));

        var outer = body.Descendants<Tag>().Single(t => t.Val?.Value == "OuterClause");
        var inner = body.Descendants<Tag>().Single(t => t.Val?.Value == "InnerParty");
        Assert.NotNull(outer.Ancestors<SdtBlock>().SingleOrDefault());
        Assert.NotNull(inner.Ancestors<SdtRun>().SingleOrDefault());

        var numbered = body.Elements<Paragraph>().Single(p => p.ParagraphId?.Value == "10000005");
        Assert.Equal(42, numbered.ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value);
        Assert.NotNull(main.NumberingDefinitionsPart?.Numbering?.Elements<NumberingInstance>()
            .SingleOrDefault(instance => instance.NumberID?.Value == 42));

        var field = body.Descendants<SimpleField>().Single();
        Assert.Equal("PAGE", field.Instruction?.Value);
        Assert.Equal("7", field.InnerText);
        Assert.Equal(2, body.Descendants<SectionProperties>().Count());

        var comment = main.WordprocessingCommentsPart?.Comments?.Elements<Comment>().Single()
            ?? throw new XunitException("Missing classic comment.");
        var commentParaId = comment.Elements<Paragraph>().Single().ParagraphId?.Value;
        Assert.Equal("20000001", commentParaId);
        Assert.Equal("Reviewer One", comment.Author?.Value);

        var commentEx = main.WordprocessingCommentsExPart?.CommentsEx?.Elements<W15.CommentEx>().Single()
            ?? throw new XunitException("Missing commentsExtended identity.");
        Assert.Equal(commentParaId, commentEx.ParaId?.Value);
        Assert.False(commentEx.Done?.Value ?? false);

        var commentId = main.WordprocessingCommentsIdsPart?.CommentsIds?.Elements<W16Cid.CommentId>().Single()
            ?? throw new XunitException("Missing commentsIds identity.");
        Assert.Equal(commentParaId, commentId.ParaId?.Value);
        Assert.Equal("D0000001", commentId.DurableId?.Value);

        var person = main.WordprocessingPeoplePart?.People?.Elements<W15.Person>().Single()
            ?? throw new XunitException("Missing people metadata.");
        Assert.Equal(comment.Author?.Value, person.Author?.Value);

        Assert.Equal("1", body.Descendants<CommentRangeStart>().Single().Id?.Value);
        Assert.Equal("1", body.Descendants<CommentRangeEnd>().Single().Id?.Value);
        Assert.Equal("1", body.Descendants<CommentReference>().Single().Id?.Value);
    }

    private static void AssertValid(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var document = WordprocessingDocument.Open(stream, isEditable: false);
        var errors = new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019)
            .Validate(document)
            .Select(error => $"{error.Path?.XPath}: {error.Description}")
            .ToArray();
        Assert.True(errors.Length == 0, string.Join("; ", errors));
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
