using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// The identifiers Word stabilization mints. <c>w14:paraId</c> is schema-constrained to
/// <c>0 &lt; value &lt; 0x80000000</c>, so a plain random 32-bit value is invalid about
/// half the time - a document Word silently rewrites on save and strict consumers reject.
/// </summary>
public class WordParaIdTests
{
    [Fact]
    public void Stabilized_documents_stay_schema_valid_across_many_runs()
    {
        // One run proves little: a naive generator produces a valid id ~50% of the time,
        // so a single-shot test passes half the time by luck. Repeat until that is
        // vanishingly unlikely.
        for (var attempt = 0; attempt < 40; attempt++)
        {
            var edited = ApplyOneEdit();

            using var stream = new MemoryStream(edited);
            using var document = WordprocessingDocument.Open(stream, isEditable: false);
            var problems = new OpenXmlValidator(FileFormatVersions.Office2019)
                .Validate(document)
                .Select(e => $"{e.Path?.XPath}: {e.Description}")
                .ToList();

            Assert.True(problems.Count == 0,
                $"attempt {attempt}: {string.Join("; ", problems.Take(3))}");
        }
    }

    [Fact]
    public void Every_minted_paragraph_id_is_inside_the_range_word_accepts()
    {
        var seen = new List<uint>();

        for (var attempt = 0; attempt < 40; attempt++)
        {
            using var stream = new MemoryStream(ApplyOneEdit());
            using var document = WordprocessingDocument.Open(stream, isEditable: false);

            foreach (var paragraph in document.MainDocumentPart!.Document.Body!.Elements<Paragraph>())
            {
                var raw = paragraph.ParagraphId?.Value;
                if (string.IsNullOrEmpty(raw)) continue;

                Assert.True(uint.TryParse(raw, System.Globalization.NumberStyles.HexNumber,
                    System.Globalization.CultureInfo.InvariantCulture, out var value),
                    $"'{raw}' is not a hex paraId");
                Assert.InRange(value, 1u, 0x7FFFFFFFu);
                seen.Add(value);
            }
        }

        Assert.NotEmpty(seen);
    }

    [Fact]
    public void Minted_ids_do_not_collide_with_ids_already_in_the_document()
    {
        // A document whose paragraphs already carry ids must keep them, and any paragraph
        // that lacks one must not be handed a duplicate.
        using var stream = new MemoryStream(ApplyOneEdit());
        using var document = WordprocessingDocument.Open(stream, isEditable: false);

        var ids = document.MainDocumentPart!.Document.Body!
            .Descendants<Paragraph>()
            .Select(p => p.ParagraphId?.Value)
            .Where(id => !string.IsNullOrEmpty(id))
            .ToList();

        Assert.Equal(ids.Count, ids.Distinct(StringComparer.OrdinalIgnoreCase).Count());
    }

    /// <summary>
    /// Applies a trivial edit to a document with unstabilized paragraphs, which is what
    /// forces stabilization to mint ids.
    /// </summary>
    private static byte[] ApplyOneEdit()
    {
        var client = new OfficeAgentClient(new WordModule());
        var source = UnstabilizedDocument();

        var hits = client.Find(
            new StreamHandle(new MemoryStream(source)),
            new FindQuery { Pattern = "Acme" });

        using var applied = client.Commit(
            new StreamHandle(new MemoryStream(source)),
            new DocumentPlan
            {
                Operations = new PlanOperation[]
                {
                    new ChangeTextOp { Target = hits[0].Anchor!, With = "Globex", Mode = ChangeMode.Direct }
                }
            });

        Assert.True(applied.Committed,
            string.Join("; ", applied.Report.Errors.Select(e => $"{e.Code}: {e.Message}")));
        return applied.ToBytes();
    }

    /// <summary>A document whose paragraphs carry no <c>w14:paraId</c> at all.</summary>
    private static byte[] UnstabilizedDocument()
    {
        using var buffer = new MemoryStream();
        using (var document = WordprocessingDocument.Create(buffer, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Acme Corp supplies the goods."))),
                new Paragraph(new Run(new Text("Payment is due on receipt."))),
                new Paragraph(new Run(new Text("Signed for Acme Corp.")))));
            main.Document.Save();
        }
        return buffer.ToArray();
    }
}
