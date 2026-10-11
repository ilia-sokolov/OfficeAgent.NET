using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// The Word compatibility mode is read and set as the <c>docProperty</c> path
/// <c>settings/compatibilityMode</c> (#44). A document without it opens in Word as a
/// Word 2007 document, in Compatibility Mode.
/// </summary>
public class CompatibilityModeTests
{
    private const string Path = "settings/compatibilityMode";

    private static OfficeAgentClient Client() => new(new WordModule());

    [Fact]
    public void Inspection_reports_the_mode_a_new_document_declares()
    {
        var node = ModeNode(new WordModule().CreateBlank());

        Assert.Equal("15", node.Summary);
        Assert.Equal("15", node.Anchor!.Expect);
    }

    [Fact]
    public void Inspection_reports_an_empty_mode_when_the_document_declares_none()
    {
        Assert.Equal(string.Empty, ModeNode(WithoutSettings()).Summary);
    }

    [Fact]
    public void Setting_the_mode_creates_the_settings_part_when_it_is_missing()
    {
        var result = Apply(WithoutSettings(), "15");

        AssertValid(result);
        Assert.Equal("15", ModeNode(result).Summary);
        using var document = Open(result);
        var setting = document.MainDocumentPart!.DocumentSettingsPart!.Settings!
            .GetFirstChild<Compatibility>()!.Elements<CompatibilitySetting>().Single();
        Assert.Equal(CompatSettingNameValues.CompatibilityMode, setting.Name!.Value);
        Assert.Equal("http://schemas.microsoft.com/office/word", setting.Uri!.Value);
    }

    [Fact]
    public void Setting_the_mode_replaces_the_existing_value_and_keeps_other_settings()
    {
        var result = Apply(new WordModule().CreateBlank(), "14");

        AssertValid(result);
        Assert.Equal("14", ModeNode(result).Summary);
        using var document = Open(result);
        var compatibility = document.MainDocumentPart!.DocumentSettingsPart!.Settings!.GetFirstChild<Compatibility>()!;
        Assert.Single(compatibility.Elements<CompatibilitySetting>(), s => s.Name!.Value == CompatSettingNameValues.CompatibilityMode);
    }

    [Fact]
    public void Preview_shows_the_mode_before_and_after()
    {
        var change = Assert.Single(Client().Preview(Handle(new WordModule().CreateBlank()), Plan("12")).Changes);

        Assert.Equal("15", change.Before);
        Assert.Equal("12", change.After);
    }

    [Theory]
    [InlineData("13")]
    [InlineData("16")]
    [InlineData("")]
    [InlineData("fifteen")]
    public void A_value_word_does_not_define_is_refused(string value)
    {
        var report = Client().Preview(Handle(new WordModule().CreateBlank()), Plan(value));

        Assert.False(report.IsValid);
        Assert.Equal(ValidationErrorCodes.InvalidOperation, Assert.Single(report.Errors).Code);
    }

    private static DocumentPlan Plan(string value) => new()
    {
        Operations = new PlanOperation[]
        {
            new SetPropertyOp { Target = new NodeAnchor { Kind = "docProperty", Path = Path }, Value = value }
        }
    };

    private static byte[] Apply(byte[] document, string value)
    {
        using var applied = Client().Commit(Handle(document), Plan(value));
        Assert.True(applied.Committed, string.Join("; ", applied.Report.Errors.Select(e => $"{e.Code}: {e.Message}")));
        return applied.ToBytes();
    }

    private static NodeInfo ModeNode(byte[] document) =>
        Client().Inspect(document).Nodes.Single(n => n.Kind == "docProperty" && n.Path == Path);

    private static byte[] WithoutSettings()
    {
        var buffer = new MemoryStream();
        using (var package = WordprocessingDocument.Create(buffer, WordprocessingDocumentType.Document))
        {
            package.AddMainDocumentPart().Document =
                new Document(new Body(new Paragraph(new Run(new Text("Body text.")))));
        }
        return buffer.ToArray();
    }

    private static StreamHandle Handle(byte[] bytes) => new(new MemoryStream(bytes));

    private static WordprocessingDocument Open(byte[] bytes) =>
        WordprocessingDocument.Open(new MemoryStream(bytes), isEditable: false);

    private static void AssertValid(byte[] document)
    {
        using var opened = Open(document);
        var problems = new OpenXmlValidator(FileFormatVersions.Office2019)
            .Validate(opened)
            .Select(e => $"{e.Path?.XPath}: {e.Description}")
            .ToList();
        Assert.True(problems.Count == 0, string.Join("; ", problems.Take(3)));
    }
}
