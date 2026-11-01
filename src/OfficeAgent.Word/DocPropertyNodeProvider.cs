using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;

namespace OfficeAgent.Word;

/// <summary>
/// Surfaces and resolves OPC core document properties (creator, title, subject, …) and the
/// document's compatibility mode (<c>settings/compatibilityMode</c>). Fully deterministic:
/// these are plain metadata, no layout involved.
/// </summary>
internal sealed class DocPropertyNodeProvider : IWordNodeProvider
{
    public string Kind => "docProperty";

    /// <summary>The path of the compatibility-mode setting.</summary>
    public const string CompatibilityModePath = "settings/compatibilityMode";

    /// <summary>
    /// The compatibility modes Word defines: 11 is Word 2003, 12 Word 2007, 14 Word 2010, and
    /// 15 Word 2013 and later. Word never wrote 13.
    /// </summary>
    public static readonly string[] CompatibilityModes = { "11", "12", "14", "15" };

    private const string WordCompatibilityUri = "http://schemas.microsoft.com/office/word";

    public static bool IsCore(string name) => PackageCoreProperties.IsCore(name);

    public IEnumerable<NodeInfo> Enumerate(WordObjectMap map)
    {
        foreach (var node in PackageCoreProperties.Nodes(map.Doc))
            yield return node;

        foreach (var node in SensitivityLabels.Nodes(map.Doc))
            yield return node;

        var mode = ReadCompatibilityMode(map.Main);
        yield return new NodeInfo
        {
            Kind = Kind,
            Path = CompatibilityModePath,
            Summary = mode ?? string.Empty,
            Anchor = new NodeAnchor { Id = "docprop:compatibilityMode", Kind = Kind, Path = CompatibilityModePath, Expect = mode }
        };
    }

    public ResolvedNode? Resolve(NodeAnchor anchor, WordObjectMap map)
    {
        if (anchor.Path == CompatibilityModePath)
            return new ResolvedNode { Kind = Kind, Value = ReadCompatibilityMode(map.Main) };

        var name = NameOf(anchor.Path);
        return IsCore(name) ? new ResolvedNode { Kind = Kind, Value = Read(map, name) } : null;
    }

    /// <summary>
    /// The value of the <c>compatSetting</c> named <c>compatibilityMode</c>, or
    /// <see langword="null"/> when the document has none - which Word reads as a Word 2007
    /// document and opens in Compatibility Mode.
    /// </summary>
    public static string? ReadCompatibilityMode(MainDocumentPart main) =>
        main.DocumentSettingsPart?.Settings?.GetFirstChild<Compatibility>()?
            .Elements<CompatibilitySetting>()
            .FirstOrDefault(IsCompatibilityMode)?.Val?.Value;

    /// <summary>
    /// Writes or replaces the compatibility mode, creating the settings part and
    /// <c>w:compat</c> when they are missing. Every other setting is left as it is.
    /// </summary>
    public static void WriteCompatibilityMode(MainDocumentPart main, string mode)
    {
        var part = main.DocumentSettingsPart ?? main.AddNewPart<DocumentSettingsPart>();
        part.Settings ??= new Settings();

        var compatibility = part.Settings.GetFirstChild<Compatibility>();
        if (compatibility is null)
        {
            compatibility = new Compatibility();
            // w:compat sits at a fixed place in the CT_Settings sequence; AddChild puts it there.
            part.Settings.AddChild(compatibility);
        }

        var setting = compatibility.Elements<CompatibilitySetting>().FirstOrDefault(IsCompatibilityMode);
        if (setting is null)
        {
            setting = new CompatibilitySetting
            {
                Name = CompatSettingNameValues.CompatibilityMode,
                Uri = WordCompatibilityUri
            };
            compatibility.AddChild(setting);
        }
        setting.Val = mode;
        part.Settings.Save();
    }

    private static bool IsCompatibilityMode(CompatibilitySetting setting) =>
        setting.Name?.Value == CompatSettingNameValues.CompatibilityMode &&
        (setting.Uri?.Value is null || setting.Uri.Value == WordCompatibilityUri);

    public static string NameOf(string path) => PackageCoreProperties.NameOf(path);

    public static string? Read(WordObjectMap map, string name) => PackageCoreProperties.Read(map.Doc, name);

    public static void Write(WordObjectMap map, string name, string? value) =>
        PackageCoreProperties.Write(map.Doc, name, value);
}
