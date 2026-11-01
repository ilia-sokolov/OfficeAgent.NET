using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeAgent.Abstractions;

namespace OfficeAgent.Core;

/// <summary>
/// Reads Microsoft Purview sensitivity labels and lists them as <c>sensitivityLabel</c> nodes,
/// so a host can decide what an agent may do with a labelled document.
/// </summary>
/// <remarks>
/// Office stores a label in one of two places or both: <c>MSIP_Label_*</c> custom properties,
/// or the <c>docMetadata/LabelInfo.xml</c> part that current versions write. The name appears
/// only in the properties. Nothing here writes: no plan verb changes or removes a label, and
/// because an edit leaves parts it does not touch alone, a label survives every save.
/// </remarks>
internal static class SensitivityLabels
{
    /// <summary>The node kind labels are listed under.</summary>
    public const string Kind = "sensitivityLabel";

    private const string LabelInfoContentType = "application/vnd.ms-office.classificationlabels+xml";

    private static readonly Regex PropertyName = new(
        "^MSIP_Label_(?<id>[0-9A-Fa-f-]{36})_(?<field>[A-Za-z]+)$",
        RegexOptions.CultureInvariant,
        TimeSpan.FromSeconds(1));

    private sealed class Label
    {
        public string? Name { get; set; }
        public string? Method { get; set; }
        public bool Active { get; set; } = true;
    }

    /// <summary>One node per active label; a document with no label has no node.</summary>
    public static IEnumerable<NodeInfo> Nodes(OpenXmlPackage package)
    {
        var labels = new SortedDictionary<string, Label>(StringComparer.Ordinal);
        ReadProperties(package, labels);
        ReadLabelInfo(package, labels);

        foreach (var pair in labels)
        {
            if (!pair.Value.Active) continue;
            yield return new NodeInfo
            {
                Kind = Kind,
                Path = $"{Kind}#{pair.Key}",
                Summary = Describe(pair.Key, pair.Value),
                Name = pair.Value.Name
            };
        }
    }

    /// <summary>
    /// A person chooses a label with the <c>Privileged</c> method; <c>Standard</c> is one applied
    /// by default or automatically.
    /// </summary>
    private static string Describe(string id, Label label)
    {
        var name = string.IsNullOrEmpty(label.Name) ? "unnamed label" : label.Name;
        var applied = label.Method switch
        {
            null or "" => "applied by an unknown method",
            var method when string.Equals(method, "Privileged", StringComparison.OrdinalIgnoreCase) =>
                "applied by a person",
            var method when string.Equals(method, "Standard", StringComparison.OrdinalIgnoreCase) =>
                "applied by default",
            var method => $"applied by method '{method}'"
        };
        return $"{name}, id {id}, {applied}";
    }

    /// <remarks>
    /// Read from the raw part, not <c>part.Properties</c>: loading the SDK's object model marks
    /// the part to be written back on save, which would change its bytes on every edit.
    /// </remarks>
    private static void ReadProperties(OpenXmlPackage package, IDictionary<string, Label> labels)
    {
        var part = CustomProperties(package);
        if (part is null) return;

        foreach (var property in Parse(part).Descendants().Where(e => e.Name.LocalName == "property"))
        {
            var match = PropertyName.Match(property.Attribute("name")?.Value ?? string.Empty);
            if (!match.Success) continue;

            var label = Find(labels, match.Groups["id"].Value);
            var value = property.Value;
            switch (match.Groups["field"].Value)
            {
                case "Name": label.Name = value; break;
                case "Method": label.Method = value; break;
                case "Enabled":
                    if (string.Equals(value, "false", StringComparison.OrdinalIgnoreCase)) label.Active = false;
                    break;
            }
        }
    }

    private static CustomFilePropertiesPart? CustomProperties(OpenXmlPackage package) => package switch
    {
        WordprocessingDocument word => word.CustomFilePropertiesPart,
        SpreadsheetDocument workbook => workbook.CustomFilePropertiesPart,
        PresentationDocument deck => deck.CustomFilePropertiesPart,
        _ => null
    };

    private static void ReadLabelInfo(OpenXmlPackage package, IDictionary<string, Label> labels)
    {
        foreach (var pair in package.Parts)
        {
            if (!string.Equals(pair.OpenXmlPart.ContentType, LabelInfoContentType, StringComparison.OrdinalIgnoreCase))
                continue;

            foreach (var element in Parse(pair.OpenXmlPart).Descendants().Where(e => e.Name.LocalName == "label"))
            {
                if (element.Attribute("id")?.Value is not { Length: > 0 } id) continue;

                var label = Find(labels, id);
                if (element.Attribute("method")?.Value is { Length: > 0 } method) label.Method = method;
                if (element.Attribute("enabled")?.Value == "0" || element.Attribute("removed")?.Value == "1")
                    label.Active = false;
            }
        }
    }

    private static XDocument Parse(OpenXmlPart part)
    {
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        using var reader = XmlReader.Create(stream, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null
        });
        return XDocument.Load(reader);
    }

    /// <summary>The entry for a label, whichever store named it and however it wrote the id.</summary>
    private static Label Find(IDictionary<string, Label> labels, string id)
    {
        var key = id.Trim('{', '}').ToLowerInvariant();
        if (!labels.TryGetValue(key, out var label))
        {
            label = new Label();
            labels[key] = label;
        }

        return label;
    }
}
