using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;

namespace OfficeAgent.Word;

/// <summary>Imports supported OPC parts directly; output never depends on deferred altChunk conversion.</summary>
internal static class WordDocumentAssembler
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace Rel = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly XNamespace Ct = "http://schemas.openxmlformats.org/package/2006/content-types";
    private static readonly XNamespace Wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    private static readonly XNamespace Pic = "http://schemas.openxmlformats.org/drawingml/2006/picture";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const string Main = "word/document.xml";
    private const string Styles = "word/styles.xml";
    private const string Numbering = "word/numbering.xml";
    private const string Settings = "word/settings.xml";
    private const string Fonts = "word/fontTable.xml";
    private const string Theme = "word/theme/theme1.xml";
    private static readonly HashSet<string> Globals = new(StringComparer.Ordinal)
        { Main, Styles, Numbering, Settings, Fonts, Theme };
    private static readonly HashSet<string> AllowedRelationships = new(StringComparer.Ordinal)
    {
        "officeDocument", "styles", "numbering", "settings", "fontTable", "theme", "header", "footer",
        "image", "hyperlink", "core-properties", "extended-properties", "custom-properties"
    };
    private static readonly HashSet<string> AllowedTypes = new(StringComparer.Ordinal)
    {
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml",
        "application/vnd.openxmlformats-officedocument.theme+xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
        "application/vnd.openxmlformats-package.core-properties+xml",
        "application/vnd.openxmlformats-officedocument.extended-properties+xml",
        "application/vnd.openxmlformats-officedocument.custom-properties+xml",
        "application/vnd.openxmlformats-package.relationships+xml",
        "image/png", "image/jpeg", "image/gif", "image/bmp", "image/tiff", "image/x-emf", "image/x-wmf"
    };
    // Unknown Word vocabulary in content or styles must not silently lose semantics during import.
    private static readonly HashSet<string> WordElements = new((
        "document body p pPr r rPr t tab br cr noBreakHyphen softHyphen lastRenderedPageBreak " +
        "pStyle rStyle b bCs i iCs caps smallCaps strike dstrike outline shadow emboss imprint vanish webHidden " +
        "color spacing w kern position sz szCs highlight u effect bdr shd fitText vertAlign rtl cs em lang eastAsianLayout specVanish oMath rFonts " +
        "keepNext keepLines pageBreakBefore widowControl outlineLvl numPr ilvl numId suppressLineNumbers pBdr top left bottom right between bar " +
        "tabs contextualSpacing mirrorIndents suppressOverlap jc textDirection textAlignment textboxTightWrap ind " +
        "snapToGrid suppressAutoHyphens kinsoku wordWrap overflowPunct autoSpaceDE autoSpaceDN adjustRightInd " +
        "tbl tblPr tblGrid gridCol tr trPr tc tcPr tblStyle tblpPr tblOverlap bidiVisual tblStyleRowBandSize tblStyleColBandSize " +
        "tblW tblInd tblBorders tblLayout tblCellMar tblLook tblCaption tblDescription start end insideH insideV " +
        "gridBefore gridAfter wBefore wAfter cantSplit trHeight tblHeader tblCellSpacing hidden cnfStyle " +
        "tcW gridSpan hMerge vMerge tcBorders noWrap tcMar tcFitText vAlign hideMark " +
        "sectPr headerReference footerReference type pgSz pgMar paperSrc pgBorders pgNumType cols col formProt vAlign " +
        "noEndnote titlePg textDirection bidi rtlGutter docGrid printerSettings " +
        "hyperlink bookmarkStart bookmarkEnd drawing fldSimple fldChar instrText " +
        "sdt sdtPr sdtContent sdtEndPr alias tag id text lock showingPlcHdr placeholder docPart temporary " +
        "styles docDefaults rPrDefault pPrDefault latentStyles lsdException style name aliases basedOn next link " +
        "autoRedefine semiHidden unhideWhenUsed qFormat uiPriority personal personalCompose personalReply rsid locked " +
        "tblStylePr numbering abstractNum nsid multiLevelType tmpl numStyleLink styleLink lvl start numFmt lvlRestart " +
        "isLgl suff lvlText lvlPicBulletId legacy lvlJc num abstractNumId lvlOverride startOverride numIdMacAtCleanup hdr ftr").Split(' '), StringComparer.Ordinal);

    public static DocumentAssemblyCandidate Assemble(IReadOnlyList<byte[]> sources, DocumentMergeOptions options,
        DocumentMergeLimits limits, CancellationToken cancellationToken)
    {
        var reports = new List<DocumentMergeSourceReport>();
        var sourceIndex = 0;
        try
        {
            var packages = new List<Package>();
            long expanded = 0;
            var parts = 0;
            foreach (var bytes in sources)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var package = Package.Read(bytes, limits, ref expanded, ref parts, cancellationToken);
                Assess(package, cancellationToken);
                packages.Add(package);
                sourceIndex++;
            }
            var first = packages[0];
            for (sourceIndex = 1; sourceIndex < packages.Count; sourceIndex++)
                CheckSharedFormatting(first, packages[sourceIndex]);

            var output = new Package();
            // Package-level properties and global rendering settings follow source 1.
            foreach (var item in first.Data.Where(item => item.Key.StartsWith("docProps/", StringComparison.Ordinal)))
                output.Add(item.Key, item.Value, first.Type(item.Key));
            foreach (var name in new[] { Settings, Fonts, Theme })
                if (first.Data.TryGetValue(name, out var data)) output.Add(name, data, first.Type(name));

            var document = new XDocument(new XElement(first.Xml(Main).Root!));
            var body = document.Root!.Element(W + "body")!;
            body.RemoveNodes();
            var styles = first.Data.ContainsKey(Styles) ? first.Xml(Styles) : new XDocument(new XElement(W + "styles"));
            styles.Root!.Elements(W + "style").Remove();
            var numbering = new XDocument(new XElement(W + "numbering"));
            var mainRels = new XElement(Rel + "Relationships");
            var rootRels = new XElement(Rel + "Relationships", new XElement(Rel + "Relationship",
                new XAttribute("Id", "officeDocument"), new XAttribute("Type", R.NamespaceName + "/officeDocument"), new XAttribute("Target", Main)));
            foreach (var rel in first.Relationships("").Where(rel => Tail((string)rel.Attribute("Type")!) != "officeDocument"))
                rootRels.Add(new XElement(rel));

            uint nextNumber = 1, nextAbstract = 0, nextDrawing = 1;
            var nextBookmark = 1;
            var nextControl = 1;
            var usedTags = new HashSet<string>(StringComparer.Ordinal);
            var evenHeaders = packages.Any(package => package.Data.ContainsKey(Settings) &&
                IsOn(package.Xml(Settings).Root!.Element(W + "evenAndOddHeaders")));
            for (sourceIndex = 0; sourceIndex < packages.Count; sourceIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var source = packages[sourceIndex];
                var prefix = "m" + (sourceIndex + 1) + "_";
                var decisions = new List<string> { "Preserve source formatting; independent list numbering; next-page source boundary." };
                var xml = source.Xml(Main);
                var sourceBody = xml.Root!.Element(W + "body")!;
                var sourceStyles = source.Data.ContainsKey(Styles) ? source.Xml(Styles) : new XDocument(new XElement(W + "styles"));
                foreach (var kind in new[] { "paragraph", "character", "table" })
                {
                    if (sourceStyles.Root!.Elements(W + "style").Any(s => (string?)s.Attribute(W + "type") == kind && IsOnAttribute(s.Attribute(W + "default")))) continue;
                    var neutral = "AssemblyDefault_" + kind;
                    while (sourceStyles.Root.Elements(W + "style").Any(s => (string?)s.Attribute(W + "styleId") == neutral)) neutral += "_";
                    sourceStyles.Root.Add(new XElement(W + "style", new XAttribute(W + "type", kind),
                        new XAttribute(W + "styleId", neutral), new XAttribute(W + "default", "1"), new XElement(W + "name", new XAttribute(W + "val", neutral))));
                }
                var styleMap = sourceStyles.Root!.Elements(W + "style").ToDictionary(
                    style => (string)style.Attribute(W + "styleId")!, style => prefix + (string)style.Attribute(W + "styleId")!, StringComparer.Ordinal);
                var defaultStyles = sourceStyles.Root.Elements(W + "style").Where(style => IsOnAttribute(style.Attribute(W + "default")))
                    .ToDictionary(style => (string)style.Attribute(W + "type")!, style => (string)style.Attribute(W + "styleId")!, StringComparer.Ordinal);
                if (source.Data.ContainsKey(Settings))
                {
                    var defaultTable = (string?)source.Xml(Settings).Root!.Element(W + "defaultTableStyle")?.Attribute(W + "val");
                    if (defaultTable is not null) defaultStyles["table"] = defaultTable;
                    if (sourceIndex == 0)
                    {
                        var settings = output.Xml(Settings);
                        settings.Root!.Element(W + "defaultTableStyle")?.Remove();
                        output.AddXml(Settings, settings, source.Type(Settings));
                    }
                }
                var nums = source.Data.ContainsKey(Numbering) ? source.Xml(Numbering) : new XDocument(new XElement(W + "numbering"));
                var numMap = nums.Root!.Elements(W + "num").ToDictionary(n => (string)n.Attribute(W + "numId")!, _ => (nextNumber++).ToString());
                var abstractMap = nums.Root.Elements(W + "abstractNum").ToDictionary(n => (string)n.Attribute(W + "abstractNumId")!, _ => (nextAbstract++).ToString());
                var contentXml = new Dictionary<string, XDocument>(StringComparer.Ordinal) { [Main] = xml };
                foreach (var name in source.Data.Keys.Where(name => source.Type(name).EndsWith(".header+xml", StringComparison.Ordinal) || source.Type(name).EndsWith(".footer+xml", StringComparison.Ordinal)))
                    contentXml.Add(name, source.Xml(name));
                var bookmarks = contentXml.Values.SelectMany(x => x.Descendants(W + "bookmarkStart")).ToArray();
                var bookmarkIds = bookmarks.ToDictionary(b => (string)b.Attribute(W + "id")!, _ => (nextBookmark++).ToString());
                var bookmarkNames = bookmarks.ToDictionary(b => (string)b.Attribute(W + "name")!, _ => "oa_" + prefix + nextBookmark++);

                foreach (var xdoc in contentXml.Values)
                {
                    Rewrite(xdoc, styleMap, numMap, abstractMap);
                    foreach (var paragraph in xdoc.Descendants(W + "p"))
                        AssignDefault(paragraph, "pPr", "pStyle", "paragraph");
                    foreach (var table in xdoc.Descendants(W + "tbl"))
                        AssignDefault(table, "tblPr", "tblStyle", "table");
                    foreach (var run in xdoc.Descendants(W + "r"))
                        AssignDefault(run, "rPr", "rStyle", "character");
                    foreach (var element in xdoc.Descendants())
                    {
                        foreach (var attr in element.Attributes().Where(a => a.Name.Namespace == R).ToArray())
                            attr.Value = prefix + attr.Value;
                        // Paragraph identity belongs to the source; remove volatile identity extensions.
                        element.Attributes().Where(a => a.Name.LocalName is "paraId" or "textId").Remove();
                        if (element.Name == Wp + "docPr" || element.Name == Pic + "cNvPr") element.SetAttributeValue("id", nextDrawing++);
                        if (element.Name == W + "bookmarkStart")
                        {
                            element.SetAttributeValue(W + "id", bookmarkIds[(string)element.Attribute(W + "id")!]);
                            element.SetAttributeValue(W + "name", bookmarkNames[(string)element.Attribute(W + "name")!]);
                        }
                        if (element.Name == W + "bookmarkEnd") element.SetAttributeValue(W + "id", Lookup(bookmarkIds, (string)element.Attribute(W + "id")!, "bookmark"));
                        if (element.Name == W + "hyperlink" && element.Attribute(W + "anchor") is XAttribute anchor)
                            anchor.Value = Lookup(bookmarkNames, anchor.Value, "bookmark target");
                        if (element.Name == W + "id" && element.Parent?.Name == W + "sdtPr") element.SetAttributeValue(W + "val", nextControl++);
                        if (element.Name == W + "tag")
                        {
                            var value = (string?)element.Attribute(W + "val") ?? "";
                            if (!usedTags.Add(value))
                            {
                                var renamed = prefix + value;
                                while (!usedTags.Add(renamed)) renamed = prefix + renamed;
                                element.SetAttributeValue(W + "val", renamed);
                                decisions.Add($"Content-control tag '{value}' renamed to '{renamed}'.");
                            }
                        }
                    }
                }
                void AssignDefault(XElement element, string propsName, string styleName, string kind)
                {
                    var props = element.Element(W + propsName);
                    if (props?.Element(W + styleName) is not null || !defaultStyles.TryGetValue(kind, out var defaultId)) return;
                    if (props is null) { props = new XElement(W + propsName); element.AddFirst(props); }
                    props.AddFirst(new XElement(W + styleName, new XAttribute(W + "val", Lookup(styleMap, defaultId, "default style"))));
                }

                Rewrite(sourceStyles, styleMap, numMap, abstractMap);
                foreach (var style in sourceStyles.Root.Elements(W + "style"))
                {
                    var oldId = (string)style.Attribute(W + "styleId")!;
                    style.SetAttributeValue(W + "styleId", styleMap[oldId]);
                    if (sourceIndex > 0) style.Attribute(W + "default")?.Remove();
                    var name = style.Element(W + "name");
                    if (sourceIndex > 0 && name is not null) name.SetAttributeValue(W + "val", prefix + (string?)name.Attribute(W + "val"));
                    styles.Root.Add(new XElement(style));
                }
                Rewrite(nums, styleMap, numMap, abstractMap);
                foreach (var number in nums.Root.Elements())
                {
                    if (number.Name == W + "num") number.SetAttributeValue(W + "numId", numMap[(string)number.Attribute(W + "numId")!]);
                    else if (number.Name == W + "abstractNum")
                    {
                        number.SetAttributeValue(W + "abstractNumId", abstractMap[(string)number.Attribute(W + "abstractNumId")!]);
                        number.Element(W + "nsid")?.Remove();
                        number.Element(W + "tmpl")?.Remove();
                    }
                    else if (number.Name == W + "numIdMacAtCleanup") continue;
                    numbering.Root!.Add(new XElement(number));
                }

                foreach (var name in source.Data.Keys.Where(name => !Globals.Contains(name) && !IsRels(name) && name != "[Content_Types].xml" && !name.StartsWith("docProps/", StringComparison.Ordinal)))
                {
                    var target = ImportedName(name, sourceIndex);
                    if (contentXml.TryGetValue(name, out var partXml)) output.AddXml(target, partXml, source.Type(name));
                    else output.Add(target, source.Data[name], source.Type(name));
                    CopyRelationships(name, target);
                }
                foreach (var rel in source.Relationships(Main))
                {
                    var type = Tail((string)rel.Attribute("Type")!);
                    if (type is "styles" or "numbering" or "settings" or "fontTable" or "theme") continue;
                    mainRels.Add(ImportRelationship(rel, Main, Main));
                }
                void CopyRelationships(string name, string target)
                {
                    var relationships = source.Relationships(name).Select(rel => ImportRelationship(rel, name, target)).ToArray();
                    if (relationships.Length > 0) output.AddXml(RelPath(target), new XDocument(new XElement(Rel + "Relationships", relationships)), RelationshipsType);
                }
                XElement ImportRelationship(XElement relationship, string owner, string targetOwner)
                {
                    var clone = new XElement(relationship);
                    clone.SetAttributeValue("Id", prefix + (string)relationship.Attribute("Id")!);
                    if ((string?)clone.Attribute("TargetMode") != "External")
                    {
                        var target = Resolve(owner, (string)clone.Attribute("Target")!);
                        clone.SetAttributeValue("Target", Relative(targetOwner, ImportedName(target, sourceIndex)));
                    }
                    return clone;
                }

                var sections = sourceBody.Descendants(W + "sectPr").ToList();
                if (sourceBody.Element(W + "sectPr") is null)
                {
                    var section = new XElement(W + "sectPr");
                    sourceBody.Add(section); sections.Add(section);
                }
                var inherited = new Dictionary<string, string>(StringComparer.Ordinal);
                var hasEven = source.Data.ContainsKey(Settings) && IsOn(source.Xml(Settings).Root!.Element(W + "evenAndOddHeaders"));
                foreach (var section in sections)
                {
                    foreach (var kind in new[] { "header", "footer" })
                    {
                        foreach (var variant in new[] { "default", "first", "even" })
                        {
                            var key = kind + variant;
                            var reference = section.Elements(W + (kind + "Reference")).SingleOrDefault(e => (string?)e.Attribute(W + "type") == variant);
                            if (reference is not null) inherited[key] = (string)reference.Attribute(R + "id")!;
                            if (!inherited.ContainsKey(key))
                            {
                                var id = prefix + "blank_" + key;
                                var name = $"word/{prefix}blank_{key}.xml";
                                output.AddXml(name, new XDocument(new XElement(W + (kind == "header" ? "hdr" : "ftr"), new XElement(W + "p"))),
                                    "application/vnd.openxmlformats-officedocument.wordprocessingml." + kind + "+xml");
                                mainRels.Add(new XElement(Rel + "Relationship", new XAttribute("Id", id), new XAttribute("Type", R.NamespaceName + "/" + kind), new XAttribute("Target", Relative(Main, name))));
                                inherited[key] = id;
                            }
                            reference?.Remove();
                            var chosen = variant == "even" && evenHeaders && !hasEven ? inherited[kind + "default"] : inherited[key];
                            section.AddFirst(new XElement(W + (kind + "Reference"), new XAttribute(W + "type", variant), new XAttribute(R + "id", chosen)));
                        }
                    }
                    var orderedReferences = section.Elements().Where(e => e.Name == W + "headerReference" || e.Name == W + "footerReference")
                        .OrderBy(e => e.Name == W + "headerReference" ? 0 : 1).Select(e => new XElement(e)).ToArray();
                    section.Elements().Where(e => e.Name == W + "headerReference" || e.Name == W + "footerReference").Remove();
                    section.AddFirst(orderedReferences);
                }
                var finalSection = sourceBody.Element(W + "sectPr")!;
                foreach (var element in sourceBody.Elements().Where(e => e != finalSection)) body.Add(new XElement(element));
                if (sourceIndex < packages.Count - 1)
                {
                    finalSection.Element(W + "type")?.Remove();
                    var references = finalSection.Elements().Where(e => e.Name == W + "headerReference" || e.Name == W + "footerReference").Last();
                    references.AddAfterSelf(new XElement(W + "type", new XAttribute(W + "val", "nextPage")));
                    // Attach the section boundary to a last paragraph where possible to avoid adding blank text.
                    var last = body.Elements().LastOrDefault();
                    if (last?.Name != W + "p" || last.Element(W + "pPr")?.Element(W + "sectPr") is not null)
                    {
                        last = new XElement(W + "p", new XElement(W + "pPr", new XElement(W + "spacing",
                            new XAttribute(W + "before", "0"), new XAttribute(W + "after", "0"), new XAttribute(W + "line", "20"), new XAttribute(W + "lineRule", "exact"))));
                        body.Add(last);
                        decisions.Add("Inserted a minimal section-boundary paragraph after non-paragraph content.");
                    }
                    var props = last.Element(W + "pPr");
                    if (props is null) { props = new XElement(W + "pPr"); last.AddFirst(props); }
                    props.Add(new XElement(finalSection));
                }
                else body.Add(new XElement(finalSection));
                decisions.Add($"Namespaced {styleMap.Count} styles, {numMap.Count} list instances, and {bookmarkIds.Count} bookmarks with {prefix}.");
                reports.Add(new DocumentMergeSourceReport
                {
                    Index = sourceIndex,
                    Paragraphs = xml.Descendants(W + "p").Count(),
                    Tables = xml.Descendants(W + "tbl").Count(),
                    Images = source.Data.Keys.Count(name => source.Type(name).StartsWith("image/", StringComparison.Ordinal)),
                    Sections = sections.Count,
                    Decisions = decisions
                });
            }
            // Abstract definitions must precede concrete numbering instances in schema order.
            var numberChildren = numbering.Root!.Elements().OrderBy(e => e.Name == W + "abstractNum" ? 0 : 1).Select(e => new XElement(e)).ToArray();
            numbering.Root.ReplaceNodes(numberChildren);
            output.AddXml(Main, document, first.Type(Main));
            output.AddXml(Styles, styles, "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml");
            output.AddXml(Numbering, numbering, "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml");
            foreach (var pair in new[] { (Styles, "styles"), (Numbering, "numbering"), (Settings, "settings"), (Fonts, "fontTable"), (Theme, "theme") })
                if (output.Data.ContainsKey(pair.Item1)) mainRels.Add(new XElement(Rel + "Relationship", new XAttribute("Id", "global_" + pair.Item2),
                    new XAttribute("Type", R.NamespaceName + "/" + pair.Item2), new XAttribute("Target", Relative(Main, pair.Item1))));
            if (evenHeaders)
            {
                var settings = output.Data.ContainsKey(Settings) ? output.Xml(Settings) : new XDocument(new XElement(W + "settings"));
                settings.Root!.Element(W + "evenAndOddHeaders")?.Remove();
                // SDK validation below enforces ordering; insert before later settings according to schema.
                var insertBefore = settings.Root.Elements().FirstOrDefault(e => !new[] { "writeProtection", "view", "zoom", "removePersonalInformation", "removeDateAndTime", "doNotDisplayPageBoundaries", "displayBackgroundShape", "printPostScriptOverText", "printFractionalCharacterWidth", "printFormsData", "embedTrueTypeFonts", "embedSystemFonts", "saveSubsetFonts", "saveFormsData", "mirrorMargins", "alignBordersAndEdges", "bordersDoNotSurroundHeader", "bordersDoNotSurroundFooter", "gutterAtTop", "hideSpellingErrors", "hideGrammaticalErrors", "activeWritingStyle", "proofState", "formsDesign", "attachedTemplate", "linkStyles", "stylePaneFormatFilter", "stylePaneSortMethod", "documentType", "mailMerge", "revisionView", "trackRevisions", "doNotTrackMoves", "doNotTrackFormatting", "documentProtection", "autoFormatOverride", "styleLockTheme", "styleLockQFSet", "defaultTabStop", "autoHyphenation", "consecutiveHyphenLimit", "hyphenationZone", "doNotHyphenateCaps", "showEnvelope", "summaryLength", "clickAndTypeStyle", "defaultTableStyle" }.Contains(e.Name.LocalName));
                if (insertBefore is null) settings.Root.Add(new XElement(W + "evenAndOddHeaders")); else insertBefore.AddBeforeSelf(new XElement(W + "evenAndOddHeaders"));
                output.AddXml(Settings, settings, "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml");
                if (!mainRels.Elements().Any(e => (string?)e.Attribute("Id") == "global_settings")) mainRels.Add(new XElement(Rel + "Relationship", new XAttribute("Id", "global_settings"), new XAttribute("Type", R.NamespaceName + "/settings"), new XAttribute("Target", "settings.xml")));
            }
            output.AddXml(RelPath(Main), new XDocument(mainRels), RelationshipsType);
            output.AddXml("_rels/.rels", new XDocument(rootRels), RelationshipsType);
            SetMetadata(output, options);
            sourceIndex = -1;
            var result = output.Write(limits, cancellationToken);
            ValidateSchema(result, cancellationToken);
            return new DocumentAssemblyCandidate { Content = result, Sources = reports };
        }
        catch (Exception ex) when (ex is MergeException or InvalidDataException or XmlException or OpenXmlPackageException or ArgumentException)
        {
            return new DocumentAssemblyCandidate
            {
                Sources = reports,
                Diagnostics = new[] { new WorkflowDiagnostic
                {
                    Code = ex is MergeException merge ? merge.Code : AssemblyDiagnosticCodes.InvalidMergePackage,
                    Message = ex.Message, Path = (sourceIndex < 0 ? "output" : "sources/" + sourceIndex) + (ex is MergeException m ? "/" + m.Part : "")
                } }
            };
        }
    }

    private static void Rewrite(XDocument document, Dictionary<string, string> styles, Dictionary<string, string> nums, Dictionary<string, string> abstracts)
    {
        foreach (var element in document.Descendants())
        {
            var value = element.Attribute(W + "val");
            if (value is null) continue;
            if (element.Name.Namespace != W) continue;
            if (element.Name.LocalName is "pStyle" or "rStyle" or "tblStyle" or "basedOn" or "next" or "link" or "numStyleLink" or "styleLink")
                value.Value = Lookup(styles, value.Value, "style");
            if (element.Name == W + "numId" && value.Value != "0") value.Value = Lookup(nums, value.Value, "numbering");
            if (element.Name == W + "abstractNumId") value.Value = Lookup(abstracts, value.Value, "abstract numbering");
        }
    }

    private static string Lookup(Dictionary<string, string> map, string value, string kind) => map.TryGetValue(value, out var result)
        ? result : throw new MergeException(AssemblyDiagnosticCodes.MergeUnresolvedReference, $"Missing {kind} definition '{value}'.", Main);

    private static void Assess(Package package, CancellationToken cancellationToken)
    {
        if (!package.Data.ContainsKey(Main) || package.Type(Main) != "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")
            throw new MergeException(AssemblyDiagnosticCodes.UnsupportedMergeFormat, "Assembly requires a standard .docx package with word/document.xml.", Main);
        foreach (var name in package.Data.Keys)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (name == "[Content_Types].xml") continue;
            var type = package.Type(name);
            if (!AllowedTypes.Contains(type) || type is "image/x-emf" or "image/x-wmf")
                throw new MergeException(AssemblyDiagnosticCodes.UnsupportedMergePart, $"Unsupported part content type '{type}'.", name);
            if (IsRels(name))
            {
                var owner = Owner(name);
                if (owner.Length > 0 && !package.Data.ContainsKey(owner)) throw new MergeException(AssemblyDiagnosticCodes.MergeUnresolvedReference, "Relationship owner is missing.", name);
                var ids = new HashSet<string>(StringComparer.Ordinal);
                foreach (var relationship in package.Xml(name).Root!.Elements())
                {
                    var relType = (string?)relationship.Attribute("Type") ?? "";
                    var tail = Tail(relType);
                    var expectedType = tail == "core-properties"
                        ? "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
                        : R.NamespaceName + "/" + tail;
                    if (!AllowedRelationships.Contains(tail) || relType != expectedType || !ids.Add((string?)relationship.Attribute("Id") ?? ""))
                        throw new MergeException(AssemblyDiagnosticCodes.UnsupportedMergeRelationship, "Unsupported or duplicate relationship.", name);
                    var target = (string?)relationship.Attribute("Target") ?? "";
                    if ((string?)relationship.Attribute("TargetMode") == "External")
                    {
                        if (tail != "hyperlink" || !Uri.TryCreate(target, UriKind.Absolute, out var uri) || uri.Scheme is not ("http" or "https" or "mailto"))
                            throw new MergeException(AssemblyDiagnosticCodes.UnsupportedMergeExternalContent, "Only http, https, and mailto external hyperlinks are supported; external content is never fetched.", name);
                    }
                    else if (!package.Data.ContainsKey(Resolve(owner, target)))
                        throw new MergeException(AssemblyDiagnosticCodes.MergeUnresolvedReference, "Internal relationship target is missing.", name);
                    else
                    {
                        var resolved = Resolve(owner, target);
                        var canonical = tail switch
                        {
                            "officeDocument" => Main,
                            "styles" => Styles,
                            "numbering" => Numbering,
                            "settings" => Settings,
                            "fontTable" => Fonts,
                            "theme" => Theme,
                            _ => null
                        };
                        if (canonical is not null && (resolved != canonical || owner != (tail == "officeDocument" ? "" : Main)))
                            throw new MergeException(AssemblyDiagnosticCodes.UnsupportedMergePartLayout, "Shared parts must use the standard Word package layout.", name);
                    }
                }
            }
            else if (type.EndsWith("+xml", StringComparison.Ordinal))
            {
                var xml = package.Xml(name);
                var content = name == Main || name == Styles || name == Numbering || type.EndsWith(".header+xml", StringComparison.Ordinal) || type.EndsWith(".footer+xml", StringComparison.Ordinal);
                if (content)
                {
                    foreach (var element in xml.Descendants())
                    {
                        if (element.Name.Namespace == W && !WordElements.Contains(element.Name.LocalName) ||
                            element.Name.Namespace != W && element.Name.Namespace != Wp && element.Name.Namespace != A && element.Name.Namespace != Pic)
                            throw new MergeException(AssemblyDiagnosticCodes.UnsupportedMergeContent, $"Unsupported content element '{element.Name.LocalName}'.", name);
                        if (element.Name == W + "sdtPr" && element.Element(W + "text") is null)
                            throw new MergeException(AssemblyDiagnosticCodes.UnsupportedMergeContentControl, "Only plain-text content controls are supported.", name);
                        if (element.Name == W + "lvlPicBulletId" || element.Name == W + "printerSettings" || element.Name == W + "placeholder")
                            throw new MergeException(AssemblyDiagnosticCodes.UnsupportedMergeContent, $"Unsupported dependency '{element.Name.LocalName}'.", name);
                        if (element.Name == A + "graphicData" && (string?)element.Attribute("uri") != Pic.NamespaceName)
                            throw new MergeException(AssemblyDiagnosticCodes.UnsupportedMergeContent, "Only embedded raster picture drawings are supported.", name);
                        foreach (var attribute in element.Attributes().Where(a => a.Name.Namespace == R))
                            if (!package.Relationships(name).Any(rel => (string?)rel.Attribute("Id") == attribute.Value))
                                throw new MergeException(AssemblyDiagnosticCodes.MergeUnresolvedReference, "Content relationship is missing.", name);
                    }
                    ValidateFields(xml, name);
                }
            }
        }
        var body = package.Xml(Main).Root?.Element(W + "body");
        if (body is null) throw new MergeException(AssemblyDiagnosticCodes.InvalidMergePackage, "Missing document body.", Main);
        if (package.Data.ContainsKey(Settings))
        {
            var settings = package.Xml(Settings);
            if (settings.Descendants().Any(e => e.Name.LocalName is "trackRevisions" or "documentProtection" or "mailMerge" or "attachedTemplate" or "saveThroughXslt"))
                throw new MergeException(AssemblyDiagnosticCodes.UnsupportedMergeSettings, "Review, protection, or external processing settings are unsupported.", Settings);
        }
        // Validate each input before importing it, including schema ordering and references known to the SDK.
        ValidateSchema(package.Original!, cancellationToken);
    }

    private static void ValidateFields(XDocument xml, string part)
    {
        static bool Allowed(string instruction) => new[] { "PAGE", "NUMPAGES", "SECTIONPAGES" }.Contains(instruction.Trim(), StringComparer.OrdinalIgnoreCase);
        foreach (var field in xml.Descendants(W + "fldSimple"))
            if (!Allowed((string?)field.Attribute(W + "instr") ?? ""))
                throw new MergeException(AssemblyDiagnosticCodes.UnsupportedMergeField, "Only PAGE, NUMPAGES, and SECTIONPAGES fields without switches are supported.", part);
        var stack = new Stack<StringBuilder>();
        foreach (var element in xml.Descendants())
        {
            if (element.Name == W + "fldChar")
            {
                var type = (string?)element.Attribute(W + "fldCharType");
                if (type == "begin") stack.Push(new StringBuilder());
                else if (type == "separate" || type == "end")
                {
                    if (stack.Count == 0 || !Allowed(stack.Peek().ToString()))
                        throw new MergeException(AssemblyDiagnosticCodes.UnsupportedMergeField, "Unsupported or unbalanced field instruction.", part);
                    if (type == "end") stack.Pop();
                }
            }
            else if (element.Name == W + "instrText")
            {
                if (stack.Count == 0) throw new MergeException(AssemblyDiagnosticCodes.UnsupportedMergeField, "Unbalanced field instruction.", part);
                stack.Peek().Append(element.Value);
            }
        }
        if (stack.Count != 0) throw new MergeException(AssemblyDiagnosticCodes.UnsupportedMergeField, "Unbalanced field instruction.", part);
    }

    private static void CheckSharedFormatting(Package first, Package next)
    {
        foreach (var name in new[] { Theme, Fonts, Settings })
        {
            var left = first.Data.ContainsKey(name) ? first.Xml(name).Root : null;
            var right = next.Data.ContainsKey(name) ? next.Xml(name).Root : null;
            if (name == Settings)
            {
                var ignored = new[] { "view", "zoom", "proofState", "rsids", "evenAndOddHeaders", "docId" };
                left?.Elements().Where(e => ignored.Contains(e.Name.LocalName)).Remove();
                right?.Elements().Where(e => ignored.Contains(e.Name.LocalName)).Remove();
                if (left is null) left = new XElement(W + "settings");
                if (right is null) right = new XElement(W + "settings");
            }
            if (!Equivalent(left, right)) throw new MergeException(AssemblyDiagnosticCodes.MergeIncompatibleFormatting,
                "Source themes, font tables, and layout-affecting settings must agree; normalize them explicitly before assembly.", name);
        }
        XElement? Defaults(Package p) => p.Data.ContainsKey(Styles) ? p.Xml(Styles).Root!.Element(W + "docDefaults") : null;
        if (!Equivalent(Defaults(first), Defaults(next))) throw new MergeException(AssemblyDiagnosticCodes.MergeIncompatibleFormatting, "Document defaults differ; assembly cannot promise source formatting preservation.", Styles);
    }

    private static bool Equivalent(XElement? left, XElement? right)
    {
        if (left is null || right is null) return left is null && right is null;
        XElement Clean(XElement element) => new(element.Name,
            element.Attributes().Where(a => !a.IsNamespaceDeclaration).OrderBy(a => a.Name.ToString()).Select(a => new XAttribute(a)),
            element.Nodes().Where(n => n is not XText t || !string.IsNullOrWhiteSpace(t.Value)).Select(n => n is XElement child ? (object)Clean(child) : n));
        return XNode.DeepEquals(Clean(left), Clean(right));
    }

    private static void ValidateSchema(byte[] bytes, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        using var stream = new MemoryStream(bytes, writable: false);
        using var document = WordprocessingDocument.Open(stream, false);
        var error = new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019).Validate(document).FirstOrDefault();
        if (error is not null) throw new MergeException(AssemblyDiagnosticCodes.InvalidMergePackage, error.Description, error.Part?.Uri.ToString() ?? Main);
        cancellationToken.ThrowIfCancellationRequested();
    }

    private static void SetMetadata(Package package, DocumentMergeOptions options)
    {
        if (options.Title is null && options.Author is null) return;
        const string part = "docProps/core.xml";
        XNamespace cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        XNamespace dc = "http://purl.org/dc/elements/1.1/";
        var xml = package.Data.ContainsKey(part) ? package.Xml(part) : new XDocument(new XElement(cp + "coreProperties"));
        if (options.Title is not null) xml.Root!.SetElementValue(dc + "title", options.Title);
        if (options.Author is not null) xml.Root!.SetElementValue(dc + "creator", options.Author);
        package.AddXml(part, xml, "application/vnd.openxmlformats-package.core-properties+xml");
        var rels = package.Xml("_rels/.rels");
        if (!rels.Root!.Elements().Any(e => Tail((string?)e.Attribute("Type") ?? "") == "core-properties"))
        {
            rels.Root.Add(new XElement(Rel + "Relationship", new XAttribute("Id", "mergeCoreProperties"),
                new XAttribute("Type", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"), new XAttribute("Target", part)));
            package.AddXml("_rels/.rels", rels, RelationshipsType);
        }
    }

    private const string RelationshipsType = "application/vnd.openxmlformats-package.relationships+xml";
    private static bool IsOn(XElement? element) => element is not null && (element.Attribute(W + "val") is null || IsOnAttribute(element.Attribute(W + "val")));
    private static bool IsOnAttribute(XAttribute? attribute) => attribute?.Value is "true" or "1" or "on";
    private static string Tail(string value) => value.Substring(value.LastIndexOf('/') + 1);
    private static bool IsRels(string name) => name.EndsWith(".rels", StringComparison.Ordinal);
    private static string ImportedName(string name, int source) => "merge/" + source + "/" + name;
    private static string RelPath(string name) { var slash = name.LastIndexOf('/'); return name.Substring(0, slash + 1) + "_rels/" + name.Substring(slash + 1) + ".rels"; }
    private static string Owner(string path) => path == "_rels/.rels" ? "" : path.Replace("/_rels/", "/").Substring(0, path.Replace("/_rels/", "/").Length - 5);
    private static string Resolve(string owner, string target)
    {
        var uri = new Uri(new Uri("http://package/" + owner), target);
        if (uri.Host != "package" || uri.Scheme != "http" || uri.Query.Length > 0 || uri.Fragment.Length > 0)
            throw new MergeException(AssemblyDiagnosticCodes.MergeUnresolvedReference, "Invalid internal package URI.", owner);
        return Uri.UnescapeDataString(uri.AbsolutePath.TrimStart('/'));
    }
    private static string Relative(string owner, string target) => new Uri("http://package/" + owner).MakeRelativeUri(new Uri("http://package/" + target)).ToString();

    private sealed class MergeException : Exception
    {
        public string Code { get; }
        public string Part { get; }
        public MergeException(string code, string message, string part) : base(message) { Code = code; Part = part; }
    }

    private sealed class Package
    {
        public Dictionary<string, byte[]> Data { get; } = new(StringComparer.Ordinal);
        private readonly Dictionary<string, string> _types = new(StringComparer.Ordinal);
        public byte[]? Original { get; private set; }
        public string Type(string name) => _types.TryGetValue(name, out var type) ? type : "";
        public void Add(string name, byte[] bytes, string type) { Data[name] = bytes; _types[name] = type; }
        public void AddXml(string name, XDocument xml, string type)
        {
            using var stream = new MemoryStream();
            xml.Save(stream, SaveOptions.DisableFormatting);
            Add(name, stream.ToArray(), type);
        }
        public XDocument Xml(string name)
        {
            using var stream = new MemoryStream(Data[name], writable: false);
            using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = Data[name].LongLength + 1 });
            return XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        }
        public IEnumerable<XElement> Relationships(string owner) => Data.ContainsKey(owner.Length == 0 ? "_rels/.rels" : RelPath(owner))
            ? Xml(owner.Length == 0 ? "_rels/.rels" : RelPath(owner)).Root!.Elements(Rel + "Relationship") : Enumerable.Empty<XElement>();
        public static Package Read(byte[] bytes, DocumentMergeLimits limits, ref long expanded, ref int parts, CancellationToken cancellationToken)
        {
            var package = new Package { Original = bytes };
            using var stream = new MemoryStream(bytes, writable: false);
            using var zip = new ZipArchive(stream, ZipArchiveMode.Read);
            foreach (var entry in zip.Entries)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (++parts > limits.MaximumParts || entry.Length > limits.MaximumExpandedBytes - expanded)
                    throw new MergeException(AssemblyDiagnosticCodes.MergeResourceLimit, "Expanded package size or part count exceeds the host limit.", entry.FullName);
                if (entry.FullName.StartsWith("/", StringComparison.Ordinal) || entry.FullName.Contains("\\") || entry.FullName.Split('/').Any(s => s is ".." or "." or "") || package.Data.ContainsKey(entry.FullName))
                    throw new MergeException(AssemblyDiagnosticCodes.InvalidMergePackage, "Noncanonical or duplicate package entry.", entry.FullName);
                using var input = entry.Open();
                using var output = new MemoryStream();
                var buffer = new byte[81920];
                int read;
                while ((read = input.Read(buffer, 0, buffer.Length)) != 0)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    expanded += read;
                    if (expanded > limits.MaximumExpandedBytes) throw new MergeException(AssemblyDiagnosticCodes.MergeResourceLimit, "Expanded input exceeds the host limit.", entry.FullName);
                    output.Write(buffer, 0, read);
                }
                package.Data.Add(entry.FullName, output.ToArray());
            }
            if (!package.Data.ContainsKey("[Content_Types].xml")) throw new InvalidDataException("Missing package content types.");
            var types = package.Xml("[Content_Types].xml");
            foreach (var name in package.Data.Keys)
            {
                var type = types.Root!.Elements(Ct + "Override").SingleOrDefault(e => (string?)e.Attribute("PartName") == "/" + name)?.Attribute("ContentType")?.Value;
                type ??= types.Root.Elements(Ct + "Default").SingleOrDefault(e => (string?)e.Attribute("Extension") == Path.GetExtension(name).TrimStart('.'))?.Attribute("ContentType")?.Value;
                package._types[name] = type ?? "";
            }
            return package;
        }
        public byte[] Write(DocumentMergeLimits limits, CancellationToken cancellationToken)
        {
            if (Data.Count + 1 > limits.MaximumParts || Data.Sum(p => p.Value.LongLength) > limits.MaximumExpandedBytes)
                throw new MergeException(AssemblyDiagnosticCodes.MergeResourceLimit, "Assembled expanded content exceeds the host limit.", "package");
            var types = new XElement(Ct + "Types", Data.Keys.OrderBy(n => n, StringComparer.Ordinal).Select(name =>
                new XElement(Ct + "Override", new XAttribute("PartName", "/" + name), new XAttribute("ContentType", Type(name)))));
            AddXml("[Content_Types].xml", new XDocument(types), "");
            using var stream = new MemoryStream();
            using (var zip = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
            {
                foreach (var item in Data.OrderBy(p => p.Key, StringComparer.Ordinal))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var entry = zip.CreateEntry(item.Key);
                    entry.LastWriteTime = new DateTimeOffset(2000, 1, 1, 0, 0, 0, TimeSpan.Zero);
                    using (var output = entry.Open()) output.Write(item.Value, 0, item.Value.Length);
                    if (stream.Length > limits.MaximumOutputBytes) throw new MergeException(AssemblyDiagnosticCodes.MergeResourceLimit, "Assembled output exceeds the host limit.", "package");
                }
            }
            if (stream.Length > limits.MaximumOutputBytes) throw new MergeException(AssemblyDiagnosticCodes.MergeResourceLimit, "Assembled output exceeds the host limit.", "package");
            return stream.ToArray();
        }
    }
}
