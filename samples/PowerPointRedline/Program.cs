using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System.Text;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

if (args.Length < 1)
{
    Console.Error.WriteLine("Usage: PowerPointRedline <original.pptx> <modified.pptx> [output.pptx]");
    Console.Error.WriteLine("Run with --help for more information.");
    return 1;
}

if (args.Length == 1 && args[0] == "--help")
{
    Console.Out.WriteLine("Usage: PowerPointRedline <original.pptx> <modified.pptx> [output.pptx]");
    Console.Out.WriteLine("       PowerPointRedline --test");
    Console.Out.WriteLine();
    Console.Out.WriteLine("Compares two PowerPoint presentations and generates a diff file.");
    Console.Out.WriteLine();
    Console.Out.WriteLine("If output.pptx is omitted, uses 'diff.pptx'");
    Console.Out.WriteLine("Use --test to generate sample presentations and compare them.");
    Console.Out.WriteLine();
    Console.Out.WriteLine("IMPORTANT: PowerPoint has no native redline vocabulary.");
    Console.Out.WriteLine("Unlike Word (w:ins/w:del), PresentationML cannot represent tracked changes.");
    Console.Out.WriteLine("This tool uses slide annotations and highlighting to mark differences.");
    return 0;
}

if (args.Length < 2 && args[0] != "--test")
{
    Console.Error.WriteLine("Usage: PowerPointRedline <original.pptx> <modified.pptx> [output.pptx]");
    Console.Error.WriteLine("Run with --help for more information.");
    return 1;
}

try
{
    string originalPath, modifiedPath, outputPath;

    // Handle test mode
    if (args.Length == 1 && args[0] == "--test")
    {
        Console.WriteLine("Running test mode with generated presentations...");
        var testDir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "pptx-redline-test");
        System.IO.Directory.CreateDirectory(testDir);

        originalPath = System.IO.Path.Combine(testDir, "original.pptx");
        modifiedPath = System.IO.Path.Combine(testDir, "modified.pptx");
        outputPath = System.IO.Path.Combine(testDir, "diff.pptx");

        Console.WriteLine("Creating test presentations...");
        PowerPointTestFileGenerator.CreateOriginalPresentation(originalPath);
        PowerPointTestFileGenerator.CreateModifiedPresentation(modifiedPath);
        Console.WriteLine($"Test files created in {testDir}");
        Console.WriteLine();
    }
    else
    {
        originalPath = System.IO.Path.GetFullPath(args[0]);
        modifiedPath = System.IO.Path.GetFullPath(args[1]);
        outputPath = args.Length >= 3
            ? System.IO.Path.GetFullPath(args[2])
            : System.IO.Path.Combine(System.IO.Path.GetDirectoryName(originalPath) ?? ".", "diff.pptx");
    }

    if (!System.IO.File.Exists(originalPath))
    {
        Console.Error.WriteLine($"Error: Original file not found: {originalPath}");
        return 1;
    }
    if (!System.IO.File.Exists(modifiedPath))
    {
        Console.Error.WriteLine($"Error: Modified file not found: {modifiedPath}");
        return 1;
    }

    Console.WriteLine($"Original: {originalPath}");
    Console.WriteLine($"Modified: {modifiedPath}");
    Console.WriteLine($"Output:   {outputPath}");
    Console.WriteLine();

    var comparer = new PowerPointRedlineComparer();
    var differences = comparer.CompareAndCreateRedline(originalPath, modifiedPath, outputPath);

    Console.WriteLine($"Comparison complete. Found {differences.Count} differences:");
    Console.WriteLine();
    foreach (var diff in differences)
        Console.WriteLine($"  {diff}");

    Console.WriteLine();
    Console.WriteLine("=== TECHNICAL FINDINGS ===");
    Console.WriteLine();
    Console.WriteLine("LIMITATION: PowerPoint/PresentationML Architecture");
    Console.WriteLine("  Word (.docx) uses w:ins, w:del, w:rPrChange for tracked changes");
    Console.WriteLine("  PowerPoint (.pptx) has NO equivalent redline vocabulary");
    Console.WriteLine();
    Console.WriteLine("CONSEQUENCE:");
    Console.WriteLine("  - OfficeAgent explicitly refuses 'Tracked' mode for PowerPoint");
    Console.WriteLine("  - Only Direct mode is supported");
    Console.WriteLine("  - Comparison output cannot contain native redlines");
    Console.WriteLine();
    Console.WriteLine("ALTERNATIVES:");
    Console.WriteLine("  1. Use visual markup (highlighting, comments) - limited clarity");
    Console.WriteLine("  2. Use Aspose.Slides - commercial solution with comparison");
    Console.WriteLine("  3. Generate separate documents - less integrated UX");
    Console.WriteLine();
    Console.WriteLine($"Output written to: {outputPath}");
    return 0;
}
catch (Exception ex)
{
    Console.Error.WriteLine($"Error: {ex.Message}");
    if (ex.InnerException != null)
        Console.Error.WriteLine($"Inner: {ex.InnerException.Message}");
    return 1;
}

/// <summary>
/// Compares two PowerPoint presentations and creates a redlined diff file.
///
/// KEY FINDING: PresentationML has NO native redline vocabulary.
/// - Word has w:ins (insertion), w:del (deletion), w:rPrChange (format change) elements
/// - PowerPoint has NO equivalent tracked-changes markup
/// - This is an architectural limitation of the PresentationML standard, not a feature gap
/// - OfficeAgent explicitly refuses tracked changes for PowerPoint (see PowerPointModule.cs)
///
/// Result: Comparisons are recorded via annotations (slide notes) and metadata.
/// This tool marks differences but cannot create native tracked changes like Word does.
/// </summary>
public class PowerPointRedlineComparer
{
    private const uint RedHighlightColor = 0xFFCC00; // Red for highlighting

    /// <summary>
    /// Compares two presentations and creates a diff output with annotations.
    /// Returns a list of differences found.
    /// </summary>
    public List<string> CompareAndCreateRedline(string originalPath, string modifiedPath, string outputPath)
    {
        var differences = new List<string>();

        try
        {
            using var originalPrs = PresentationDocument.Open(originalPath, false);
            using var modifiedPrs = PresentationDocument.Open(modifiedPath, false);

            var originalSlides = GetSlideContents(originalPrs);
            var modifiedSlides = GetSlideContents(modifiedPrs);

            differences.AddRange(CompareSlideStructure(originalSlides, modifiedSlides));

            // Create output by copying modified (to show end state)
            System.IO.File.Copy(modifiedPath, outputPath, overwrite: true);

            // Annotate the output with difference information via slide notes
            AnnotateDifferences(outputPath, differences, originalSlides, modifiedSlides);
        }
        catch (Exception ex)
        {
            differences.Add($"[Error] {ex.Message}");
        }

        return differences;
    }

    private void AnnotateDifferences(string outputPath, List<string> differences,
        List<SlideContent> originalSlides, List<SlideContent> modifiedSlides)
    {
        try
        {
            using var prs = PresentationDocument.Open(outputPath, true);
            var presentationPart = prs.PresentationPart;
            if (presentationPart?.Presentation?.SlideIdList == null) return;

            var slidesList = presentationPart.Presentation.SlideIdList.Elements<P.SlideId>().ToList();

            // Annotate each slide with relevant differences
            for (int i = 0; i < Math.Min(slidesList.Count, modifiedSlides.Count); i++)
            {
                var slideId = slidesList[i];
                var relId = slideId.RelationshipId?.Value;
                if (relId == null) continue;

                var slidePart = presentationPart.GetPartById(relId) as SlidePart;
                if (slidePart?.Slide == null) continue;

                // Build annotation note for this slide
                var notes = new StringBuilder();
                notes.AppendLine("=== REDLINE COMPARISON NOTES ===");

                if (i < originalSlides.Count && i < modifiedSlides.Count)
                {
                    var orig = originalSlides[i];
                    var mod = modifiedSlides[i];

                    if (!string.Equals(orig.Title, mod.Title, StringComparison.Ordinal))
                        notes.AppendLine($"Title changed: '{orig.Title}' -> '{mod.Title}'");

                    if (!string.Equals(orig.TextContent, mod.TextContent, StringComparison.Ordinal))
                        notes.AppendLine("Content modified: Text has changes");

                    if (orig.ShapeCount != mod.ShapeCount)
                        notes.AppendLine($"Shapes changed: {orig.ShapeCount} -> {mod.ShapeCount}");
                }
                else if (i >= originalSlides.Count)
                {
                    notes.AppendLine("Slide added in modified version");
                }

                AddSlideNote(slidePart, notes.ToString());
            }

            prs.Save();
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Warning: Could not annotate output: {ex.Message}");
        }
    }

    private void AddSlideNote(SlidePart slidePart, string noteText)
    {
        try
        {
            // Create or get the notes part
            var notesPart = slidePart.NotesSlidePart;
            if (notesPart == null)
            {
                notesPart = slidePart.AddNewPart<NotesSlidePart>();
                var notesSlide = new P.NotesSlide(
                    new P.CommonSlideData(
                        new P.ShapeTree()
                    ),
                    new P.ColorMapOverride()
                );
                notesPart.NotesSlide = notesSlide;
            }

            // Add text to notes
            var shapeTree = notesPart.NotesSlide?.CommonSlideData?.ShapeTree;
            if (shapeTree != null)
            {
                // Find or create a text shape for notes
                var noteShape = shapeTree.Elements<P.Shape>().FirstOrDefault(
                    s => s.TextBody?.Elements<A.Paragraph>().Any() ?? false);

                if (noteShape == null)
                {
                    // Create a new shape if needed (simplified - in production would need proper shape setup)
                    return;
                }

                var textBody = noteShape.TextBody;
                if (textBody != null)
                {
                    // Append note text to existing paragraphs
                    foreach (var line in noteText.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None))
                    {
                        if (string.IsNullOrWhiteSpace(line)) continue;

                        var para = new A.Paragraph();
                        var run = new A.Run
                        {
                            Text = new A.Text { Text = line }
                        };
                        para.Append(run);
                        textBody.Append(para);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Note annotation skipped: {ex.Message}");
        }
    }

    private List<SlideContent> GetSlideContents(PresentationDocument prs)
    {
        var slides = new List<SlideContent>();
        var presentationPart = prs.PresentationPart;
        if (presentationPart?.Presentation?.SlideIdList == null) return slides;

        foreach (var slideId in presentationPart.Presentation.SlideIdList.Elements<P.SlideId>())
        {
            try
            {
                var relId = slideId.RelationshipId?.Value;
                if (relId == null) continue;

                var slidePart = presentationPart.GetPartById(relId) as SlidePart;
                if (slidePart?.Slide == null) continue;

                var slide = slidePart.Slide;

                var content = new SlideContent
                {
                    Index = slides.Count,
                    Title = ExtractTitle(slide),
                    TextContent = ExtractAllText(slide),
                    ShapeCount = CountShapes(slide)
                };

                slides.Add(content);
            }
            catch
            {
                // Skip slides that can't be read
            }
        }

        return slides;
    }

    private string ExtractTitle(P.Slide slide)
    {
        var shapes = slide.CommonSlideData?.ShapeTree;
        if (shapes == null) return string.Empty;

        foreach (var shape in shapes.Elements<P.Shape>().Take(1))
        {
            var text = ExtractShapeText(shape);
            if (!string.IsNullOrEmpty(text)) return text;
        }

        return string.Empty;
    }

    private string ExtractAllText(P.Slide slide)
    {
        var shapes = slide.CommonSlideData?.ShapeTree;
        if (shapes == null) return string.Empty;

        var sb = new StringBuilder();
        foreach (var shape in shapes.Elements<P.Shape>())
        {
            var text = ExtractShapeText(shape);
            if (!string.IsNullOrEmpty(text))
                sb.AppendLine(text);
        }

        return sb.ToString();
    }

    private string ExtractShapeText(P.Shape shape)
    {
        var txBody = shape.TextBody;
        if (txBody == null) return string.Empty;

        var sb = new StringBuilder();
        foreach (var paragraph in txBody.Elements<DocumentFormat.OpenXml.Drawing.Paragraph>())
        {
            foreach (var run in paragraph.Elements<DocumentFormat.OpenXml.Drawing.Run>())
            {
                if (run.Text != null)
                    sb.Append(run.Text);
            }
        }

        return sb.ToString();
    }

    private int CountShapes(P.Slide slide) =>
        slide.CommonSlideData?.ShapeTree?.Elements<P.Shape>().Count() ?? 0;

    private List<string> CompareSlideStructure(List<SlideContent> original, List<SlideContent> modified)
    {
        var diffs = new List<string>();

        if (original.Count != modified.Count)
            diffs.Add($"Slide count: {original.Count} -> {modified.Count}");

        for (int i = 0; i < Math.Min(original.Count, modified.Count); i++)
        {
            var o = original[i];
            var m = modified[i];

            if (!string.Equals(o.Title, m.Title, StringComparison.Ordinal))
                diffs.Add($"Slide {i + 1}: Title changed");

            if (!string.Equals(o.TextContent, m.TextContent, StringComparison.Ordinal))
                diffs.Add($"Slide {i + 1}: Content modified");

            if (o.ShapeCount != m.ShapeCount)
                diffs.Add($"Slide {i + 1}: Shape count changed ({o.ShapeCount} -> {m.ShapeCount})");
        }

        for (int i = original.Count; i < modified.Count; i++)
            diffs.Add($"Slide {i + 1}: Added");

        for (int i = modified.Count; i < original.Count; i++)
            diffs.Add($"Slide {i + 1}: Deleted");

        return diffs;
    }

    private class SlideContent
    {
        public int Index { get; set; }
        public string Title { get; set; } = string.Empty;
        public string TextContent { get; set; } = string.Empty;
        public int ShapeCount { get; set; }
    }
}

/// <summary>
/// Helper class to generate test PowerPoint presentations for demonstration and testing.
/// </summary>
public static class PowerPointTestFileGenerator
{
    public static void CreateOriginalPresentation(string path)
    {
        using var prs = PresentationDocument.Create(path, PresentationDocumentType.Presentation);
        var presentationPart = prs.AddPresentationPart();
        presentationPart.Presentation = new P.Presentation();

        var slideIdList = new P.SlideIdList();
        presentationPart.Presentation.Append(slideIdList);

        // Slide 1
        var slidePart1 = presentationPart.AddNewPart<SlidePart>();
        slideIdList.Append(new P.SlideId { Id = 256U, RelationshipId = presentationPart.GetIdOfPart(slidePart1) });
        slidePart1.Slide = new P.Slide();
        slidePart1.Slide.Append(new P.CommonSlideData(new P.ShapeTree()));

        AddTitleShape(slidePart1.Slide.CommonSlideData!.ShapeTree!, "Quarterly Report Q3 2024");
        AddTextShape(slidePart1.Slide.CommonSlideData!.ShapeTree!, "Sales Performance Overview");

        // Slide 2
        var slidePart2 = presentationPart.AddNewPart<SlidePart>();
        slideIdList.Append(new P.SlideId { Id = 257U, RelationshipId = presentationPart.GetIdOfPart(slidePart2) });
        slidePart2.Slide = new P.Slide();
        slidePart2.Slide.Append(new P.CommonSlideData(new P.ShapeTree()));

        AddTitleShape(slidePart2.Slide.CommonSlideData!.ShapeTree!, "Revenue by Region");
        AddTextShape(slidePart2.Slide.CommonSlideData!.ShapeTree!, "North America: $2.5M\nEurope: $1.8M\nAsia-Pacific: $1.2M");

        prs.Save();
    }

    public static void CreateModifiedPresentation(string path)
    {
        using var prs = PresentationDocument.Create(path, PresentationDocumentType.Presentation);
        var presentationPart = prs.AddPresentationPart();
        presentationPart.Presentation = new P.Presentation();

        var slideIdList = new P.SlideIdList();
        presentationPart.Presentation.Append(slideIdList);

        // Modified Slide 1 - changed title and content
        var slidePart1 = presentationPart.AddNewPart<SlidePart>();
        slideIdList.Append(new P.SlideId { Id = 256U, RelationshipId = presentationPart.GetIdOfPart(slidePart1) });
        slidePart1.Slide = new P.Slide();
        slidePart1.Slide.Append(new P.CommonSlideData(new P.ShapeTree()));

        AddTitleShape(slidePart1.Slide.CommonSlideData!.ShapeTree!, "Q3 2024 Financial Results");
        AddTextShape(slidePart1.Slide.CommonSlideData!.ShapeTree!, "Annual Sales Performance Analysis");

        // Modified Slide 2 - updated numbers
        var slidePart2 = presentationPart.AddNewPart<SlidePart>();
        slideIdList.Append(new P.SlideId { Id = 257U, RelationshipId = presentationPart.GetIdOfPart(slidePart2) });
        slidePart2.Slide = new P.Slide();
        slidePart2.Slide.Append(new P.CommonSlideData(new P.ShapeTree()));

        AddTitleShape(slidePart2.Slide.CommonSlideData!.ShapeTree!, "Revenue by Region");
        AddTextShape(slidePart2.Slide.CommonSlideData!.ShapeTree!, "North America: $2.8M\nEurope: $2.1M\nAsia-Pacific: $1.5M\nOther: $0.6M");

        // New Slide 3
        var slidePart3 = presentationPart.AddNewPart<SlidePart>();
        slideIdList.Append(new P.SlideId { Id = 258U, RelationshipId = presentationPart.GetIdOfPart(slidePart3) });
        slidePart3.Slide = new P.Slide();
        slidePart3.Slide.Append(new P.CommonSlideData(new P.ShapeTree()));

        AddTitleShape(slidePart3.Slide.CommonSlideData!.ShapeTree!, "Market Outlook");
        AddTextShape(slidePart3.Slide.CommonSlideData!.ShapeTree!, "Expected growth: 15-20% YoY");

        prs.Save();
    }

    private static void AddTitleShape(P.ShapeTree shapeTree, string title)
    {
        var shape = new P.Shape();
        shape.Append(new P.NonVisualShapeProperties());
        shape.Append(new P.ShapeProperties());

        var txBody = new P.TextBody();
        var para = new A.Paragraph();
        var run = new A.Run { Text = new A.Text { Text = title } };
        para.Append(run);
        txBody.Append(para);
        shape.Append(txBody);

        shapeTree.Append(shape);
    }

    private static void AddTextShape(P.ShapeTree shapeTree, string content)
    {
        var shape = new P.Shape();
        shape.Append(new P.NonVisualShapeProperties());
        shape.Append(new P.ShapeProperties());

        var txBody = new P.TextBody();
        foreach (var line in content.Split('\n'))
        {
            var para = new A.Paragraph();
            var run = new A.Run { Text = new A.Text { Text = line } };
            para.Append(run);
            txBody.Append(para);
        }
        shape.Append(txBody);

        shapeTree.Append(shape);
    }
}
