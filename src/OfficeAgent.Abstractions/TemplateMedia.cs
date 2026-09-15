namespace OfficeAgent.Abstractions;

/// <summary>
/// A value bound to a template slot. Text, an image, or a native chart.
/// </summary>
/// <remarks>
/// Typed bindings sit beside the existing string <see cref="TemplateBinding.Values"/>
/// rather than replacing them, so a scalar-only request keeps behaving exactly as it did.
/// A typed value aimed at a slot that cannot hold it is an error, never a silent
/// stringification and never a quietly dropped field.
/// </remarks>
public abstract class TemplateValue
{
    /// <summary>Gets the slot kind this value can be bound to.</summary>
    public abstract string RequiredSlotKind { get; }
}

/// <summary>Text bound to a content control or a named shape.</summary>
public sealed class TemplateTextValue : TemplateValue
{
    /// <summary>Initializes a text value.</summary>
    public TemplateTextValue(string? text) => Text = text;

    /// <summary>Gets the text to place in the slot.</summary>
    public string? Text { get; }

    /// <inheritdoc />
    public override string RequiredSlotKind => "text";
}

/// <summary>
/// An image bound to a discovered image slot.
/// </summary>
/// <remarks>
/// The bytes come from one of two places and never from the open internet: inline
/// base64 the caller already holds, or a document the caller previously put in a
/// provider connection they are authorized for. There is no URL form, because fetching
/// one would make the engine a client of whatever address a plan happened to carry.
/// </remarks>
public sealed class TemplateImageValue : TemplateValue
{
    /// <summary>Gets inline image bytes. Mutually exclusive with the provider reference.</summary>
    public string? Base64Bytes { get; init; }

    /// <summary>Gets the connection holding the image. Required with <see cref="ImageDocumentId"/>.</summary>
    public string? ImageConnectionId { get; init; }

    /// <summary>Gets the provider-assigned image id. Mutually exclusive with inline bytes.</summary>
    public string? ImageDocumentId { get; init; }

    /// <summary>Gets the image format: <c>png</c>, <c>jpeg</c>, <c>gif</c>, <c>bmp</c>, or <c>tiff</c>.</summary>
    public string ImageType { get; init; } = "png";

    /// <summary>Gets the display width in pixels at 96 DPI.</summary>
    public int WidthPx { get; init; } = 200;

    /// <summary>Gets the display height in pixels at 96 DPI.</summary>
    public int HeightPx { get; init; } = 200;

    /// <summary>
    /// Gets the alternative text describing the image. Required: an image with no alt
    /// text is inaccessible to anyone reading with a screen reader, and a template that
    /// generates many documents multiplies that.
    /// </summary>
    public string? AltText { get; init; }

    /// <inheritdoc />
    public override string RequiredSlotKind => "image";
}

/// <summary>
/// Data bound to a chart that already exists in the template.
/// </summary>
/// <remarks>
/// This updates a native chart's series and categories; it does not create one, and it
/// is PowerPoint only, because that is where this engine has native chart handling. A
/// chart binding aimed at a Word template is refused rather than ignored.
/// </remarks>
public sealed class TemplateChartValue : TemplateValue
{
    /// <summary>Gets the chart type to apply.</summary>
    public ChartKind Kind { get; init; } = ChartKind.ClusteredColumn;

    /// <summary>Gets the category labels.</summary>
    public IReadOnlyList<string> Categories { get; init; } = Array.Empty<string>();

    /// <summary>Gets the data series.</summary>
    public IReadOnlyList<ChartSeries> Series { get; init; } = Array.Empty<ChartSeries>();

    /// <summary>Gets an optional chart title.</summary>
    public string? Title { get; init; }

    /// <summary>Gets whether the legend is shown.</summary>
    public bool ShowLegend { get; init; } = true;

    /// <inheritdoc />
    public override string RequiredSlotKind => "chart";
}

/// <summary>A slot that can hold something other than text.</summary>
public sealed class TemplateMediaSlot
{
    /// <summary>Gets the name to bind by.</summary>
    public string Name { get; init; } = string.Empty;

    /// <summary>Gets what the slot holds: <c>image</c> or <c>chart</c>.</summary>
    public string MediaKind { get; init; } = string.Empty;

    /// <summary>Gets the anchor path or tag the binding resolves to.</summary>
    public string Path { get; init; } = string.Empty;

    /// <summary>
    /// Gets whether this slot sits inside a repeating table row. Media in a repeating row
    /// is not supported, and a binding that targets one is refused rather than producing
    /// a document with one image shared by every row.
    /// </summary>
    public bool InRepeatingRow { get; init; }
}

/// <summary>Host budgets for media carried by a template batch.</summary>
public sealed class TemplateMediaLimits
{
    /// <summary>The budgets applied when a host configures nothing.</summary>
    public static TemplateMediaLimits Default { get; } = new();

    /// <summary>Gets the maximum decoded size of one image, in bytes.</summary>
    public long MaximumImageBytes { get; init; } = 8L * 1024 * 1024;

    /// <summary>Gets the maximum images one output may carry.</summary>
    public int MaximumImagesPerDocument { get; init; } = 20;

    /// <summary>Gets the maximum decoded image bytes the whole batch may carry.</summary>
    public long MaximumTotalImageBytes { get; init; } = 64L * 1024 * 1024;

    /// <summary>Gets the maximum data points one chart binding may carry.</summary>
    public int MaximumChartPoints { get; init; } = 1_000;

    /// <summary>Returns the stricter of these budgets and a request's own.</summary>
    public TemplateMediaLimits Restrict(TemplateMediaLimits? requested)
    {
        if (requested is null) return this;

        return new TemplateMediaLimits
        {
            MaximumImageBytes = Math.Min(MaximumImageBytes, requested.MaximumImageBytes),
            MaximumImagesPerDocument =
                Math.Min(MaximumImagesPerDocument, requested.MaximumImagesPerDocument),
            MaximumTotalImageBytes = Math.Min(MaximumTotalImageBytes, requested.MaximumTotalImageBytes),
            MaximumChartPoints = Math.Min(MaximumChartPoints, requested.MaximumChartPoints)
        };
    }

    /// <summary>Throws when any budget is not positive.</summary>
    public void Validate()
    {
        if (MaximumImageBytes <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaximumImageBytes));
        if (MaximumImagesPerDocument <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaximumImagesPerDocument));
        if (MaximumTotalImageBytes <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaximumTotalImageBytes));
        if (MaximumChartPoints <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaximumChartPoints));
    }
}
