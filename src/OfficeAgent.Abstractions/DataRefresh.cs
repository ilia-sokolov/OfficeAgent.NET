namespace OfficeAgent.Abstractions;

/// <summary>
/// How a Word document or PowerPoint deck records that part of it shows workbook data: a name
/// or tag that starts with <see cref="Prefix"/> followed by the range, such as
/// <c>xl:Summary!B2</c>. The marker lives where Office keeps names - a content control's tag, a
/// Word table's alt-text title, a PowerPoint shape's name - so it survives editing and saving in
/// Office, and the same document can be refreshed again next month.
/// </summary>
public static class DataBindingMarker
{
    /// <summary>The prefix every binding marker starts with.</summary>
    public const string Prefix = "xl:";

    /// <summary>
    /// Reads the workbook reference out of a name or tag, or returns <see langword="false"/> when
    /// it is not a binding. A PowerPoint name qualified with its slide, <c>slide256/xl:…</c>, is
    /// read too. Anything after <c>#</c> is ignored: it tells apart two places bound to the same
    /// range, because Word content-control tags must be unique.
    /// </summary>
    public static bool TryParse(string? name, out string reference)
    {
        reference = string.Empty;
        if (string.IsNullOrEmpty(name)) return false;

        var start = name!.StartsWith(Prefix, StringComparison.OrdinalIgnoreCase)
            ? 0
            : name.IndexOf("/" + Prefix, StringComparison.OrdinalIgnoreCase) is var slash and >= 0 ? slash + 1 : -1;
        if (start < 0) return false;

        var value = name.Substring(start + Prefix.Length);
        var hash = value.IndexOf('#');
        if (hash >= 0) value = value.Substring(0, hash);
        reference = value.Trim();
        return reference.Length > 0;
    }

    /// <summary>The marker for a workbook reference.</summary>
    public static string For(string reference) => Prefix + reference;
}

/// <summary>What a data binding updates.</summary>
public enum DataBindingKind
{
    /// <summary>One value: a Word content control or a PowerPoint text shape, filled with one cell's displayed value.</summary>
    Value,

    /// <summary>A table, filled cell by cell from a range.</summary>
    Table,

    /// <summary>A native chart, whose data is replaced from a range with series in columns.</summary>
    Chart
}

/// <summary>One binding found in a document, and what refreshing it will write.</summary>
public sealed class DataRefreshBinding
{
    /// <summary>Gets what the binding updates.</summary>
    public DataBindingKind Kind { get; init; }

    /// <summary>Gets the marker as the document carries it, such as <c>xl:Summary!B2</c>.</summary>
    public string Marker { get; init; } = string.Empty;

    /// <summary>Gets the workbook reference the marker names.</summary>
    public string Reference { get; init; } = string.Empty;

    /// <summary>Gets the part of the document the binding updates.</summary>
    public Anchor Target { get; init; } = null!;

    /// <summary>Gets the range the reference resolved to, as <c>Sheet!A1:D6</c>, or <see langword="null"/> when it did not resolve.</summary>
    public string? Range { get; init; }

    /// <summary>
    /// Gets what will be written: the displayed value for a value binding, or the shape of the
    /// data - rows and columns for a table, categories and series for a chart.
    /// </summary>
    public string? Summary { get; init; }

    /// <summary>Gets why the binding cannot be refreshed, from <see cref="DataRefreshErrorCodes"/>, or <see langword="null"/>.</summary>
    public string? ErrorCode { get; init; }

    /// <summary>Gets a message explaining <see cref="ErrorCode"/>.</summary>
    public string? Message { get; init; }
}

/// <summary>
/// A data refresh worked out but not applied: every binding the document carries, what each one
/// will write, and the plan that writes it. Apply <see cref="Plan"/> like any other plan, to
/// preview it or to save it with a receipt.
/// </summary>
/// <remarks>
/// A refresh is all or nothing. When any binding cannot be read, <see cref="IsComplete"/> is
/// <see langword="false"/> and <see cref="Plan"/> has no operations, because a report showing
/// this month's revenue beside last month's costs is worse than one left as it was.
/// </remarks>
public sealed class DataRefreshPlan
{
    /// <summary>Gets the plan that writes every binding, or an empty plan when the refresh is incomplete.</summary>
    public DocumentPlan Plan { get; init; } = new();

    /// <summary>Gets every binding found, in document order.</summary>
    public IReadOnlyList<DataRefreshBinding> Bindings { get; init; } = Array.Empty<DataRefreshBinding>();

    /// <summary>Gets whether every binding resolved, so <see cref="Plan"/> refreshes the whole document.</summary>
    public bool IsComplete => Bindings.Count > 0 && Bindings.All(binding => binding.ErrorCode is null);
}
