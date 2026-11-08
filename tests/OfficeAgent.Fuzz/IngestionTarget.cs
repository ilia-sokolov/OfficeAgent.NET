using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Excel;
using OfficeAgent.PowerPoint;
using OfficeAgent.Word;

namespace OfficeAgent.Fuzz;

/// <summary>What one input produced when it was inspected.</summary>
internal enum Outcome
{
    /// <summary>The bytes were a readable document and inspection returned a result.</summary>
    Result,

    /// <summary>Refused with <see cref="OpenXmlIngestionLimitException"/> (<c>input-too-large</c>).</summary>
    LimitRefusal,

    /// <summary>Refused with <see cref="OpenXmlPackageRejectedException"/> (<c>malformed-package</c>).</summary>
    PackageRejected,

    /// <summary>Refused with <see cref="EncryptedDocumentException"/> (<c>encrypted-document</c>).</summary>
    Encrypted,

    /// <summary>
    /// An exception matching an entry in <c>OFFICEAGENT_FUZZ_KNOWN</c>: a finding already reported,
    /// set aside so the campaign can look for others.
    /// </summary>
    Known
}

/// <summary>
/// Raised when an input breaks a documented promise without throwing on its own: a refusal
/// whose diagnostic is not short, or an inspection that allocated past the bound.
/// </summary>
internal sealed class FuzzFindingException : Exception
{
    public FuzzFindingException(string message, Exception? inner = null) : base(message, inner)
    {
    }
}

/// <summary>
/// The fuzz target: arbitrary bytes, wrapped in a <see cref="StreamHandle"/>, inspected through a
/// client with the Word, PowerPoint and Excel modules registered.
/// </summary>
/// <remarks>
/// <para>
/// The expected outcomes are a result or one of the three documented refusals in
/// docs/ingestion-limits.md. Anything else escapes this method and is a finding: any other
/// exception type, including an <see cref="OutOfMemoryException"/> past the managed heap limit set
/// in the project file. Two more findings are raised here because nothing would otherwise throw:
/// a refusal diagnostic of 600 characters or more, the bound <c>BoundedIngestionTests</c> holds
/// every refusal to, and an inspection that allocates more than
/// <see cref="MaximumAllocatedBytes"/>. A hang is a finding too; libFuzzer's <c>-timeout</c>
/// reports it, since a hung call cannot report itself.
/// </para>
/// <para>
/// The client uses the default <see cref="OpenXmlIngestionLimits"/>, which are what a host gets
/// when it configures nothing.
/// </para>
/// </remarks>
internal static class IngestionTarget
{
    /// <summary>The refusal diagnostic bound that <c>BoundedIngestionTests</c> asserts.</summary>
    public const int MaximumDiagnosticLength = 600;

    private static readonly OfficeAgentClient Client =
        new(new WordModule(), new PowerPointModule(), new ExcelModule());

    /// <summary>
    /// Gets the most an inspection of one input may allocate, in bytes, over its lifetime.
    /// </summary>
    /// <remarks>
    /// libfuzzer-dotnet passes at most 1 MiB per input, and the default expansion ceiling of 200
    /// bounds what that can expand to at about 200 MiB. Allocating more than 1 GiB to inspect it
    /// means the work was not bounded by the ceilings. Set
    /// <c>OFFICEAGENT_FUZZ_MAX_ALLOCATED_MB</c> to change it.
    /// </remarks>
    public static long MaximumAllocatedBytes { get; } =
        long.TryParse(Environment.GetEnvironmentVariable("OFFICEAGENT_FUZZ_MAX_ALLOCATED_MB"), out var mb) && mb > 0
            ? mb * 1024 * 1024
            : 1024L * 1024 * 1024;

    /// <summary>
    /// Findings already reported, from <c>OFFICEAGENT_FUZZ_KNOWN</c>: entries separated by <c>;</c>,
    /// each an exception type's full name, <c>@</c>, and text its stack trace must contain, such as
    /// <c>System.ArgumentOutOfRangeException@OpenXmlPartContainer.GetPartById</c>. Empty unless set.
    /// </summary>
    private static readonly (string Type, string Frame)[] KnownFindings =
        (Environment.GetEnvironmentVariable("OFFICEAGENT_FUZZ_KNOWN") ?? string.Empty)
            .Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(entry => entry.Split('@', 2))
            .Where(parts => parts.Length == 2)
            .Select(parts => (parts[0], parts[1]))
            .ToArray();

    /// <summary>Gets the bytes allocated by the most recent call to <see cref="Run"/>.</summary>
    public static long LastAllocatedBytes { get; private set; }

    /// <summary>Runs one input. Returns its outcome or throws a finding.</summary>
    public static Outcome Run(ReadOnlySpan<byte> data)
    {
        var input = data.ToArray();
        long before = GC.GetTotalAllocatedBytes(precise: false);

        Outcome outcome;
        try
        {
            using var stream = new MemoryStream(input, writable: false);
            var result = Client.InspectAsync(new StreamHandle(stream, "fuzz-input"))
                .GetAwaiter().GetResult();
            if (result is null) throw new FuzzFindingException("InspectAsync returned null.");
            outcome = Outcome.Result;
        }
        catch (OpenXmlIngestionLimitException ex)
        {
            Bounded(ex);
            outcome = Outcome.LimitRefusal;
        }
        catch (OpenXmlPackageRejectedException ex)
        {
            Bounded(ex);
            outcome = Outcome.PackageRejected;
        }
        catch (EncryptedDocumentException ex)
        {
            Bounded(ex);
            outcome = Outcome.Encrypted;
        }
        catch (Exception ex) when (IsKnown(ex))
        {
            outcome = Outcome.Known;
        }

        LastAllocatedBytes = GC.GetTotalAllocatedBytes(precise: false) - before;
        if (LastAllocatedBytes > MaximumAllocatedBytes)
            throw new FuzzFindingException(
                $"Inspecting a {input.Length}-byte input allocated {LastAllocatedBytes} bytes, past the " +
                $"{MaximumAllocatedBytes}-byte bound ({outcome}).");

        return outcome;
    }

    private static bool IsKnown(Exception ex) =>
        ex is not FuzzFindingException &&
        KnownFindings.Any(known =>
            ex.GetType().FullName == known.Type &&
            (ex.StackTrace ?? string.Empty).Contains(known.Frame, StringComparison.Ordinal));

    private static void Bounded(Exception refusal)
    {
        if (refusal.Message.Length >= MaximumDiagnosticLength)
            throw new FuzzFindingException(
                $"{refusal.GetType().Name} diagnostic is {refusal.Message.Length} characters; refusals are " +
                $"documented to stay short (under {MaximumDiagnosticLength}).", refusal);
    }
}
