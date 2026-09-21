namespace OfficeAgent.Tests;

/// <summary>
/// Pins what the public API baseline records, so the compatibility gate cannot quietly lose
/// precision again.
/// </summary>
/// <remarks>
/// The baseline in docs/csharp-api.md is compared verbatim, so a change is caught only if the
/// renderer puts it into the text. The first renderer did not: probed with five deliberate
/// breaking changes against a passing control, it missed all five. Each test here fixes one of
/// those classes, plus the shapes that break a subclass or an implementer. A test that renders a
/// fixture is fast and deterministic, where the original probe had to edit source and rebuild.
/// </remarks>
public sealed class ApiSurfaceFidelityTests
{
    private static IReadOnlyList<string> Render(Type type) => ApiSurface.RenderType(type);

    [Fact]
    public void Constant_values_are_part_of_the_baseline()
    {
        // A const is copied into every caller at compile time; changing it changes what
        // already-compiled code believes the value is.
        Assert.Contains("public const string Code = \"fixture-code\";", Render(typeof(FidelityFixture)));
        Assert.Contains("public const int Limit = 42;", Render(typeof(FidelityFixture)));
    }

    [Fact]
    public void Enum_members_carry_their_numeric_values()
    {
        // Renumbering keeps every name and order, and still breaks every caller that stored
        // or serialised the number.
        var lines = Render(typeof(FidelityEnum));
        Assert.Contains("First = 10", lines);
        Assert.Contains("Second = 20", lines);
    }

    [Fact]
    public void Non_int_enum_underlying_types_are_recorded()
    {
        Assert.Contains("// underlying type: byte", Render(typeof(FidelityByteEnum)));
    }

    [Fact]
    public void Default_arguments_are_part_of_the_baseline()
    {
        // The default is compiled into each call site that omits the argument.
        var method = Assert.Single(Render(typeof(FidelityFixture)), line => line.Contains("Compare("));
        Assert.Contains("string author = \"OfficeAgent Compare\"", method);
        Assert.Contains("int attempts = 3", method);
        Assert.Contains("CancellationToken cancellationToken = default", method);
        Assert.Contains("FidelityEnum kind = FidelityEnum.Second", method);
    }

    [Fact]
    public void Init_only_and_settable_properties_are_distinguished()
    {
        var lines = Render(typeof(FidelityFixture));
        Assert.Contains("public string Settable { get; set; }", lines);
        Assert.Contains("public string InitOnly { get; init; }", lines);
        Assert.Contains("public string ReadOnly { get; }", lines);
    }

    [Fact]
    public void Reference_nullability_is_part_of_the_baseline()
    {
        var lines = Render(typeof(FidelityFixture));
        Assert.Contains("public string? Maybe { get; init; }", lines);
        Assert.Contains("public string InitOnly { get; init; }", lines);
        Assert.Contains("public IReadOnlyList<string?> MaybeItems { get; init; }", lines);
        Assert.Single(lines, line => line.Contains("string? Find(string key)"));
    }

    [Fact]
    public void Base_types_and_interfaces_are_part_of_the_declaration()
    {
        var declaration = Render(typeof(FidelityDerived))[0];
        Assert.Contains(": OfficeAgent.Tests.FidelityBase", declaration);
        Assert.Contains("System.IDisposable", declaration);
    }

    [Fact]
    public void Protected_members_of_inheritable_types_are_visible_and_modifiers_are_recorded()
    {
        // A subclass depends on protected members and on which methods it may override.
        var baseLines = Render(typeof(FidelityBase));
        Assert.Contains("protected FidelityBase();", baseLines);
        Assert.Contains("protected abstract string Name();", baseLines);
        Assert.Contains("public virtual int Size();", baseLines);

        // An override on a type that can itself be derived from stays part of the surface.
        Assert.Contains("protected override string Name();", Render(typeof(FidelityMiddle)));
        Assert.Contains("public sealed override int Size();", Render(typeof(FidelityDerived)));
    }

    [Fact]
    public void Protected_members_of_sealed_types_are_not_surface()
    {
        // Nobody can derive from a sealed type, so its protected members are unreachable.
        Assert.DoesNotContain(Render(typeof(FidelityDerived)), line => line.StartsWith("protected FidelityDerived"));
    }

    [Fact]
    public void Params_arrays_and_generic_constraints_are_recorded()
    {
        var lines = Render(typeof(FidelityFixture));
        Assert.Single(lines, line => line.Contains("Join(params string[] parts)"));
        Assert.Single(lines, line => line.Contains("Make<T>()") && line.EndsWith(" where T : class, new();"));
    }

    [Fact]
    public void Stability_class_is_part_of_the_baseline()
    {
        // Dropping the attribute promises 1.x stability for a seam; adding it withdraws a
        // promise from every host built on the type. Either must show up as a diff.
        Assert.Equal("[Experimental(\"OFFICEAGENT001\")]", Render(typeof(FidelityExperimental))[0]);
        Assert.DoesNotContain(Render(typeof(FidelityFixture)), line => line.StartsWith("[Experimental", StringComparison.Ordinal));
    }

    [Fact]
    public void The_committed_baseline_records_what_a_caller_depends_on()
    {
        // The real surface, not a fixture: the members the original probes changed must be
        // rendered with the detail that made each probe detectable.
        var baseline = File.ReadAllText(Path.Combine(RepositoryRoot(), "docs", "csharp-api.md"));
        Assert.Contains("public string Id { get; init; }", baseline);
        Assert.Contains("public const string Code = \"input-too-large\";", baseline);
        Assert.Contains("Previewed = 0", baseline);
        Assert.Contains("string revisionAuthor = \"OfficeAgent Compare\"", baseline);
        Assert.Contains("public string ConnectionId { get; set; }", baseline);
    }

    private static string RepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "OfficeAgent.NET.sln")))
            directory = directory.Parent;
        return directory?.FullName ?? throw new DirectoryNotFoundException("repository root not found");
    }
}

#pragma warning disable CS1591 // Fixture types exist only to be rendered.
public enum FidelityEnum
{
    First = 10,
    Second = 20
}

public enum FidelityByteEnum : byte
{
    Only = 1
}

public sealed class FidelityFixture
{
    public const string Code = "fixture-code";
    public const int Limit = 42;

    public string Settable { get; set; } = string.Empty;
    public string InitOnly { get; init; } = string.Empty;
    public string ReadOnly { get; } = string.Empty;
    public string? Maybe { get; init; }
    public IReadOnlyList<string?> MaybeItems { get; init; } = Array.Empty<string?>();

    public string? Find(string key) => key.Length == 0 ? null : key;

    public Task Compare(
        string author = "OfficeAgent Compare",
        int attempts = 3,
        FidelityEnum kind = FidelityEnum.Second,
        CancellationToken cancellationToken = default) => Task.CompletedTask;

    public string Join(params string[] parts) => string.Concat(parts);

    public T Make<T>() where T : class, new() => new();
}

public abstract class FidelityBase
{
    protected FidelityBase()
    {
    }

    protected abstract string Name();

    public virtual int Size() => 0;
}

public class FidelityMiddle : FidelityBase
{
    protected override string Name() => "middle";
}

[System.Diagnostics.CodeAnalysis.Experimental("OFFICEAGENT001")]
public interface FidelityExperimental
{
}

public sealed class FidelityDerived : FidelityBase, IDisposable
{
    protected override string Name() => "derived";

    public sealed override int Size() => 1;

    public void Dispose()
    {
    }
}
#pragma warning restore CS1591
