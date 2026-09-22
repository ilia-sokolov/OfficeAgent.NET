namespace OfficeAgent.Tests;

/// <summary>
/// Keeps a runtime CI leg honest. The .NET 10 leg runs these net8.0 tests with
/// <c>DOTNET_ROLL_FORWARD=LatestMajor</c>; if the roll-forward did not happen, the suite would
/// pass on .NET 8 and the leg would claim coverage it never had.
/// </summary>
public sealed class RuntimeLegTests
{
    /// <summary>
    /// When <c>OFFICEAGENT_EXPECT_RUNTIME_MAJOR</c> is set, the tests must be running on that
    /// major version of .NET. Unset, as in the ordinary build, there is nothing to check.
    /// </summary>
    [Fact]
    public void Tests_run_on_the_runtime_the_leg_asks_for()
    {
        var expected = Environment.GetEnvironmentVariable("OFFICEAGENT_EXPECT_RUNTIME_MAJOR");
        if (string.IsNullOrWhiteSpace(expected)) return;

        Assert.Equal(int.Parse(expected, System.Globalization.CultureInfo.InvariantCulture), Environment.Version.Major);
    }
}
