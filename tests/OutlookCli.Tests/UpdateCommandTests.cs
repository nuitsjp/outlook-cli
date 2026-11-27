using OutlookCli.Commands;
using Xunit;

namespace OutlookCli.Tests;

public sealed class UpdateCommandTests
{
    [Theory]
    [InlineData("y\n", true)]
    [InlineData("Y\n", true)]
    [InlineData("\n", true)]
    [InlineData("n\n", false)]
    [InlineData("unexpected\n", false)]
    [InlineData("", false)]
    public void ConfirmUpdate_displays_version_and_requires_an_affirmative_reply(string input, bool confirmed)
    {
        using var reader = new StringReader(input);
        using var error = new StringWriter();

        var result = UpdateCommand.ConfirmUpdate(reader, error, "1.2.3", "1.2.4");

        Assert.Equal(confirmed, result);
        Assert.Contains("1.2.3", error.ToString(), StringComparison.Ordinal);
        Assert.Contains("1.2.4", error.ToString(), StringComparison.Ordinal);
        Assert.Contains("[Y/n]", error.ToString(), StringComparison.Ordinal);
    }
}
