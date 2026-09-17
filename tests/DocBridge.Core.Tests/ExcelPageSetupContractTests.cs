using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class ExcelPageSetupContractTests
{
    [Theory]
    [InlineData("7:7", false, "$7:$7")]
    [InlineData("$7:$7", false, "$7:$7")]
    [InlineData("1:3", false, "$1:$3")]
    [InlineData("A:B", true, "$A:$B")]
    [InlineData("$a:$c", true, "$A:$C")]
    public void Title_ranges_are_absolutized_before_com_assign(string input, bool columns, string expected) =>
        Assert.Equal(expected, ExcelPageSetupContract.AbsolutizeTitleRange(input, columns));

    [Theory]
    [InlineData("7", false)]
    [InlineData("0:3", false)]
    [InlineData("3:1", false)]
    [InlineData("A:1", false)]
    [InlineData("1:A", true)]
    public void Invalid_title_ranges_pass_through_for_validator_rejection(string input, bool columns) =>
        Assert.Equal(input, ExcelPageSetupContract.AbsolutizeTitleRange(input, columns));
}
