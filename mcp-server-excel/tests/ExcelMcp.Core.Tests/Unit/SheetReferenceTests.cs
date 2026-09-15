using Sbroenne.ExcelMcp.Core.Utilities;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "Chart")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class SheetReferenceTests
{
    /// <summary>
    /// The regression this pins down: a worksheet named with an apostrophe was interpolated raw,
    /// producing 'O'Brien'!A1:D20, which Excel rejects with a bare 0x800A03EC. Every legal sheet
    /// name must survive quoting.
    /// </summary>
    [Theory]
    [InlineData("Plain", "A1:B2", "'Plain'!A1:B2")]
    [InlineData("With Space", "A1:B2", "'With Space'!A1:B2")]
    [InlineData("O'Brien", "A1:C7", "'O''Brien'!A1:C7")]
    [InlineData("'Quoted'", "A1", "'''Quoted'''!A1")]
    [InlineData("a'b'c", "$A$1:$B$2", "'a''b''c'!$A$1:$B$2")]
    public void BuildRangeReference_EscapesApostrophes(string sheetName, string address, string expected)
    {
        Assert.Equal(expected, SheetReference.BuildRangeReference(sheetName, address));
    }

    [Fact]
    public void QuoteSheetName_LeavesNamesWithoutApostrophesUntouched()
    {
        Assert.Equal("Sales Data 2026", SheetReference.QuoteSheetName("Sales Data 2026"));
    }

    [Fact]
    public void QuoteSheetName_DoublesEveryApostrophe()
    {
        Assert.Equal("''''", SheetReference.QuoteSheetName("''"));
    }
}
