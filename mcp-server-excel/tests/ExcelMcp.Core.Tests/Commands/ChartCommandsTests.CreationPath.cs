using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands;

/// <summary>
/// Integration tests for the chart-creation paths.
///
/// Not every chart type can be built the same way: the Excel 2016+ charts reject
/// Chart.SetSourceData (0x800A03EC) and must be filled series by series, while the stock charts
/// cannot be created directly at all. These tests pin both facts against real Excel, including
/// the read-back check that catches a silent fallback to a different chart type.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "Core")]
[Trait("Feature", "Charts")]
[Trait("RequiresExcel", "true")]
public partial class ChartCommandsTests
{
    [Theory]
    [InlineData("A1:C6", ChartType.Treemap)]
    [InlineData("A1:C6", ChartType.Histogram)]
    [InlineData("A1:C6", ChartType.Pareto)]
    [InlineData("A1:C6", ChartType.BoxWhisker)]
    [InlineData("A1:C6", ChartType.Waterfall)]
    [InlineData("A1:C6", ChartType.Funnel)]
    [InlineData("A1:C6", ChartType.RegionMap)]
    [InlineData("A1:C6", ChartType.ColumnLineCombo)]
    // Stock charts reject the creation when the series count does not match the variant
    // exactly (HLC=3, OHLC=4, VHLC=4, VOHLC=5 numeric columns — no category column).
    [InlineData("A1:C6", ChartType.StockHLC)]
    [InlineData("A1:D6", ChartType.StockOHLC)]
    [InlineData("A1:D6", ChartType.StockVHLC)]
    [InlineData("A1:E6", ChartType.StockVOHLC)]
    public void CreateFromRange_TypeWithItsOwnCreationPath_BuildsTheRequestedType(
        string sourceRange,
        ChartType chartType)
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.SharedTestFile);

        // Act - the creation path differs per type (series attach / stock switch)
        var createResult = _commands.CreateFromRange(
            batch, "Sheet1", sourceRange, chartType, 50, 50, 400, 300);

        // Assert - Excel must report the type that was asked for, never a silent fallback
        Assert.Equal(chartType, createResult.ChartType);

        var readResult = _commands.Read(batch, createResult.ChartName);
        Assert.Equal(chartType, readResult.ChartType);
    }

    [Fact]
    public void AddSeries_UnqualifiedAddress_ResolvesAgainstTheChartsWorksheet()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.SharedTestFile);
        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:B6", ChartType.Line, 50, 50);
        var readBefore = _commands.Read(batch, createResult.ChartName);

        // Act - "C2:C6" carries no sheet name, which is what a caller sends when the data sits
        // on the chart's own sheet. Series.Values needs a Range object, not the address string.
        var addResult = _commands.AddSeries(
            batch, createResult.ChartName, "Series4", "C2:C6", "A2:A6");

        // Assert
        Assert.Equal("Series4", addResult.Name);
        var readAfter = _commands.Read(batch, createResult.ChartName);
        Assert.Equal(readBefore.Series.Count + 1, readAfter.Series.Count);
        Assert.Contains(readAfter.Series, series => series.Name == "Series4");
    }
}
