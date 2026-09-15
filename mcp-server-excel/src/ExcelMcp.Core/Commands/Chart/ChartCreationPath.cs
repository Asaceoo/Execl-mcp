using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;

namespace Sbroenne.ExcelMcp.Core.Commands.Chart;

/// <summary>
/// Chooses how a chart is created and how its data is attached.
///
/// One path does not fit every chart type (verified against Excel 16.0 build 20326):
///
/// 1. Default - <c>Shapes.AddChart(type)</c> then <c>Chart.SetSourceData(range)</c>.
///    Works for every classic type (column/bar/line/pie/area/scatter/radar/surface/bubble/3D/cylinder/cone/pyramid).
///
/// 2. Series attach - <c>Shapes.AddChart2(-1, type)</c> then one <c>SeriesCollection.NewSeries()</c>
///    per data column. The Excel 2016+ charts (treemap/sunburst/histogram/pareto/box-whisker/
///    waterfall/funnel/region map/combo) reject <c>SetSourceData</c> with 0x800A03EC even though
///    the chart itself is created fine, so the data has to be attached series by series.
///
/// 3. Stock switch - create a clustered column chart, attach the source, then set <c>ChartType</c>.
///    The four stock charts (HLC/OHLC/VHLC/VOHLC) cannot be created directly: AddChart fails with
///    0x800A03EC while this order succeeds.
/// </summary>
internal static class ChartCreationPath
{
    /// <summary>xlColumnClustered - the scaffold chart used for stock charts.</summary>
    private const int ColumnClusteredCode = 51;

    /// <summary>xlColumns - explicit plot orientation for the stock scaffold's SetSourceData.</summary>
    private const int PlotByColumns = 2;

    /// <summary>Chart types that require the series-attach path (Excel 2016+ charts).</summary>
    private static readonly HashSet<int> SeriesAttachCodes = new()
    {
        116, // Sunburst
        117, // Treemap
        118, // Histogram
        119, // Waterfall
        120, // ColumnLineCombo
        121, // BoxWhisker
        122, // Pareto
        123, // Funnel
        140  // RegionMap
    };

    /// <summary>Chart types that require the stock-switch path.</summary>
    private static readonly HashSet<int> StockCodes = new() { 88, 89, 90, 91 };

    /// <summary>True when the type needs its data attached series by series.</summary>
    internal static bool NeedsSeriesAttach(ChartType chartType) =>
        SeriesAttachCodes.Contains((int)chartType);

    /// <summary>True when the type must be created as a column chart and switched afterwards.</summary>
    internal static bool IsStockChart(ChartType chartType) =>
        StockCodes.Contains((int)chartType);

    /// <summary>
    /// Creates the chart shape and attaches <paramref name="sourceRange"/> using the path that
    /// matches <paramref name="chartType"/>. Caller owns the returned shape.
    /// </summary>
    internal static dynamic CreateShape(
        dynamic shapes,
        ChartType chartType,
        double left,
        double top,
        double width,
        double height,
        dynamic sourceRange)
    {
        int code = (int)chartType;

        if (StockCodes.Contains(code))
        {
            dynamic stockShape = shapes.AddChart(ColumnClusteredCode, left, top, width, height);
            try
            {
                dynamic scaffoldChart = stockShape.Chart;
                // PlotBy must be explicit: with it omitted, AddChart auto-fills the scaffold from
                // the sheet's whole current region first, and SetSourceData(range) then leaves that
                // region's series count in place — the ChartType switch below rejects stock charts
                // whose series count does not match the variant (HLC=3, OHLC/VHLC=4, VOHLC=5).
                // Passing xlColumns makes SetSourceData authoritative, so the requested range wins.
                scaffoldChart.SetSourceData(sourceRange, PlotByColumns);
                scaffoldChart.ChartType = code;
                ComUtilities.Release(ref scaffoldChart!);
                return stockShape;
            }
            catch
            {
                TryDeleteShape(stockShape);
                throw;
            }
        }

        if (SeriesAttachCodes.Contains(code))
        {
            dynamic modernShape = shapes.AddChart2(-1, code, left, top, width, height);
            try
            {
                dynamic modernChart = modernShape.Chart;
                AttachSeriesFromRange(modernChart, sourceRange, code);
                ComUtilities.Release(ref modernChart!);
                return modernShape;
            }
            catch
            {
                TryDeleteShape(modernShape);
                throw;
            }
        }

        dynamic shape = shapes.AddChart(code, left, top, width, height);
        try
        {
            dynamic chart = shape.Chart;
            chart.SetSourceData(sourceRange);
            ComUtilities.Release(ref chart!);
            return shape;
        }
        catch
        {
            TryDeleteShape(shape);
            throw;
        }
    }

    /// <summary>
    /// Removes a shape that could not become the requested chart, then releases it, so a failed
    /// create never leaves a stray empty chart behind on the worksheet.
    /// </summary>
    private static void TryDeleteShape(dynamic shape)
    {
        try
        {
            shape.Delete();
        }
        catch (COMException)
        {
            // Best-effort cleanup: the original failure is the one that must surface.
        }
        finally
        {
            ComUtilities.Release(ref shape!);
        }
    }

    /// <summary>
    /// Verifies Excel kept the requested chart type. A silent fallback (for example a sunburst
    /// that comes back as an unrelated type) must surface as an error, never as a success.
    /// </summary>
    internal static void VerifyChartType(dynamic shape, ChartType requested, string sheetName)
    {
        dynamic? chart = null;
        try
        {
            chart = shape.Chart;
            int actual = Convert.ToInt32(chart.ChartType, System.Globalization.CultureInfo.InvariantCulture);
            if (actual != (int)requested)
            {
                // Do not leave a mis-typed chart behind for the caller to trip over.
                try
                {
                    shape.Delete();
                }
                catch (COMException)
                {
                    // Best-effort cleanup: the type mismatch is the error that matters.
                }

                throw new InvalidOperationException(
                    $"Excel fell back to a different chart type on sheet '{sheetName}': " +
                    $"requested {requested} ({(int)requested}) but the created chart reports {actual}. " +
                    "That type is not available in this Excel installation.");
            }
        }
        finally
        {
            ComUtilities.Release(ref chart!);
        }
    }

    /// <summary>
    /// Reads the source range column by column and attaches one series per data column.
    /// Text columns become the category (Sunburst: the hierarchy) while numeric columns carry values.
    /// </summary>
    private static void AttachSeriesFromRange(dynamic chart, dynamic sourceRange, int chartTypeCode)
    {
        int firstRow = Convert.ToInt32(sourceRange.Row, System.Globalization.CultureInfo.InvariantCulture);
        int firstColumn = Convert.ToInt32(sourceRange.Column, System.Globalization.CultureInfo.InvariantCulture);
        int rowCount = Convert.ToInt32(sourceRange.Rows.Count, System.Globalization.CultureInfo.InvariantCulture);
        int columnCount = Convert.ToInt32(sourceRange.Columns.Count, System.Globalization.CultureInfo.InvariantCulture);

        if (rowCount < 2 || columnCount < 1)
        {
            throw new InvalidOperationException(
                "Modern charts need at least one header row (or one category row) plus data rows.");
        }

        bool hasHeader = IsTextValue2(sourceRange.Cells[1, 1].Value2);

        // First row is a header only when the first cell is text; otherwise all rows carry data.
        int firstDataOffset = hasHeader ? 2 : 1;
        int lastRow = firstRow + rowCount - 1;
        int firstDataRow = firstRow + (hasHeader ? 1 : 0);

        var textOffsets = new List<int>();
        var valueOffsets = new List<int>();
        for (int offset = 0; offset < columnCount; offset++)
        {
            object? probe = sourceRange.Cells[firstDataOffset, offset + 1].Value2;
            if (IsText(probe))
            {
                textOffsets.Add(offset);
            }
            else
            {
                valueOffsets.Add(offset);
            }
        }

        if (valueOffsets.Count == 0)
        {
            throw new InvalidOperationException(
                "Modern charts need at least one numeric column; the supplied range has none. " +
                "Add a numeric column (or include the header row so the first column can act as categories).");
        }

        // Sunburst nests every text column as hierarchy levels; the combo chart plots every
        // numeric column as its own series; the remaining modern charts are single-series.
        List<int> categoryOffsets = chartTypeCode == 116
            ? textOffsets
            : (textOffsets.Count > 0 && textOffsets[0] == 0 ? new List<int> { 0 } : new List<int>());
        List<int> seriesOffsets = chartTypeCode == 120 ? valueOffsets : new List<int> { valueOffsets[0] };

        dynamic? worksheet = null;
        dynamic? seriesCollection = null;
        try
        {
            worksheet = sourceRange.Worksheet;
            seriesCollection = chart.SeriesCollection();
            foreach (int offset in seriesOffsets)
            {
                dynamic? values = null;
                dynamic? series = null;
                dynamic? categories = null;
                try
                {
                    values = worksheet.Range[
                        worksheet.Cells[firstDataRow, firstColumn + offset],
                        worksheet.Cells[lastRow, firstColumn + offset]];

                    series = seriesCollection.NewSeries();
                    series.Values = values;

                    if (categoryOffsets.Count > 0)
                    {
                        int catOffset = categoryOffsets[0];
                        categories = worksheet.Range[
                            worksheet.Cells[firstDataRow, firstColumn + catOffset],
                            worksheet.Cells[lastRow, firstColumn + catOffset]];
                        series.XValues = categories;
                    }

                    if (hasHeader && worksheet.Cells[firstRow, firstColumn + offset].Value2 is { } header)
                    {
                        series.Name = header.ToString() ?? $"Series{offset + 1}";
                    }
                }
                finally
                {
                    ComUtilities.Release(ref categories!);
                    ComUtilities.Release(ref series!);
                    ComUtilities.Release(ref values!);
                }
            }
        }
        catch (COMException ex) when (ex.HResult == unchecked((int)0x800A03EC))
        {
            throw new InvalidOperationException(
                $"Cannot attach data to a {chartTypeCode} chart. Modern charts need a header row " +
                "plus at least one numeric column, and the values must be in a contiguous range.", ex);
        }
        finally
        {
            ComUtilities.Release(ref seriesCollection!);
            ComUtilities.Release(ref worksheet!);
        }
    }

    private static bool IsTextValue2(object? value) => IsText(value);

    private static bool IsText(object? value) =>
        value is string text && !string.IsNullOrWhiteSpace(text);
}
