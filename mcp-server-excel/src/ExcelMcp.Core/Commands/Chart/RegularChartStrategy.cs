using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;

namespace Sbroenne.ExcelMcp.Core.Commands.Chart;

/// <summary>
/// Strategy for Regular Charts (created from ranges/tables).
/// Handles Shapes.AddChart(), SeriesCollection operations, explicit data source management.
/// </summary>
public class RegularChartStrategy : IChartStrategy
{
    /// <inheritdoc />
    public bool CanHandle(dynamic chart)
    {
        // Regular charts: chart.PivotLayout is null or doesn't exist
        dynamic? pivotLayout = null;
        try
        {
            pivotLayout = chart.PivotLayout;
            return pivotLayout == null;
        }
        catch (COMException)
        {
            return true; // No PivotLayout property = Regular chart
        }
        catch (NotImplementedException)
        {
            // Modern charts (treemap/waterfall/funnel/...) have no PivotLayout member at all:
            // the late binder reports a missing COM member as NotImplementedException, not
            // COMException. Listing or reading such a chart must not crash the whole call.
            return true;
        }
        finally
        {
            ComUtilities.Release(ref pivotLayout);
        }
    }

    /// <inheritdoc />
    public ChartInfo GetInfo(dynamic chart, string chartName, string sheetName, dynamic shape)
    {
        var info = new ChartInfo
        {
            Name = chartName,
            SheetName = sheetName,
            ChartType = (ChartType)Convert.ToInt32(chart.ChartType),
            IsPivotChart = false,
            Left = Convert.ToDouble(shape.Left),
            Top = Convert.ToDouble(shape.Top),
            Width = Convert.ToDouble(shape.Width),
            Height = Convert.ToDouble(shape.Height)
        };

        // Get anchor cells and placement mode
        dynamic? topLeftCell = null;
        dynamic? bottomRightCell = null;
        try
        {
            topLeftCell = shape.TopLeftCell;
            info.TopLeftCell = topLeftCell.Address?.ToString();
        }
        catch (COMException)
        {
            // TopLeftCell not available - optional COM property
        }
        finally
        {
            ComUtilities.Release(ref topLeftCell!);
        }

        try
        {
            bottomRightCell = shape.BottomRightCell;
            info.BottomRightCell = bottomRightCell.Address?.ToString();
        }
        catch (COMException)
        {
            // BottomRightCell not available - optional COM property
        }
        finally
        {
            ComUtilities.Release(ref bottomRightCell!);
        }

        try
        {
            info.Placement = Convert.ToInt32(shape.Placement);
        }
        catch (COMException)
        {
            // Placement not available - optional COM property
        }

        // Count series
        dynamic? seriesCollection = null;
        try
        {
            seriesCollection = chart.SeriesCollection();
            info.SeriesCount = Convert.ToInt32(seriesCollection.Count);
        }
        finally
        {
            ComUtilities.Release(ref seriesCollection!);
        }

        return info;
    }

    /// <inheritdoc />
    public ChartInfoResult GetDetailedInfo(dynamic chart, string chartName, string sheetName, dynamic shape)
    {
        var info = new ChartInfoResult
        {
            Success = true,
            Name = chartName,
            SheetName = sheetName,
            ChartType = (ChartType)Convert.ToInt32(chart.ChartType),
            IsPivotChart = false,
            Left = Convert.ToDouble(shape.Left),
            Top = Convert.ToDouble(shape.Top),
            Width = Convert.ToDouble(shape.Width),
            Height = Convert.ToDouble(shape.Height)
        };

        // Get anchor cells and placement mode
        dynamic? topLeftCell = null;
        dynamic? bottomRightCell = null;
        try
        {
            topLeftCell = shape.TopLeftCell;
            info.TopLeftCell = topLeftCell.Address?.ToString();
        }
        catch (COMException)
        {
            // TopLeftCell not available - optional COM property
        }
        finally
        {
            ComUtilities.Release(ref topLeftCell!);
        }

        try
        {
            bottomRightCell = shape.BottomRightCell;
            info.BottomRightCell = bottomRightCell.Address?.ToString();
        }
        catch (COMException)
        {
            // BottomRightCell not available - optional COM property
        }
        finally
        {
            ComUtilities.Release(ref bottomRightCell!);
        }

        try
        {
            info.Placement = Convert.ToInt32(shape.Placement);
        }
        catch (COMException)
        {
            // Placement not available - optional COM property
        }

        // Get title
        try
        {
            if (chart.HasTitle)
            {
                info.Title = chart.ChartTitle.Text?.ToString() ?? string.Empty;
            }
        }
        catch (COMException)
        {
            // No title - optional COM property, safe to ignore
        }

        // Get legend
        try
        {
            info.HasLegend = chart.HasLegend;
        }
        catch (COMException)
        {
            info.HasLegend = false; // Safe fallback for optional COM property
        }

        // Get source range
        try
        {
            dynamic sourceData = chart.ChartArea.Parent.SeriesCollection(1).Formula;
            info.SourceRange = sourceData?.ToString() ?? string.Empty;
        }
        catch (COMException)
        {
            // No source range or no series - optional COM property, safe to ignore
        }
        catch (NotImplementedException)
        {
            // Modern charts (treemap/waterfall/funnel/...) expose no SERIES formula at all:
            // the late binder reports the missing member as E_NOTIMPL, not as a COM error.
        }

        // Get series
        dynamic? seriesCollection = null;
        try
        {
            seriesCollection = chart.SeriesCollection();
            int seriesCount = Convert.ToInt32(seriesCollection.Count);

            for (int i = 1; i <= seriesCount; i++)
            {
                dynamic? series = null;
                try
                {
                    series = seriesCollection.Item(i);
                    var seriesInfo = new SeriesInfo
                    {
                        Name = SafeText(() => series.Name),
                        // Modern charts refuse to read Values / XValues back (E_NOTIMPL) even though
                        // writing them worked, so both reads are optional by nature.
                        ValuesRange = SafeText(() => series.Values),
                        CategoryRange = SafeText(() => series.XValues)
                    };
                    info.Series.Add(seriesInfo);
                }
                finally
                {
                    if (series != null)
                    {
                        ComUtilities.Release(ref series!);
                    }
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref seriesCollection!);
        }

        return info;
    }

    /// <summary>
    /// Reads an optional chart member and degrades to empty text when Excel refuses it.
    /// The Excel 2016+ charts raise E_NOTIMPL - surfaced by the late binder as
    /// NotImplementedException - when Values / XValues / Formula are read back, so a read that
    /// fails here must not fail the whole inspection.
    /// </summary>
    private static string SafeText(Func<object?> read)
    {
        try
        {
            return read()?.ToString() ?? string.Empty;
        }
        catch (COMException)
        {
            return string.Empty;
        }
        catch (NotImplementedException)
        {
            return string.Empty;
        }
    }

    /// <inheritdoc />
    public void SetSourceRange(dynamic chart, string sourceRange)
    {
        dynamic? sourceRangeObj = null;
        try
        {
            // Get workbook from chart
            dynamic workbook = chart.Parent.Parent.Parent;

            // Get the range object from the address string
            sourceRangeObj = workbook.Application.Range(sourceRange);
            chart.SetSourceData(sourceRangeObj);
        }
        finally
        {
            if (sourceRangeObj != null)
            {
                ComUtilities.Release(ref sourceRangeObj!);
            }
        }
    }

    /// <inheritdoc />
    public SeriesInfo AddSeries(dynamic chart, string seriesName, string valuesRange, string? categoryRange)
    {
        dynamic? seriesCollection = null;
        dynamic? newSeries = null;
        dynamic? valuesRangeObj = null;
        dynamic? categoryRangeObj = null;
        dynamic? worksheet = null;

        try
        {
            // Series.Values and Series.XValues accept a Range object or an array - never an address
            // string. Assigning the string throws 0x800A03EC on every workbook, so resolve the
            // addresses against the worksheet that hosts the chart first.
            worksheet = ResolveWorksheet(chart);
            valuesRangeObj = ResolveRange(worksheet, valuesRange);

            if (!string.IsNullOrWhiteSpace(categoryRange))
            {
                categoryRangeObj = ResolveRange(worksheet, categoryRange);
            }

            seriesCollection = chart.SeriesCollection();
            newSeries = seriesCollection.NewSeries();
            newSeries.Name = seriesName;
            newSeries.Values = valuesRangeObj;

            if (categoryRangeObj != null)
            {
                newSeries.XValues = categoryRangeObj;
            }

            return new SeriesInfo
            {
                Name = seriesName,
                ValuesRange = valuesRange,
                CategoryRange = categoryRange
            };
        }
        finally
        {
            ComUtilities.Release(ref categoryRangeObj!);
            ComUtilities.Release(ref valuesRangeObj!);
            ComUtilities.Release(ref worksheet!);
            if (newSeries != null)
            {
                ComUtilities.Release(ref newSeries!);
            }
            if (seriesCollection != null)
            {
                ComUtilities.Release(ref seriesCollection!);
            }
        }
    }

    /// <summary>
    /// The worksheet that owns the chart's data. An unqualified address like "B2:B7" has no
    /// context of its own, so series ranges are resolved against this worksheet.
    /// </summary>
    private static dynamic ResolveWorksheet(dynamic chart)
    {
        dynamic? chartObject = null;
        try
        {
            chartObject = chart.Parent; // ChartObject when the chart is embedded
            dynamic worksheet = chartObject.Parent; // Worksheet that hosts the embedded chart
            return worksheet;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException(
                $"Cannot resolve the worksheet that hosts chart '{chart.Name}'. " +
                "Series ranges are resolved against the chart's own sheet.", ex);
        }
        finally
        {
            ComUtilities.Release(ref chartObject!);
        }
    }

    /// <summary>
    /// Resolves a range address against the chart's worksheet. A sheet-qualified address
    /// (e.g. "Sheet1!B2:B7") is resolved through the application instead.
    /// </summary>
    private static dynamic ResolveRange(dynamic worksheet, string address)
    {
        if (address.Contains('!'))
        {
            dynamic? application = null;
            try
            {
                application = worksheet.Application;
                return application.Range[address];
            }
            finally
            {
                ComUtilities.Release(ref application!);
            }
        }

        return worksheet.Range[address];
    }

    /// <inheritdoc />
    public void RemoveSeries(dynamic chart, int seriesIndex)
    {
        dynamic? seriesCollection = null;
        dynamic? series = null;

        try
        {
            seriesCollection = chart.SeriesCollection();
            series = seriesCollection.Item(seriesIndex);
            series.Delete();
        }
        finally
        {
            if (series != null)
            {
                ComUtilities.Release(ref series!);
            }
            if (seriesCollection != null)
            {
                ComUtilities.Release(ref seriesCollection!);
            }
        }
    }
}


