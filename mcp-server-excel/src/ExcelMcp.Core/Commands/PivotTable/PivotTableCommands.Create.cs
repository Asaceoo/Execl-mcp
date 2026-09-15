using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

/// <summary>
/// PivotTable creation operations
/// </summary>
public partial class PivotTableCommands
{
    /// <summary>
    /// Creates a PivotTable from an Excel range
    /// Following VBA pattern from ReneNyffenegger/about-MS-Office-object-model
    /// </summary>
    public PivotTableCreateResult CreateFromRange(IExcelBatch batch,
        string sourceSheet, string sourceRange,
        string destinationSheet, string destinationCell,
        string pivotTableName,
        string? shareCacheFrom = null)
    {
        using var timeoutCts = new CancellationTokenSource(TimeSpan.FromMinutes(5));

        return batch.Execute((ctx, ct) =>
        {
            dynamic? sourceWorksheet = null;
            dynamic? sourceRangeObj = null;
            dynamic? destWorksheet = null;
            dynamic? destRangeObj = null;
            dynamic? pivotCaches = null;
            dynamic? pivotCache = null;
            dynamic? pivotTable = null;

            // STEP 1: Validate source range has headers and data
            sourceWorksheet = ctx.Book.Worksheets[sourceSheet];
            sourceRangeObj = sourceWorksheet.Range[sourceRange];

            if (sourceRangeObj.Rows.Count < 2)
            {
                throw new InvalidOperationException($"Source range must contain headers and at least one data row. Found {sourceRangeObj.Rows.Count} rows");
            }

            // STEP 2: Obtain the PivotCache - either a brand new one, or the cache of an existing
            // PivotTable when shareCacheFrom is supplied. Sharing one PivotCache is what allows a
            // single slicer to drive several PivotTables (see ConnectPivotTables).
            pivotCaches = ctx.Book.PivotCaches();
            // Sheet names with spaces or special characters must be quoted: 'Sheet Name'!A1:D6
            string sourceDataRef = $"'{sourceSheet}'!{sourceRange}";

            // xlDatabase = 1, xlPivotTableVersion14 = 4
            // NOTE: all arguments are positional on purpose - the first parameter is dynamic, so this
            // is a dynamic call, and dynamic calls forbid named arguments before positional ones.
            //
            // expectedSourceRef is built through Excel rather than reused verbatim: Excel reports a
            // range PivotCache's SourceData in R1C1 form (SalesData!R1C1:R6C4), so comparing the A1
            // string the caller typed would falsely reject every legitimate share.
            string expectedSourceRef = BuildExpectedRangeSourceRef(sourceSheet, sourceRangeObj, sourceDataRef);
            pivotCache = ResolvePivotCache(
                ctx.Book,
                pivotCaches,
                shareCacheFrom,
                sourceDataRef,
                1,
                4,
                out bool cacheReused,
                expectedSourceRef
            );

            // STEP 3: Create PivotTable from cache
            // VBA: Set pivot_table = pivot_cache.CreatePivotTable(TableDestination:=pivot_table_upper_left)
            destWorksheet = ctx.Book.Worksheets[destinationSheet];
            destRangeObj = destWorksheet.Range[destinationCell];

            pivotTable = pivotCache.CreatePivotTable(
                TableDestination: destRangeObj,
                TableName: pivotTableName
            );

            // STEP 4: Refresh to materialize the PivotTable structure
            pivotTable.RefreshTable();

            // STEP 5: Get available fields from PivotTable (VBA pattern)
            // VBA: Set pf_col_1 = pivot_table.PivotFields("col_1")
            // STEP 5: Get available fields from source range headers
            // These are the fields that CAN be added to the PivotTable
            var availableFields = new List<string>();

            dynamic? headerRow = null;
            try
            {
                headerRow = sourceRangeObj.Rows[1];
                object[,] headers = headerRow.Value2;

                for (int col = 1; col <= headers.GetLength(1); col++)
                {
                    var header = headers[1, col]?.ToString();
                    if (!string.IsNullOrWhiteSpace(header))
                    {
                        availableFields.Add(header);
                    }
                }

                if (availableFields.Count == 0)
                {
                    throw new InvalidOperationException($"No field headers found in source range. Header row has {headers.GetLength(1)} columns.");
                }
            }
            finally
            {
                ComUtilities.Release(ref headerRow);
            }

            try
            {
                return new PivotTableCreateResult
                {
                    Success = true,
                    PivotTableName = pivotTableName,
                    SheetName = destinationSheet,
                    Range = pivotTable.TableRange2.Address,
                    SourceData = sourceDataRef,
                    SourceRowCount = sourceRangeObj.Rows.Count - 1,
                    AvailableFields = availableFields,
                    SharedCacheFrom = cacheReused ? shareCacheFrom : null,
                    WorkflowHint = BuildSharedCacheHint(cacheReused, pivotTableName, shareCacheFrom),
                    FilePath = batch.WorkbookPath
                };
            }
            finally
            {
                ComUtilities.Release(ref pivotTable);
                ComUtilities.Release(ref pivotCache);
                ComUtilities.Release(ref pivotCaches);
                ComUtilities.Release(ref destRangeObj);
                ComUtilities.Release(ref destWorksheet);
                ComUtilities.Release(ref sourceRangeObj);
                ComUtilities.Release(ref sourceWorksheet);
            }
        }, timeoutCts.Token);
    }

    /// <summary>
    /// Creates a PivotTable from an Excel Table
    /// </summary>
    public PivotTableCreateResult CreateFromTable(IExcelBatch batch,
        string tableName,
        string destinationSheet, string destinationCell,
        string pivotTableName,
        string? shareCacheFrom = null)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? table = null;
            dynamic? destWorksheet = null;
            dynamic? destRangeObj = null;
            dynamic? pivotCaches = null;
            dynamic? pivotCache = null;
            dynamic? pivotTable = null;

            // Find the table
            dynamic? sheets = null;
            bool tableFound = false;

            try
            {
                sheets = ctx.Book.Worksheets;
                for (int i = 1; i <= sheets.Count; i++)
                {
                    dynamic? sheet = null;
                    dynamic? listObjects = null;
                    try
                    {
                        sheet = sheets.Item(i);
                        listObjects = sheet.ListObjects;

                        for (int j = 1; j <= listObjects.Count; j++)
                        {
                            dynamic? tbl = null;
                            try
                            {
                                tbl = listObjects.Item(j);
                                if (tbl.Name == tableName)
                                {
                                    table = tbl;
                                    tableFound = true;
                                    break;
                                }
                            }
                            finally
                            {
                                if (tbl != null && tbl != table)
                                {
                                    ComUtilities.Release(ref tbl);
                                }
                            }
                        }

                        if (tableFound) break;
                    }
                    finally
                    {
                        ComUtilities.Release(ref listObjects);
                        ComUtilities.Release(ref sheet);
                    }
                }
            }
            finally
            {
                ComUtilities.Release(ref sheets);
            }

            if (!tableFound || table == null)
            {
                throw new InvalidOperationException($"Table '{tableName}' not found in workbook");
            }

            // Get table range and headers
            dynamic? tableRange = null;
            dynamic? headerRow = null;
            var headers = new List<string>();
            int rowCount = 0;

            try
            {
                tableRange = table.Range;
                rowCount = tableRange.Rows.Count;

                if (rowCount < 2)
                {
                    throw new InvalidOperationException($"Table '{tableName}' must contain at least one data row (has {rowCount} rows including header)");
                }

                // Get headers
                dynamic? headerRowCol = null;
                try
                {
                    headerRowCol = table.HeaderRowRange;
                    object[,] headerValues = headerRowCol.Value2;

                    for (int col = 1; col <= headerValues.GetLength(1); col++)
                    {
                        var header = headerValues[1, col]?.ToString();
                        if (!string.IsNullOrWhiteSpace(header))
                        {
                            headers.Add(header);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref headerRowCol);
                }

                // Obtain the PivotCache - either a brand new one, or the cache of an existing
                // PivotTable when shareCacheFrom is supplied. Sharing one PivotCache is what
                // allows a single slicer to drive several PivotTables.
                pivotCaches = ctx.Book.PivotCaches();
                string sourceDataRef = $"{table.Parent.Name}!{table.Name}";

                // xlDatabase = 1. All positional: dynamic call (see CreateFromRange).
                // expectedSourceData is null on purpose: Excel reports a ListObject-backed cache's
                // SourceData either as the bare table name or sheet-qualified, and the notation is not
                // pinned down by any test here. A wrong guess would falsely reject a legitimate share,
                // which is worse than skipping the guard - the range path, where the notation is known,
                // keeps it.
                pivotCache = ResolvePivotCache(
                    ctx.Book,
                    pivotCaches,
                    shareCacheFrom,
                    sourceDataRef,
                    1,
                    null,
                    out bool cacheReused,
                    null
                );

                // Create PivotTable
                destWorksheet = ctx.Book.Worksheets[destinationSheet];
                destRangeObj = destWorksheet.Range[destinationCell];

                pivotTable = pivotCache.CreatePivotTable(
                    TableDestination: destRangeObj,
                    TableName: pivotTableName
                );

                // Refresh to materialize layout
                pivotTable.RefreshTable();

                try
                {
                    return new PivotTableCreateResult
                    {
                        Success = true,
                        PivotTableName = pivotTableName,
                        SheetName = destinationSheet,
                        Range = pivotTable.TableRange2.Address,
                        SourceData = sourceDataRef,
                        SourceRowCount = rowCount - 1, // Exclude header
                        AvailableFields = headers,
                        SharedCacheFrom = cacheReused ? shareCacheFrom : null,
                        WorkflowHint = BuildSharedCacheHint(cacheReused, pivotTableName, shareCacheFrom),
                        FilePath = batch.WorkbookPath
                    };
                }
                finally
                {
                    ComUtilities.Release(ref pivotTable);
                    ComUtilities.Release(ref pivotCache);
                    ComUtilities.Release(ref pivotCaches);
                    ComUtilities.Release(ref destRangeObj);
                    ComUtilities.Release(ref destWorksheet);
                    ComUtilities.Release(ref table);
                }
            }
            finally
            {
                ComUtilities.Release(ref headerRow);
                ComUtilities.Release(ref tableRange);
            }
        });
    }

    /// <summary>
    /// Creates a PivotTable from a Power Pivot Data Model table
    /// Uses xlExternal source type with "ThisWorkbookDataModel" connection
    /// </summary>
    public PivotTableCreateResult CreateFromDataModel(IExcelBatch batch,
        string tableName,
        string destinationSheet, string destinationCell,
        string pivotTableName,
        string? shareCacheFrom = null)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? model = null;
            dynamic? modelTable = null;
            dynamic? destWorksheet = null;
            dynamic? destRangeObj = null;
            dynamic? pivotCaches = null;
            dynamic? pivotCache = null;
            dynamic? pivotTable = null;

            // STEP 1: Verify Data Model exists and find the table
            // NOTE: Every workbook has a Model object, but it may be empty
            model = ctx.Book.Model;

            // Find the table in the Data Model
            dynamic? modelTables = null;
            bool tableFound = false;
            try
            {
                modelTables = model.ModelTables;

                // Check if Data Model has any tables
                if (modelTables == null || modelTables.Count == 0)
                {
                    throw new InvalidOperationException("Data Model does not contain any tables");
                }

                for (int i = 1; i <= modelTables.Count; i++)
                {
                    dynamic? tbl = null;
                    try
                    {
                        tbl = modelTables.Item(i);
                        if (tbl.Name == tableName)
                        {
                            modelTable = tbl;
                            tableFound = true;
                            break;
                        }
                    }
                    finally
                    {
                        if (tbl != null && tbl != modelTable)
                        {
                            ComUtilities.Release(ref tbl);
                        }
                    }
                }
            }
            finally
            {
                ComUtilities.Release(ref modelTables);
            }

            if (!tableFound || modelTable == null)
            {
                throw new InvalidOperationException($"Table '{tableName}' not found in Data Model");
            }

            // Get columns from the Data Model table
            var headers = new List<string>();
            int recordCount = 0;

            try
            {
                recordCount = ComUtilities.SafeGetInt(modelTable, "RecordCount");

                // Get columns
                dynamic? modelColumns = null;
                try
                {
                    modelColumns = modelTable.ModelTableColumns;
                    for (int i = 1; i <= modelColumns.Count; i++)
                    {
                        dynamic? column = null;
                        try
                        {
                            column = modelColumns.Item(i);
                            var colName = ComUtilities.SafeGetString(column, "Name");
                            if (!string.IsNullOrWhiteSpace(colName))
                            {
                                headers.Add(colName);
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref column);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref modelColumns);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to read columns from Data Model table '{tableName}': {ex.Message}");
            }

            if (headers.Count == 0)
            {
                throw new InvalidOperationException($"Data Model table '{tableName}' has no columns");
            }

            // STEP 2: Obtain the PivotCache - either a brand new one, or the cache of an existing
            // PivotTable. Using xlExternal (2) with "ThisWorkbookDataModel" connection.
            // No source validation here: Excel reports OLAP SourceData as a provider string that
            // cannot be compared reliably, and every Data Model cache points at the same model.
            pivotCaches = ctx.Book.PivotCaches();

            // xlExternal = 2. All positional: dynamic call (see CreateFromRange).
            // expectedSourceData omitted (null): OLAP/Data Model SourceData is a provider string that
            // cannot be compared, so the source guard is intentionally skipped here.
            pivotCache = ResolvePivotCache(
                ctx.Book,
                pivotCaches,
                shareCacheFrom,
                "ThisWorkbookDataModel",
                2,
                null,
                out bool cacheReused,
                null
            );

            // STEP 3: Create PivotTable from cache
            destWorksheet = ctx.Book.Worksheets[destinationSheet];
            destRangeObj = destWorksheet.Range[destinationCell];

            pivotTable = pivotCache.CreatePivotTable(
                TableDestination: destRangeObj,
                TableName: pivotTableName
            );

            // STEP 4: Refresh to materialize the PivotTable structure
            pivotTable.RefreshTable();

            try
            {
                return new PivotTableCreateResult
                {
                    Success = true,
                    PivotTableName = pivotTableName,
                    SheetName = destinationSheet,
                    Range = pivotTable.TableRange2.Address,
                    SourceData = $"ThisWorkbookDataModel[{tableName}]",
                    SourceRowCount = recordCount,
                    AvailableFields = headers,
                    SharedCacheFrom = cacheReused ? shareCacheFrom : null,
                    WorkflowHint = BuildSharedCacheHint(cacheReused, pivotTableName, shareCacheFrom),
                    FilePath = batch.WorkbookPath
                };
            }
            finally
            {
                ComUtilities.Release(ref pivotTable);
                ComUtilities.Release(ref pivotCache);
                ComUtilities.Release(ref pivotCaches);
                ComUtilities.Release(ref destRangeObj);
                ComUtilities.Release(ref destWorksheet);
                ComUtilities.Release(ref modelTable);
                ComUtilities.Release(ref model);
            }
        });
    }

    /// <summary>
    /// Builds the hint returned when a PivotTable was created on a shared PivotCache.
    ///
    /// All three create paths route through here so the guidance does not depend on which source kind
    /// was used - telling a caller about 'connect-pivots' from only one of them left the other two
    /// silently unexplained. Returns null when the cache was not shared.
    /// </summary>
    private static string? BuildSharedCacheHint(bool cacheReused, string pivotTableName, string? shareCacheFrom)
        => cacheReused
            ? $"PivotTable '{pivotTableName}' reuses the PivotCache of '{shareCacheFrom}'. "
              + "Use slicer 'create-slicer' on any of them, then 'connect-pivots' to add the rest."
            : null;

    /// <summary>
    /// Resolves the PivotCache a new PivotTable should be built on.
    ///
    /// When <paramref name="shareCacheFrom"/> is null the behaviour is unchanged: a brand new
    /// PivotCache is created from the supplied source. When it names an existing PivotTable, that
    /// PivotTable's cache is reused instead. A shared PivotCache is the precondition for one
    /// slicer to filter several PivotTables - Excel only lets a slicer drive PivotTables that
    /// share a single PivotCache (see the slicer 'connect-pivots' action).
    /// </summary>
    /// <param name="workbook">Workbook that owns the PivotTables</param>
    /// <param name="pivotCaches">The workbook's PivotCaches collection, already resolved by the caller</param>
    /// <param name="shareCacheFrom">Optional name of an existing PivotTable whose cache is reused</param>
    /// <param name="sourceDataRef">Source reference used when a new cache has to be created</param>
    /// <param name="sourceType">xlDatabase (1) for ranges and tables, xlExternal (2) for the Data Model</param>
    /// <param name="sourceVersion">PivotTable version for newly created caches; null omits the argument</param>
    /// <param name="cacheReused">True when an existing cache was reused rather than created</param>
    /// <param name="expectedSourceData">
    /// Best-effort guard. When supplied and the reused cache demonstrably points at a different
    /// source, the operation is rejected instead of silently producing PivotTables that no slicer
    /// can link. The caller must pass the reference in the notation Excel itself reports from
    /// PivotCache.SourceData (for ranges that is R1C1 - see BuildExpectedRangeSourceRef). Pass null
    /// for sources Excel reports in a form that cannot be canonicalised (ListObject and Data Model
    /// sources), because a false rejection would be worse than skipping the guard.
    /// </param>
    /// <returns>The PivotCache the caller should build the PivotTable on; caller owns the reference</returns>
    private static dynamic ResolvePivotCache(
        dynamic workbook,
        dynamic pivotCaches,
        string? shareCacheFrom,
        string sourceDataRef,
        int sourceType,
        int? sourceVersion,
        out bool cacheReused,
        string? expectedSourceData = null)
    {
        cacheReused = false;

        if (string.IsNullOrWhiteSpace(shareCacheFrom))
        {
            dynamic created = sourceVersion.HasValue
                ? pivotCaches.Create(
                    SourceType: sourceType,
                    SourceData: sourceDataRef,
                    Version: sourceVersion.Value)
                : pivotCaches.Create(
                    SourceType: sourceType,
                    SourceData: sourceDataRef);

            return created;
        }

        dynamic? sourcePivot = null;
        dynamic? sharedCache = null;

        try
        {
            sourcePivot = FindPivotTable(workbook, shareCacheFrom);
            sharedCache = sourcePivot.PivotCache;

            if (expectedSourceData != null)
            {
                string actualSourceData = ReadPivotCacheSourceData(sharedCache);
                if (SourceDataConflict(expectedSourceData, actualSourceData))
                {
                    throw new InvalidOperationException(
                        $"Cannot share PivotCache: PivotTable '{shareCacheFrom}' is built on '{actualSourceData}', "
                        + $"but this request targets '{expectedSourceData}'. Point share_cache_from at a PivotTable "
                        + "built from the same source, or omit it to create an independent PivotCache.");
                }
            }

            cacheReused = true;

            dynamic resolved = sharedCache;
            sharedCache = null; // ownership moves to the caller - do not release it here
            return resolved;
        }
        finally
        {
            ComUtilities.Release(ref sharedCache);
            ComUtilities.Release(ref sourcePivot);
        }
    }

    /// <summary>
    /// Builds the reference the shared-cache guard compares against, expressed in the notation Excel
    /// itself uses for PivotCache.SourceData.
    ///
    /// Excel rewrites a range source into R1C1 form: a cache built from 'SalesData'!A1:D6 is reported
    /// as SalesData!R1C1:R6C4. Comparing the A1 string the caller typed would therefore reject every
    /// legitimate share_cache_from call, so the range is converted through Excel first. When Excel
    /// refuses to produce an address the raw fallback is used, which restores the previous
    /// (potentially stricter) behaviour instead of dropping the guard.
    /// </summary>
    private static string BuildExpectedRangeSourceRef(string sheetName, dynamic rangeObject, string fallbackRef)
    {
        try
        {
            // xlR1C1 = -4150. The default absolute flags produce the anchor-free R1C1:R6C4 form Excel stores.
            string r1c1Address = rangeObject.Address(ReferenceStyle: -4150);
            return string.IsNullOrWhiteSpace(r1c1Address)
                ? fallbackRef
                : $"{sheetName}!{r1c1Address}";
        }
        catch (Exception)
        {
            return fallbackRef;
        }
    }

    /// <summary>
    /// Reads a PivotCache's SourceData. Returns an empty string when the property is unavailable.
    /// </summary>
    private static string ReadPivotCacheSourceData(dynamic cache)
    {
        try
        {
            return cache.SourceData?.ToString() ?? string.Empty;
        }
        catch (Exception)
        {
            return string.Empty;
        }
    }

    /// <summary>
    /// Best-effort comparison of two source references.
    ///
    /// Excel rewrites SourceData when it stores a cache: 'Sheet1'!A1:D6 comes back as
    /// Sheet1!R1C1:R6C4, and it quotes sheet names that need it. Raw string equality is therefore
    /// useless. This returns true only when both references are readable and demonstrably describe
    /// different targets - an unreadable value never blocks the operation, because a false rejection
    /// would be worse than a missing guard.
    ///
    /// The comparison is deliberately exact once normalised. An earlier version also accepted one
    /// reference being a suffix of the other, meant to cover a ListObject reported as either the bare
    /// table name or Sheet!Table; that path has since been excluded (the ListObject and Data Model
    /// callers pass null), and the tolerance turned two different sheets whose names nest - 'Data' and
    /// 'SourceData' - into a silent match, letting a PivotTable adopt a cache built on another sheet.
    /// </summary>
    private static bool SourceDataConflict(string expected, string actual)
    {
        string normalizedExpected = NormalizeSourceRef(expected);
        string normalizedActual = NormalizeSourceRef(actual);

        if (normalizedExpected.Length == 0 || normalizedActual.Length == 0)
        {
            return false;
        }

        return !string.Equals(normalizedExpected, normalizedActual, StringComparison.Ordinal);
    }

    /// <summary>
    /// Upper-cases a source reference and strips the quoting and anchors Excel adds
    /// ('Sheet Name'!$A$1:$D$6), so equivalent references compare equal.
    ///
    /// Whitespace is preserved on purpose: removing it would make distinct sheet names equal
    /// ('Sales Data' and 'SalesData'), which is exactly the confusion the guard exists to catch.
    /// See SlicerLink_ShareCacheFromSuffixSheetName_IsRejectedAsADifferentSource for the pinned case.
    /// </summary>
    private static string NormalizeSourceRef(string reference)
    {
        if (string.IsNullOrWhiteSpace(reference))
        {
            return string.Empty;
        }

        var builder = new System.Text.StringBuilder(reference.Length);
        foreach (char c in reference)
        {
            if (c is '$' or '\'' or '`' or '"')
            {
                continue;
            }

            builder.Append(char.ToUpperInvariant(c));
        }

        return builder.ToString().Trim();
    }
}



