using System.Runtime.InteropServices;

using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

/// <summary>
/// PivotTable slicer operations (CreateSlicer, ListSlicers, SetSlicerSelection, DeleteSlicer)
/// </summary>
public partial class PivotTableCommands
{
    /// <summary>
    /// Creates a slicer for a PivotTable field
    /// </summary>
    public SlicerResult CreateSlicer(IExcelBatch batch, string pivotTableName,
        string fieldName, string slicerName, string destinationSheet, string position)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? pivot = null;
            dynamic? slicerCaches = null;
            dynamic? slicerCache = null;
            dynamic? slicers = null;
            dynamic? slicer = null;
            dynamic? destSheet = null;
            dynamic? destRange = null;

            try
            {
                pivot = FindPivotTable(ctx.Book, pivotTableName);
                slicerCaches = ctx.Book.SlicerCaches;

                // Check if a SlicerCache already exists for this field+PivotTable
                // If so, we add a new visual Slicer to the existing cache
                slicerCache = FindExistingSlicerCache(slicerCaches, pivot, fieldName);

                if (slicerCache == null)
                {
                    // Create new SlicerCache for this field
                    // SlicerCaches.Add(source, sourceField, name, slicerCacheType)
                    // source = PivotTable object
                    // sourceField = field name string
                    // name = cache name (optional, auto-generated if not provided)

                    // For regular PivotTables, use field name directly
                    // For OLAP, may need the hierarchical name
                    slicerCache = slicerCaches.Add2(pivot, fieldName);
                }

                // Get destination sheet and calculate position from cell reference
                destSheet = ctx.Book.Worksheets[destinationSheet];
                destRange = destSheet.Range[position];

                // Get position in points from the cell reference
                double top = Convert.ToDouble(destRange.Top);
                double left = Convert.ToDouble(destRange.Left);

                // Add visual Slicer to the cache
                // Slicers.Add(SlicerDestination, Level, Name, Caption, Top, Left, Width, Height)
                // For non-OLAP sources, Level should be Type.Missing or omitted
                slicers = slicerCache.Slicers;
                slicer = slicers.Add(destSheet, Type.Missing, slicerName, slicerName, top, left);

                // Build result
                var result = BuildSlicerResult(slicer, slicerCache, fieldName);
                result.Success = true;
                result.WorkflowHint = $"Slicer '{slicerName}' created for field '{fieldName}'. It currently filters "
                    + $"{result.ConnectedPivotTables.Count} PivotTable(s). "
                    + "Use 'set-slicer-selection' to filter, and 'connect-pivots' to make it drive further "
                    + "PivotTables that share the same PivotCache (create those with pivot table 'share_cache_from').";

                return result;
            }
            finally
            {
                ComUtilities.Release(ref destRange);
                ComUtilities.Release(ref destSheet);
                ComUtilities.Release(ref slicer);
                ComUtilities.Release(ref slicers);
                ComUtilities.Release(ref slicerCache);
                ComUtilities.Release(ref slicerCaches);
                ComUtilities.Release(ref pivot);
            }
        });
    }

    /// <summary>
    /// Lists all slicers in the workbook, optionally filtered by PivotTable
    /// </summary>
    public SlicerListResult ListSlicers(IExcelBatch batch, string? pivotTableName = null)
    {
        return batch.Execute((ctx, ct) =>
        {
            var result = new SlicerListResult { Success = true };
            dynamic? slicerCaches = null;
            dynamic? targetPivot = null;

            try
            {
                slicerCaches = ctx.Book.SlicerCaches;

                // If filtering by PivotTable, find it first
                if (!string.IsNullOrEmpty(pivotTableName))
                {
                    targetPivot = FindPivotTable(ctx.Book, pivotTableName);
                }

                for (int cacheIndex = 1; cacheIndex <= slicerCaches.Count; cacheIndex++)
                {
                    dynamic? cache = null;
                    dynamic? slicers = null;

                    try
                    {
                        cache = slicerCaches.Item(cacheIndex);

                        // If filtering by PivotTable, check if this cache is connected
                        if (targetPivot != null && !IsSlicerCacheConnectedToPivot(cache, targetPivot))
                        {
                            continue;
                        }

                        slicers = cache.Slicers;
                        for (int slicerIndex = 1; slicerIndex <= slicers.Count; slicerIndex++)
                        {
                            dynamic? slicer = null;
                            try
                            {
                                slicer = slicers.Item(slicerIndex);
                                var slicerInfo = BuildSlicerInfo(slicer, cache);
                                result.Slicers.Add(slicerInfo);
                            }
                            finally
                            {
                                ComUtilities.Release(ref slicer);
                            }
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref slicers);
                        ComUtilities.Release(ref cache);
                    }
                }

                return result;
            }
            finally
            {
                ComUtilities.Release(ref targetPivot);
                ComUtilities.Release(ref slicerCaches);
            }
        });
    }

    /// <summary>
    /// Sets the selection for a slicer
    /// </summary>
    public SlicerResult SetSlicerSelection(IExcelBatch batch, string slicerName,
        List<string> selectedItems, bool clearFirst = true)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? slicerCaches = null;
            dynamic? targetCache = null;
            dynamic? targetSlicer = null;
            dynamic? slicerItems = null;

            try
            {
                slicerCaches = ctx.Book.SlicerCaches;

                // Find the slicer by name
                var searchResult = FindSlicerByName(slicerCaches, slicerName);
                targetCache = searchResult.Cache;
                targetSlicer = searchResult.Slicer;

                if (targetSlicer == null || targetCache == null)
                {
                    return new SlicerResult
                    {
                        Success = false,
                        ErrorMessage = $"Slicer '{slicerName}' not found in workbook"
                    };
                }

                // Get slicer items from the cache
                slicerItems = targetCache.SlicerItems;

                // Build set of items to select for fast lookup
                var itemsToSelect = new HashSet<string>(selectedItems, StringComparer.OrdinalIgnoreCase);

                // If no items specified, select all (clear filter)
                bool selectAll = selectedItems.Count == 0;

                // Iterate through slicer items and set selection
                for (int i = 1; i <= slicerItems.Count; i++)
                {
                    dynamic? item = null;
                    try
                    {
                        item = slicerItems.Item(i);
                        string itemName = item.Name?.ToString() ?? string.Empty;

                        if (selectAll)
                        {
                            item.Selected = true;
                        }
                        else if (clearFirst)
                        {
                            // Clear first mode: select only specified items
                            item.Selected = itemsToSelect.Contains(itemName);
                        }
                        else
                        {
                            // Additive mode: add to existing selection
                            if (itemsToSelect.Contains(itemName))
                            {
                                item.Selected = true;
                            }
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref item);
                    }
                }

                // Build result with updated state
                string fieldName = GetSlicerCacheFieldName(targetCache);
                var result = BuildSlicerResult(targetSlicer, targetCache, fieldName);
                result.Success = true;
                result.WorkflowHint = selectAll
                    ? $"Slicer '{slicerName}' filter cleared - all items are now visible."
                    : $"Slicer '{slicerName}' selection updated to {selectedItems.Count} item(s).";

                return result;
            }
            finally
            {
                ComUtilities.Release(ref slicerItems);
                ComUtilities.Release(ref targetSlicer);
                ComUtilities.Release(ref targetCache);
                ComUtilities.Release(ref slicerCaches);
            }
        });
    }

    /// <summary>
    /// Deletes a slicer from the workbook
    /// </summary>
    public OperationResult DeleteSlicer(IExcelBatch batch, string slicerName)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? slicerCaches = null;
            dynamic? targetCache = null;
            dynamic? targetSlicer = null;

            try
            {
                slicerCaches = ctx.Book.SlicerCaches;

                // Find the slicer by name
                var searchResult = FindSlicerByName(slicerCaches, slicerName);
                targetCache = searchResult.Cache;
                targetSlicer = searchResult.Slicer;

                if (targetSlicer == null)
                {
                    return new OperationResult
                    {
                        Success = false,
                        ErrorMessage = $"Slicer '{slicerName}' not found in workbook"
                    };
                }

                // Delete the visual slicer
                targetSlicer.Delete();

                // Note: The SlicerCache will be automatically deleted if this was the last slicer
                // connected to it. Excel handles this automatically.

                return new OperationResult { Success = true };
            }
            finally
            {
                ComUtilities.Release(ref targetSlicer);
                ComUtilities.Release(ref targetCache);
                ComUtilities.Release(ref slicerCaches);
            }
        });
    }

    /// <summary>
    /// Connects additional PivotTables to an existing PivotTable slicer so a single slicer filters
    /// them all at once.
    ///
    /// Excel only lets a slicer drive PivotTables that share one PivotCache, and it enforces that by
    /// raising a bare COM error (0x800A03EC) from SlicerPivotTables.AddPivotTable - it neither
    /// documents the rule nor explains the failure. This method therefore applies each connection,
    /// then re-reads the slicer's live connection list through Excel and reports any requested
    /// PivotTable that did not stick, replacing the opaque COM error with an actionable message.
    /// It deliberately does not try to predict the outcome from PivotCache.Index: that property was
    /// measured on a real machine to return 0 for two independent caches and 0/1 for one shared
    /// cache, so it cannot identify a cache.
    /// </summary>
    public SlicerResult ConnectPivotTables(IExcelBatch batch, string slicerName,
        List<string> pivotTableNames)
    {
        // Rejected before touching Excel: an empty request would otherwise report a misleading
        // success ("already connected"), and a request made only of blank names would do the same
        // without ever looking a single PivotTable up.
        if (!TryCleanPivotTableNames(pivotTableNames, out var requestedNames))
        {
            return new SlicerResult
            {
                Success = false,
                ErrorMessage = "connect-pivots requires at least one non-empty PivotTable name "
                    + "in pivotTableNames."
            };
        }

        return batch.Execute((ctx, ct) =>
        {
            dynamic? slicerCaches = null;
            dynamic? targetCache = null;
            dynamic? targetSlicer = null;
            dynamic? pivotTables = null;

            try
            {
                slicerCaches = ctx.Book.SlicerCaches;

                var searchResult = FindSlicerByName(slicerCaches, slicerName);
                targetCache = searchResult.Cache;
                targetSlicer = searchResult.Slicer;

                if (targetSlicer == null || targetCache == null)
                {
                    return new SlicerResult
                    {
                        Success = false,
                        ErrorMessage = $"Slicer '{slicerName}' not found in workbook"
                    };
                }

                if (targetCache.List == true)
                {
                    return new SlicerResult
                    {
                        Success = false,
                        ErrorMessage = $"Slicer '{slicerName}' is a Table slicer and cannot be connected to PivotTables."
                    };
                }

                // Resolve every requested name before mutating anything, so a typo is reported
                // without leaving the slicer partially connected.
                var alreadyConnected = new HashSet<string>(
                    GetConnectedPivotTableNames(targetCache), StringComparer.OrdinalIgnoreCase);
                var toAdd = new List<string>();

                foreach (string name in requestedNames)
                {
                    if (alreadyConnected.Contains(name))
                    {
                        continue;
                    }

                    if (!PivotTableExists(ctx.Book, name))
                    {
                        return new SlicerResult
                        {
                            Success = false,
                            ErrorMessage = $"PivotTable '{name}' not found in workbook"
                        };
                    }

                    toAdd.Add(name);
                }

                var added = new List<string>();
                if (toAdd.Count > 0)
                {
                    pivotTables = targetCache.PivotTables;

                    foreach (string name in toAdd)
                    {
                        dynamic? pivot = null;
                        try
                        {
                            pivot = FindPivotTable(ctx.Book, name);
                            pivotTables.AddPivotTable(pivot);
                            added.Add(name);
                        }
                        catch (System.Runtime.InteropServices.COMException)
                        {
                            // Excel refuses PivotTables that do not share the slicer's PivotCache.
                            // Swallowed here so every requested name is attempted and the refreshed
                            // connection list below decides what actually took effect.
                        }
                        finally
                        {
                            ComUtilities.Release(ref pivot);
                        }
                    }
                }

                // BuildSlicerResult reads the slicer's current connections from Excel, so this is
                // verification of the applied state rather than a replay of our own bookkeeping.
                var result = BuildSlicerResult(targetSlicer, targetCache, GetSlicerCacheFieldName(targetCache));
                var connectedNow = new HashSet<string>(result.ConnectedPivotTables, StringComparer.OrdinalIgnoreCase);
                var applied = toAdd.Where(connectedNow.Contains).ToList();
                var refused = toAdd.Where(name => !connectedNow.Contains(name)).ToList();

                if (refused.Count > 0)
                {
                    // A partly applied request must not read as "nothing happened": Excel does not roll
                    // back the connections that did stick, so the message has to state the live list.
                    string state = result.ConnectedPivotTables.Count == 0
                        ? "The slicer currently filters no PivotTables."
                        : $"The slicer now filters {result.ConnectedPivotTables.Count}: "
                          + $"{string.Join(", ", result.ConnectedPivotTables)}.";

                    result.Success = false;
                    result.ErrorMessage =
                        $"Excel refused to connect {string.Join(", ", refused)} to slicer '{slicerName}'. A slicer can "
                        + "only drive PivotTables that share a single PivotCache."
                        + (applied.Count == 0
                            ? string.Empty
                            : applied.Count == 1
                                ? $" {applied[0]} was connected before the request failed and stays connected."
                                : $" {string.Join(", ", applied)} were connected before the request failed and stay connected.")
                        + $" {state} To link the refused PivotTables, recreate each with pivot table "
                        + "'share_cache_from' set to a PivotTable this slicer already filters, then retry.";
                    return result;
                }

                result.Success = true;
                result.WorkflowHint = added.Count == 0
                    ? $"Slicer '{slicerName}' was already connected to the requested PivotTable(s)."
                    : $"Slicer '{slicerName}' now filters {result.ConnectedPivotTables.Count} PivotTable(s); "
                      + $"added {string.Join(", ", added)}. Use 'set-slicer-selection' to filter them together.";

                return result;
            }
            finally
            {
                ComUtilities.Release(ref pivotTables);
                ComUtilities.Release(ref targetSlicer);
                ComUtilities.Release(ref targetCache);
                ComUtilities.Release(ref slicerCaches);
            }
        });
    }

    /// <summary>
    /// Removes PivotTables from a slicer's connection list so the slicer stops filtering them.
    /// A slicer must keep at least one connection, so removing the last one is rejected.
    /// </summary>
    public SlicerResult DisconnectPivotTables(IExcelBatch batch, string slicerName,
        List<string> pivotTableNames)
    {
        // Same contract as ConnectPivotTables: distinct, non-blank names only. Counting list entries
        // rather than distinct PivotTables would let a repeated name trip the last-connection guard,
        // and would also aim a second RemovePivotTable at an entry that is already gone.
        if (!TryCleanPivotTableNames(pivotTableNames, out var requestedNames))
        {
            return new SlicerResult
            {
                Success = false,
                ErrorMessage = "disconnect-pivots requires at least one non-empty PivotTable name "
                    + "in pivotTableNames."
            };
        }

        return batch.Execute((ctx, ct) =>
        {
            dynamic? slicerCaches = null;
            dynamic? targetCache = null;
            dynamic? targetSlicer = null;
            dynamic? pivotTables = null;

            try
            {
                slicerCaches = ctx.Book.SlicerCaches;

                var searchResult = FindSlicerByName(slicerCaches, slicerName);
                targetCache = searchResult.Cache;
                targetSlicer = searchResult.Slicer;

                if (targetSlicer == null || targetCache == null)
                {
                    return new SlicerResult
                    {
                        Success = false,
                        ErrorMessage = $"Slicer '{slicerName}' not found in workbook"
                    };
                }

                if (targetCache.List == true)
                {
                    return new SlicerResult
                    {
                        Success = false,
                        ErrorMessage = $"Slicer '{slicerName}' is a Table slicer; its connections are managed by Excel."
                    };
                }

                var connected = GetConnectedPivotTableNames(targetCache);

                // Match requested names against the actual connection list, preserving Excel's casing.
                var toRemove = new List<string>();
                foreach (string name in requestedNames)
                {
                    foreach (string existing in connected)
                    {
                        if (string.Equals(existing, name, StringComparison.OrdinalIgnoreCase))
                        {
                            toRemove.Add(existing);
                            break;
                        }
                    }
                }

                if (toRemove.Count == 0)
                {
                    return new SlicerResult
                    {
                        Success = false,
                        ErrorMessage = $"None of the requested PivotTables are connected to slicer '{slicerName}'. "
                            + $"Currently connected: {string.Join(", ", connected)}."
                    };
                }

                // toRemove holds distinct PivotTables, so this counts real removals: a slicer has to
                // keep at least one connection.
                if (connected.Count - toRemove.Count < 1)
                {
                    return new SlicerResult
                    {
                        Success = false,
                        ErrorMessage = $"A slicer must stay connected to at least one PivotTable, so "
                            + $"{string.Join(", ", toRemove)} cannot be removed. "
                            + $"Currently connected: {string.Join(", ", connected)}."
                    };
                }

                pivotTables = targetCache.PivotTables;
                foreach (string name in toRemove)
                {
                    try
                    {
                        // The parameter is declared as a PivotTable, but Excel resolves a name here as
                        // well - which this path relies on and SlicerLink_DisconnectPivots_* pins down.
                        pivotTables.RemovePivotTable(name);
                    }
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        // Swallowed to match ConnectPivotTables: the re-read below decides what actually
                        // happened, so a refusal is reported as an actionable message instead of a raw
                        // 0x8002000B/0x800A03EC escaping to the caller.
                    }
                }

                var result = BuildSlicerResult(targetSlicer, targetCache, GetSlicerCacheFieldName(targetCache));

                // Verify against Excel's live connection list rather than trusting the call to have
                // taken effect - the same discipline ConnectPivotTables uses.
                var stillConnected = new HashSet<string>(result.ConnectedPivotTables, StringComparer.OrdinalIgnoreCase);
                var notRemoved = toRemove.Where(name => stillConnected.Contains(name)).ToList();

                if (notRemoved.Count > 0)
                {
                    result.Success = false;
                    result.ErrorMessage =
                        $"Excel kept {string.Join(", ", notRemoved)} connected to slicer '{slicerName}'. "
                        + $"Currently connected: {string.Join(", ", result.ConnectedPivotTables)}.";
                    return result;
                }

                result.Success = true;
                result.WorkflowHint = $"Removed {string.Join(", ", toRemove)} from slicer '{slicerName}'; "
                    + $"it now filters {result.ConnectedPivotTables.Count} PivotTable(s).";

                return result;
            }
            finally
            {
                ComUtilities.Release(ref pivotTables);
                ComUtilities.Release(ref targetSlicer);
                ComUtilities.Release(ref targetCache);
                ComUtilities.Release(ref slicerCaches);
            }
        });
    }

    #region Slicer Helper Methods

    /// <summary>
    /// Result of searching for a slicer by name (avoids dynamic tuple deconstruction)
    /// </summary>
    private readonly struct SlicerSearchResult
    {
        public dynamic? Cache { get; init; }
        public dynamic? Slicer { get; init; }
    }

    /// <summary>
    /// Finds an existing SlicerCache for a field on a specific PivotTable
    /// </summary>
    private static dynamic? FindExistingSlicerCache(dynamic slicerCaches, dynamic pivot, string fieldName)
    {
        for (int i = 1; i <= slicerCaches.Count; i++)
        {
            dynamic? cache = null;
            try
            {
                cache = slicerCaches.Item(i);

                // Check if cache is for the same field
                string cacheFieldName = GetSlicerCacheFieldName(cache);
                if (!string.Equals(cacheFieldName, fieldName, StringComparison.OrdinalIgnoreCase))
                {
                    ComUtilities.Release(ref cache);
                    continue;
                }

                // Check if this cache is connected to our PivotTable
                if (IsSlicerCacheConnectedToPivot(cache, pivot))
                {
                    return cache; // Don't release - returning to caller
                }

                ComUtilities.Release(ref cache);
            }
            catch (COMException)
            {
                // COM property access may fail for certain cache types - continue searching
                ComUtilities.Release(ref cache);
            }
        }

        return null;
    }

    /// <summary>
    /// Gets the source field name from a SlicerCache
    /// </summary>
    private static string GetSlicerCacheFieldName(dynamic cache)
    {
        dynamic? sourceField = null;
        try
        {
            // Try to get SourceName first (OLAP), then fall back to checking PivotField
            try
            {
                string? sourceName = cache.SourceName?.ToString();
                if (!string.IsNullOrEmpty(sourceName))
                    return sourceName;
            }
            catch (COMException)
            {
                // SourceName property not available for this cache type - fall back to Name
            }

            // For regular slicers, get from PivotTables collection
            dynamic? pivotTables = null;
            try
            {
                pivotTables = cache.PivotTables;
                if (pivotTables != null && pivotTables.Count > 0)
                {
                    dynamic? pt = null;
                    try
                    {
                        pt = pivotTables.Item(1);
                        // The cache Name often contains the field name
                        string cacheName = cache.Name?.ToString() ?? string.Empty;
                        // SlicerCache names are typically "Slicer_FieldName" format
                        if (cacheName.StartsWith("Slicer_", StringComparison.OrdinalIgnoreCase))
                        {
                            return cacheName[7..]; // Remove "Slicer_" prefix
                        }
                        return cacheName;
                    }
                    finally
                    {
                        ComUtilities.Release(ref pt);
                    }
                }
            }
            finally
            {
                ComUtilities.Release(ref pivotTables);
            }

            return cache.Name?.ToString() ?? "Unknown";
        }
        finally
        {
            ComUtilities.Release(ref sourceField);
        }
    }

    /// <summary>
    /// Checks if a SlicerCache is connected to a specific PivotTable.
    /// Returns false for Table slicers (cache.List == true) since they don't connect to PivotTables.
    /// </summary>
    private static bool IsSlicerCacheConnectedToPivot(dynamic cache, dynamic targetPivot)
    {
        // Per MS docs: List property is true for Table slicers, false for PivotTable slicers
        // https://learn.microsoft.com/en-us/office/vba/api/excel.slicercache.list
        // Table slicers don't connect to PivotTables
        if (cache.List == true)
        {
            return false;
        }

        dynamic? pivotTables = null;
        try
        {
            pivotTables = cache.PivotTables;
            string targetName = targetPivot.Name?.ToString() ?? string.Empty;

            for (int i = 1; i <= pivotTables.Count; i++)
            {
                dynamic? pt = null;
                try
                {
                    pt = pivotTables.Item(i);
                    string ptName = pt.Name?.ToString() ?? string.Empty;
                    if (string.Equals(ptName, targetName, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
                finally
                {
                    ComUtilities.Release(ref pt);
                }
            }
            return false;
        }
        finally
        {
            ComUtilities.Release(ref pivotTables);
        }
    }

    /// <summary>
    /// Finds a slicer by name across all SlicerCaches
    /// </summary>
    private static SlicerSearchResult FindSlicerByName(dynamic slicerCaches, string slicerName)
    {
        for (int cacheIndex = 1; cacheIndex <= slicerCaches.Count; cacheIndex++)
        {
            dynamic? cache = null;
            dynamic? slicers = null;

            try
            {
                cache = slicerCaches.Item(cacheIndex);
                slicers = cache.Slicers;

                for (int slicerIndex = 1; slicerIndex <= slicers.Count; slicerIndex++)
                {
                    dynamic? slicer = null;
                    try
                    {
                        slicer = slicers.Item(slicerIndex);
                        string name = slicer.Name?.ToString() ?? string.Empty;

                        if (string.Equals(name, slicerName, StringComparison.OrdinalIgnoreCase))
                        {
                            // Found it - return both cache and slicer (don't release)
                            ComUtilities.Release(ref slicers);
                            return new SlicerSearchResult { Cache = cache, Slicer = slicer };
                        }
                        ComUtilities.Release(ref slicer);
                    }
                    catch (COMException)
                    {
                        // COM access may fail for certain slicer types - continue searching
                        ComUtilities.Release(ref slicer);
                    }
                }

                ComUtilities.Release(ref slicers);
                ComUtilities.Release(ref cache);
            }
            catch (COMException)
            {
                // COM access may fail for certain cache types - continue searching
                ComUtilities.Release(ref slicers);
                ComUtilities.Release(ref cache);
            }
        }

        return new SlicerSearchResult { Cache = null, Slicer = null };
    }

    /// <summary>
    /// Builds a SlicerInfo from COM objects
    /// </summary>
    private static SlicerInfo BuildSlicerInfo(dynamic slicer, dynamic cache)
    {
        var info = new SlicerInfo
        {
            Name = slicer.Name?.ToString() ?? string.Empty,
            Caption = slicer.Caption?.ToString() ?? string.Empty,
            FieldName = GetSlicerCacheFieldName(cache),
            ColumnCount = Convert.ToInt32(slicer.NumberOfColumns)
        };

        // Get sheet name and position
        dynamic? parent = null;
        try
        {
            parent = slicer.Parent;
            info.SheetName = parent.Name?.ToString() ?? string.Empty;
        }
        finally
        {
            ComUtilities.Release(ref parent);
        }

        // Get position (top-left cell) - per Microsoft docs, TopLeftCell is on Shape object
        // https://learn.microsoft.com/en-us/office/vba/api/excel.shape.topleftcell
        dynamic? shape = null;
        dynamic? topLeftCell = null;
        try
        {
            shape = slicer.Shape;
            topLeftCell = shape.TopLeftCell;
            info.Position = topLeftCell?.Address?.ToString()?.Replace("$", "") ?? string.Empty;
        }
        finally
        {
            ComUtilities.Release(ref topLeftCell);
            ComUtilities.Release(ref shape);
        }

        // Get selected and available items from cache
        var items = GetSlicerItems(cache);
        info.SelectedItems = items.Selected;
        info.AvailableItems = items.Available;

        // Get connected PivotTables
        info.ConnectedPivotTables = GetConnectedPivotTableNames(cache);

        return info;
    }

    /// <summary>
    /// Builds a SlicerResult from COM objects
    /// </summary>
    private static SlicerResult BuildSlicerResult(dynamic slicer, dynamic cache, string fieldName)
    {
        var result = new SlicerResult
        {
            Name = slicer.Name?.ToString() ?? string.Empty,
            Caption = slicer.Caption?.ToString() ?? string.Empty,
            FieldName = fieldName
        };

        // Get sheet name and position
        dynamic? parent = null;
        try
        {
            parent = slicer.Parent;
            result.SheetName = parent.Name?.ToString() ?? string.Empty;
        }
        finally
        {
            ComUtilities.Release(ref parent);
        }

        // Get position - per Microsoft docs, TopLeftCell is on Shape object
        // https://learn.microsoft.com/en-us/office/vba/api/excel.shape.topleftcell
        dynamic? shape = null;
        dynamic? topLeftCell = null;
        try
        {
            shape = slicer.Shape;
            topLeftCell = shape.TopLeftCell;
            result.Position = topLeftCell?.Address?.ToString()?.Replace("$", "") ?? string.Empty;
        }
        finally
        {
            ComUtilities.Release(ref topLeftCell);
            ComUtilities.Release(ref shape);
        }

        // Get items
        var items = GetSlicerItems(cache);
        result.SelectedItems = items.Selected;
        result.AvailableItems = items.Available;

        // Get connected PivotTables
        result.ConnectedPivotTables = GetConnectedPivotTableNames(cache);

        return result;
    }

    /// <summary>
    /// Result of getting slicer items (avoids dynamic tuple deconstruction)
    /// </summary>
    private readonly struct SlicerItemsResult
    {
        public List<string> Selected { get; init; }
        public List<string> Available { get; init; }
    }

    /// <summary>
    /// Gets selected and available items from a SlicerCache
    /// </summary>
    private static SlicerItemsResult GetSlicerItems(dynamic cache)
    {
        var selected = new List<string>();
        var available = new List<string>();
        dynamic? slicerItems = null;

        try
        {
            slicerItems = cache.SlicerItems;

            for (int i = 1; i <= slicerItems.Count; i++)
            {
                dynamic? item = null;
                try
                {
                    item = slicerItems.Item(i);
                    string itemName = item.Name?.ToString() ?? string.Empty;

                    if (!string.IsNullOrEmpty(itemName))
                    {
                        available.Add(itemName);
                        if (item.Selected)
                        {
                            selected.Add(itemName);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref item);
                }
            }
        }
        catch (COMException)
        {
            // SlicerItems collection may not be accessible for certain cache types
        }
        finally
        {
            ComUtilities.Release(ref slicerItems);
        }

        return new SlicerItemsResult { Selected = selected, Available = available };
    }

    /// <summary>
    /// Gets names of PivotTables connected to a SlicerCache.
    /// Returns empty list for Table slicers (cache.List == true).
    /// </summary>
    private static List<string> GetConnectedPivotTableNames(dynamic cache)
    {
        var names = new List<string>();

        // Per MS docs: List property is true for Table slicers
        // Table slicers don't have PivotTables collection
        if (cache.List == true)
        {
            return names;
        }

        dynamic? pivotTables = null;
        try
        {
            pivotTables = cache.PivotTables;

            for (int i = 1; i <= pivotTables.Count; i++)
            {
                dynamic? pt = null;
                try
                {
                    pt = pivotTables.Item(i);
                    string name = pt.Name?.ToString() ?? string.Empty;
                    if (!string.IsNullOrEmpty(name))
                    {
                        names.Add(name);
                    }
                }
                finally
                {
                    ComUtilities.Release(ref pt);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref pivotTables);
        }

        return names;
    }

    /// <summary>
    /// Normalises a caller-supplied PivotTable name list into the distinct, non-blank names it
    /// actually asks for, preserving the caller's order and casing.
    ///
    /// Both connect-pivots and disconnect-pivots need this for the same reason: the request is a
    /// list, but the work is per distinct PivotTable. A repeated name would otherwise be applied
    /// twice (a second RemovePivotTable on an entry that is already gone raises a bare COM error) and
    /// would inflate the count checked against the "a slicer keeps at least one connection" rule.
    /// Returning false for a list with no usable entry keeps a meaningless request from being
    /// reported as completed.
    /// </summary>
    /// <param name="requested">Names supplied by the caller, possibly null, blank or repeated</param>
    /// <param name="cleaned">Distinct non-blank names in the caller's order</param>
    /// <returns>True when at least one usable name was supplied</returns>
    private static bool TryCleanPivotTableNames(List<string>? requested, out List<string> cleaned)
    {
        cleaned = new List<string>();

        if (requested == null || requested.Count == 0)
        {
            return false;
        }

        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (string name in requested)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                continue;
            }

            string trimmed = name.Trim();
            if (seen.Add(trimmed))
            {
                cleaned.Add(trimmed);
            }
        }

        return cleaned.Count > 0;
    }

    /// <summary>
    /// Reports whether a named PivotTable exists, without surfacing the lookup failure as an
    /// exception. Lets callers turn a typo into a proper failure result before mutating anything.
    /// </summary>
    private static bool PivotTableExists(dynamic workbook, string pivotTableName)
    {
        dynamic? pivot = null;

        try
        {
            pivot = FindPivotTable(workbook, pivotTableName);
            return true;
        }
        catch (InvalidOperationException)
        {
            return false;
        }
        finally
        {
            ComUtilities.Release(ref pivot);
        }
    }

    #endregion
}


