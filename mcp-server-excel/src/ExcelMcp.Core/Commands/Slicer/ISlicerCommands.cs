using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Slicer;

/// <summary>
/// Slicer visual filters for PivotTables and Excel Tables.
///
/// PIVOTTABLE SLICERS: create-slicer, list-slicers, set-slicer-selection, delete-slicer,
/// connect-pivots, disconnect-pivots.
/// TABLE SLICERS: create-table-slicer, list-table-slicers, set-table-slicer-selection, delete-table-slicer.
///
/// NAMING: Auto-generate descriptive names like {FieldName}Slicer (e.g., RegionSlicer).
///
/// SELECTION: selectedItems as list of strings.
/// Empty list clears filter (shows all items). Set clearFirst=false to add to existing selection.
///
/// MULTI-TABLE FILTERING: one slicer can filter several PivotTables, but only when they share a
/// single PivotCache. Create them with pivottable 'share_cache_from', then use 'connect-pivots'.
/// </summary>
[ServiceCategory("slicer", "Slicer")]
[McpTool("slicer", Title = "Slicer Operations", Destructive = true, Category = "analysis",
    Description = "Slicer management: create, list, configure, connect, delete visual filtering controls for PivotTables and Tables. NAMING: Auto-generate descriptive names like RegionSlicer, CategorySlicer. PIVOTTABLE SLICERS: create-slicer, list-slicers, set-slicer-selection, delete-slicer, connect-pivots, disconnect-pivots. TABLE SLICERS: create-table-slicer, list-table-slicers, set-table-slicer-selection, delete-table-slicer. SELECTION: selectedItems as JSON array of strings. Use clearFirst=false to add to existing selection. MULTI-TABLE: one slicer filters several PivotTables only when they share one PivotCache - build them with pivottable share_cache_from, then use connect-pivots to add connections and disconnect-pivots to remove them; a slicer always keeps at least one connection, and list-slicers reports connectedPivotTables.")]
public interface ISlicerCommands
{
    /// <summary>
    /// Creates a slicer for a PivotTable field.
    /// Slicers provide visual filtering for PivotTable data.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable to create slicer for</param>
    /// <param name="fieldName">Name of the field to use for the slicer</param>
    /// <param name="slicerName">Name for the new slicer</param>
    /// <param name="destinationSheet">Worksheet where slicer will be placed</param>
    /// <param name="position">Top-left cell position for the slicer (e.g., "H2")</param>
    /// <returns>Created slicer details with available items</returns>
    [ServiceAction("create-slicer")]
    SlicerResult CreateSlicer(IExcelBatch batch, string pivotTableName,
        string fieldName, string slicerName, string destinationSheet, string position);

    /// <summary>
    /// Lists all slicers in the workbook, optionally filtered by PivotTable.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Optional PivotTable name to filter slicers (null = all slicers)</param>
    /// <returns>List of slicers with names, fields, positions, and selections</returns>
    [ServiceAction("list-slicers")]
    SlicerListResult ListSlicers(IExcelBatch batch, string? pivotTableName = null);

    /// <summary>
    /// Sets the selection for a slicer, filtering the connected PivotTable(s).
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="slicerName">Name of the slicer to modify</param>
    /// <param name="selectedItems">Items to select (show in PivotTable)</param>
    /// <param name="clearFirst">If true, clears existing selection before setting new items (default: true)</param>
    /// <returns>Updated slicer state with current selection</returns>
    [ServiceAction("set-slicer-selection")]
    SlicerResult SetSlicerSelection(IExcelBatch batch, string slicerName,
        List<string> selectedItems, bool clearFirst = true);

    /// <summary>
    /// Deletes a slicer from the workbook.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="slicerName">Name of the slicer to delete</param>
    /// <returns>Operation result indicating success or failure</returns>
    [ServiceAction("delete-slicer")]
    OperationResult DeleteSlicer(IExcelBatch batch, string slicerName);

    /// <summary>
    /// Connects additional PivotTables to a PivotTable slicer, so one slicer filters them all.
    /// Every target must share the slicer's PivotCache; create such PivotTables with the
    /// pivottable 'share_cache_from' option. Excel enforces the rule by raising a bare COM error
    /// from SlicerPivotTables.AddPivotTable without explaining why, so on failure this action
    /// re-reads the slicer's live connections and reports which PivotTables were refused, plus the
    /// share_cache_from remedy. Connecting itself never fails when the caches already match.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="slicerName">Name of the PivotTable slicer to extend</param>
    /// <param name="pivotTableNames">Names of the PivotTables to connect. Matched case-insensitively;
    /// blanks are ignored, and a request containing no usable name is rejected rather than reported
    /// as already done.</param>
    /// <returns>Slicer state with the full list of connected PivotTables</returns>
    [ServiceAction("connect-pivots")]
    SlicerResult ConnectPivotTables(IExcelBatch batch, string slicerName,
        List<string> pivotTableNames);

    /// <summary>
    /// Stops a slicer from filtering the given PivotTables. A slicer must keep at least one
    /// connection; removing the last one is rejected. The applied state is re-read from Excel
    /// before success is reported.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="slicerName">Name of the PivotTable slicer to shrink</param>
    /// <param name="pivotTableNames">Names of the PivotTables to disconnect. Matched
    /// case-insensitively and de-duplicated before counting, so repeating a name never consumes more
    /// than one connection.</param>
    /// <returns>Slicer state with the remaining connected PivotTables</returns>
    [ServiceAction("disconnect-pivots")]
    SlicerResult DisconnectPivotTables(IExcelBatch batch, string slicerName,
        List<string> pivotTableNames);

    /// <summary>
    /// Creates a slicer for an Excel Table column.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="tableName">Name of the Excel Table</param>
    /// <param name="columnName">Name of the column to use for the slicer</param>
    /// <param name="slicerName">Name for the new slicer</param>
    /// <param name="destinationSheet">Worksheet where slicer will be placed</param>
    /// <param name="position">Top-left cell position for the slicer (e.g., "H2")</param>
    /// <returns>Created slicer details with available items</returns>
    [ServiceAction("create-table-slicer")]
    SlicerResult CreateTableSlicer(IExcelBatch batch, string tableName,
        string columnName, string slicerName, string destinationSheet, string position);

    /// <summary>
    /// Lists all table slicers in the workbook, optionally filtered by table.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="tableName">Optional table name to filter slicers (null = all table slicers)</param>
    /// <returns>List of slicers with names, columns, positions, and selections</returns>
    [ServiceAction("list-table-slicers")]
    SlicerListResult ListTableSlicers(IExcelBatch batch, string? tableName = null);

    /// <summary>
    /// Sets the selection for a table slicer, filtering the connected table.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="slicerName">Name of the slicer to modify</param>
    /// <param name="selectedItems">Items to select (show in table)</param>
    /// <param name="clearFirst">If true, clears existing selection before setting new items (default: true)</param>
    /// <returns>Updated slicer state with current selection</returns>
    [ServiceAction("set-table-slicer-selection")]
    SlicerResult SetTableSlicerSelection(IExcelBatch batch, string slicerName,
        List<string> selectedItems, bool clearFirst = true);

    /// <summary>
    /// Deletes a table slicer from the workbook.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="slicerName">Name of the slicer to delete</param>
    /// <returns>Operation result indicating success or failure</returns>
    [ServiceAction("delete-table-slicer")]
    OperationResult DeleteTableSlicer(IExcelBatch batch, string slicerName);
}
