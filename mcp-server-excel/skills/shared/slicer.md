# slicer - Server Quirks

**Slicer Types**:

Two distinct slicer types exist:
- **PivotTable Slicers**: Filter PivotTables. One slicer can drive several PivotTables, provided
  they share a single PivotCache (see *Multi-PivotTable Linkage* below).
- **Table Slicers**: Filter Excel Tables (single table only).

**Actions**:

| Action | Description | Required Parameters |
|--------|-------------|---------------------|
| `create-slicer` | Create PivotTable slicer | pivot_table_name, field_name |
| `list-slicers` | List all PivotTable slicers | (none) |
| `set-slicer-selection` | Set PivotTable slicer filter | slicer_name, selected_items |
| `delete-slicer` | Delete PivotTable slicer | slicer_name |
| `connect-pivots` | Attach more PivotTables to one slicer | slicer_name, pivot_table_names |
| `disconnect-pivots` | Detach PivotTables from one slicer | slicer_name, pivot_table_names |
| `create-table-slicer` | Create Table slicer | table_name, column_name |
| `list-table-slicers` | List all Table slicers | (none) |
| `set-table-slicer-selection` | Set Table slicer filter | slicer_name, selected_items |
| `delete-table-slicer` | Delete Table slicer | slicer_name |

**CRITICAL: Required Parameters** - The "Required Parameters" column above is strict. Missing any required parameter will cause an error. Pay special attention to `pivot_table_name` (singular, for `create-slicer`) versus `pivot_table_names` (plural, for `connect-pivots` / `disconnect-pivots`), and to `slicer_name` for selection/deletion operations.

**Naming Convention**:

- If `slicer_name` is not provided, the tool auto-generates `{FieldName}Slicer` or `{ColumnName}Slicer`
- Slicer names must be unique within workbook
- Use `list-slicers` or `list-table-slicers` to check existing names

**Multi-PivotTable Linkage** (one slicer filtering several PivotTables):

Excel only lets a slicer drive PivotTables that share ONE PivotCache. It enforces that silently:
a violation comes back as a bare COM error naming nothing. The working sequence is:

1. `pivottable` `create` with `share_cache_from` set to a PivotTable the slicer already filters —
   repeat once per additional PivotTable. This is the only way to make two PivotTables share a cache.
2. `slicer` `create-slicer` on the first of them (`pivot_table_name`, `field_name`).
3. `slicer` `connect-pivots` with `slicer_name` and `pivot_table_names` to attach the rest.

`disconnect-pivots` reverses step 3. A slicer must always keep at least one connection, so removing
the last one is refused. `list-slicers` reports `connectedPivotTables` for every slicer, so the
applied state can be checked rather than assumed. Names are matched case-insensitively, blank
entries are ignored, and a request with no usable name is refused instead of reported as done.

**Selection Behavior**:

- `selected_items` carries a list of strings. The wire format differs by surface: the MCP tool
  takes it as a JSON array **inside a string** (`'["Value1", "Value2"]'`), while the CLI takes a
  plain JSON array argument. A native array sent to the MCP tool is rejected by the argument binder.
- Empty list `[]` clears all filters (shows all items)
- Values must match exactly (case-sensitive)
- Invalid values are silently ignored

**CLI: JSON Array Quoting** (important for `--selected-items`):

The `--selected-items` parameter requires a JSON array. Use proper shell escaping:

```powershell
# PowerShell: use single quotes around the JSON, double quotes inside
--selected-items '["West","East"]'

# Or escape inner quotes with backtick
--selected-items "[`"West`",`"East`"]"

# Clear filter (show all items)
--selected-items '[]'
```

**Positioning**:

- `destination_sheet` specifies which worksheet hosts the slicer
- `position` is a cell address for top-left corner (e.g., `'E1'`, `'G5'`)
- The slicer's top-left corner aligns to the specified cell
- Default position if not specified: Excel chooses

**Common Mistakes**:

- Creating slicer for field not in PivotTable → Error
- Creating table slicer for column not in table → Error
- Setting selection with wrong case → Values ignored (filter shows nothing)
- Deleting slicer that doesn't exist → Error
- Calling `connect-pivots` on PivotTables that do not share the slicer's PivotCache → the refused
  names are reported together with the connections that did stick; recreate them with
  `share_cache_from` and retry
- Sending a native JSON array to an MCP list parameter → argument-binder rejection; send a JSON
  array inside a string instead

**Best Practices**:

1. Call `list-slicers` before creating to avoid name conflicts
2. Use `list-slicers` to get exact slicer names for selection/deletion
3. For multi-PivotTable filtering use `share_cache_from` + `connect-pivots` (see above) — no manual
   work in the Excel UI is required
4. Styling: built-in slicer styles are applied through `vba` (`.xlsm` + `vba import` then
   `vba run`), setting `Slicer.Style` to one of `SlicerStyleLight1..6` / `SlicerStyleDark1..6`;
   all twelve were verified to apply and persist

**CLI Usage**:

```powershell
# Create PivotTable slicer
excelcli slicer create-slicer --session <id> --pivot-table-name "SalesPivot" --field-name "Region" --destination-sheet "Dashboard"

# Set slicer filter
excelcli slicer set-slicer-selection --session <id> --slicer-name "RegionSlicer" --selected-items "[`"West`",`"East`"]"

# Clear slicer filter (show all)
excelcli slicer set-slicer-selection --session <id> --slicer-name "RegionSlicer" --selected-items "[]"

# Attach a second PivotTable to the same slicer
excelcli slicer connect-pivots --session <id> --slicer-name "RegionSlicer" --pivot-table-names "[`"SalesPivot`",`"SalesPivot2`"]"

# Create Table slicer
excelcli slicer create-table-slicer --session <id> --table-name "SalesTable" --column-name "Category"

# List all slicers
excelcli slicer list-slicers --session <id>
excelcli slicer list-table-slicers --session <id>
```
