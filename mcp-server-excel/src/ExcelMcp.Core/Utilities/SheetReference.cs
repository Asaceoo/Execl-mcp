namespace Sbroenne.ExcelMcp.Core.Utilities;

/// <summary>
/// Builds A1-style references that survive every legal worksheet name.
///
/// Excel resolves a sheet-qualified reference only when the sheet name is wrapped in single quotes
/// and every apostrophe inside the name is doubled: a sheet called <c>O'Brien</c> is addressed as
/// <c>'O''Brien'!A1:D20</c>. Interpolating the raw name produces <c>'O'Brien'!A1:D20</c>, which Excel
/// rejects with a bare <c>0x800A03EC</c> - an error that names neither the sheet nor the cause.
///
/// The escaping rule was already applied by <c>DrawingCommands.Sparklines</c>; it lives here so the
/// next caller does not have to rediscover it.
/// </summary>
internal static class SheetReference
{
    /// <summary>Doubles embedded apostrophes so a sheet name can be quoted safely.</summary>
    /// <param name="sheetName">Raw worksheet name, for example <c>O'Brien</c></param>
    internal static string QuoteSheetName(string sheetName) =>
        sheetName.Replace("'", "''", StringComparison.Ordinal);

    /// <summary>
    /// Qualifies an A1-style address with a worksheet name, for example
    /// <c>BuildRangeReference("O'Brien", "A1:D20")</c> returns <c>'O''Brien'!A1:D20</c>.
    /// </summary>
    /// <param name="sheetName">Raw worksheet name</param>
    /// <param name="rangeAddress">A1-style address relative to that sheet</param>
    internal static string BuildRangeReference(string sheetName, string rangeAddress) =>
        $"'{QuoteSheetName(sheetName)}'!{rangeAddress}";
}
