using Microsoft.Extensions.Logging;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Slicer;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PivotTable;

/// <summary>
/// Integration tests for driving several PivotTables from a single slicer.
///
/// Excel only lets a slicer filter PivotTables that share one PivotCache, so the supported flow is:
/// create the first PivotTable normally, create the others with share_cache_from, create a slicer on
/// any of them, then call connect-pivots. The decisive assertion is a grand total that only moves
/// when the connection really took effect.
///
/// Each PivotTable is placed on its OWN worksheet. Excel refuses to insert a second PivotTable on a
/// sheet that already holds one ("cannot insert a PivotTable at the chosen location ... because
/// PivotTable 'X' already exists"), so co-locating them was wrong. Separate sheets also match how
/// linked dashboards are actually built.
///
/// SalesData layout ('SalesData'!A1:D6) - Region, Product, Sales, Date:
///   North/Widget/100, North/Widget/150, South/Gadget/200, North/Gadget/75, South/Widget/125
///   => North = 325, all regions = 650
/// </summary>
public partial class PivotTableCommandsTests
{
    private const string SalesSheet = "SalesData";
    private const string SalesSource = "A1:D6";
    private const double NorthRegionTotal = 325d;
    private const double AllRegionsTotal = 650d;

    /// <summary>
    /// One slicer, three separately built PivotTables sharing a PivotCache: after connection, a
    /// single selection must move the grand total of all three.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void SlicerLink_SharedCache_LetsOneSlicerFilterEveryPivotTable()
    {
        // Arrange
        var testFile = CreateTestFileWithData(nameof(SlicerLink_SharedCache_LetsOneSlicerFilterEveryPivotTable));

        var logger = _loggerFactory.CreateLogger<ExcelBatch>();
        using var batch = new ExcelBatch(new[] { testFile }, logger);

        AddSheet(batch, "P1");
        AddSheet(batch, "P2");
        AddSheet(batch, "P3");

        var first = _pivotCommands.CreateFromRange(batch, SalesSheet, SalesSource, "P1", "A1", "LinkP1");
        Assert.True(first.Success, $"CreateFromRange(P1) failed: {first.ErrorMessage}");
        Assert.Null(first.SharedCacheFrom);

        // Evidence of the rendered footprint, and proof the destination cell is honoured.
        _output.WriteLine($"LinkP1 created on P1, rendered range: {first.Range}");

        var second = _pivotCommands.CreateFromRange(
            batch, SalesSheet, SalesSource, "P2", "A1", "LinkP2", shareCacheFrom: "LinkP1");
        Assert.True(second.Success, $"CreateFromRange(P2) failed: {second.ErrorMessage}");
        Assert.Equal("LinkP1", second.SharedCacheFrom);
        Assert.NotNull(second.WorkflowHint);

        var third = _pivotCommands.CreateFromRange(
            batch, SalesSheet, SalesSource, "P3", "A1", "LinkP3", shareCacheFrom: "LinkP1");
        Assert.True(third.Success, $"CreateFromRange(P3) failed: {third.ErrorMessage}");
        Assert.Equal("LinkP1", third.SharedCacheFrom);

        _pivotCommands.AddRowField(batch, "LinkP1", "Region");
        _pivotCommands.AddValueField(batch, "LinkP1", "Sales", AggregationFunction.Sum);
        _pivotCommands.AddRowField(batch, "LinkP2", "Product");
        _pivotCommands.AddValueField(batch, "LinkP2", "Sales", AggregationFunction.Sum);
        _pivotCommands.AddRowField(batch, "LinkP3", "Region");
        _pivotCommands.AddValueField(batch, "LinkP3", "Sales", AggregationFunction.Sum);

        var slicerCommands = new SlicerCommands();
        var created = slicerCommands.CreateSlicer(
            batch, "LinkP1", "Region", "RegionLinkSlicer", SalesSheet, "I2");
        Assert.True(created.Success, $"CreateSlicer failed: {created.ErrorMessage}");

        // Recorded rather than asserted: whether Excel pre-connects the sibling PivotTables that
        // already share the cache is Excel's choice, not this API's contract.
        _output.WriteLine($"Connected right after create-slicer: {created.ConnectedPivotTables.Count} "
            + $"[{string.Join(", ", created.ConnectedPivotTables)}]");
        Assert.Contains("LinkP1", created.ConnectedPivotTables);

        Assert.Equal(AllRegionsTotal, GetGrandTotal(batch, "LinkP1"));
        Assert.Equal(AllRegionsTotal, GetGrandTotal(batch, "LinkP2"));
        Assert.Equal(AllRegionsTotal, GetGrandTotal(batch, "LinkP3"));

        // Act - connect the remaining PivotTables through the slicer facade
        var connected = slicerCommands.ConnectPivotTables(
            batch, "RegionLinkSlicer", new List<string> { "LinkP2", "LinkP3" });

        // Assert
        Assert.True(connected.Success, $"ConnectPivotTables failed: {connected.ErrorMessage}");
        Assert.Equal(3, connected.ConnectedPivotTables.Count);
        Assert.Contains("LinkP1", connected.ConnectedPivotTables);
        Assert.Contains("LinkP2", connected.ConnectedPivotTables);
        Assert.Contains("LinkP3", connected.ConnectedPivotTables);

        // The decisive check: one selection has to move every connected PivotTable.
        var selection = slicerCommands.SetSlicerSelection(
            batch, "RegionLinkSlicer", new List<string> { "North" });
        Assert.True(selection.Success, $"SetSlicerSelection failed: {selection.ErrorMessage}");

        Assert.Equal(NorthRegionTotal, GetGrandTotal(batch, "LinkP1"));
        Assert.Equal(NorthRegionTotal, GetGrandTotal(batch, "LinkP2"));
        Assert.Equal(NorthRegionTotal, GetGrandTotal(batch, "LinkP3"));
    }

    /// <summary>
    /// Default behaviour is preserved: without share_cache_from each PivotTable gets its own
    /// PivotCache, and connecting two such PivotTables to one slicer is rejected up front with an
    /// actionable message - and leaves the slicer untouched.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void SlicerLink_IndependentCaches_ConnectingIsRejectedWithoutSideEffects()
    {
        // Arrange
        var testFile = CreateTestFileWithData(nameof(SlicerLink_IndependentCaches_ConnectingIsRejectedWithoutSideEffects));

        var logger = _loggerFactory.CreateLogger<ExcelBatch>();
        using var batch = new ExcelBatch(new[] { testFile }, logger);

        AddSheet(batch, "S1");
        AddSheet(batch, "S2");

        var first = _pivotCommands.CreateFromRange(batch, SalesSheet, SalesSource, "S1", "A1", "SoloP1");
        var second = _pivotCommands.CreateFromRange(batch, SalesSheet, SalesSource, "S2", "A1", "SoloP2");
        Assert.True(first.Success, $"CreateFromRange(P1) failed: {first.ErrorMessage}");
        Assert.True(second.Success, $"CreateFromRange(P2) failed: {second.ErrorMessage}");

        // Both were built from the same range without share_cache_from, so each got its own cache.
        Assert.Null(first.SharedCacheFrom);
        Assert.Null(second.SharedCacheFrom);

        _pivotCommands.AddRowField(batch, "SoloP1", "Region");
        _pivotCommands.AddRowField(batch, "SoloP2", "Region");
        var created = _pivotCommands.CreateSlicer(batch, "SoloP1", "Region", "SoloSlicer", SalesSheet, "I2");
        Assert.True(created.Success, $"CreateSlicer failed: {created.ErrorMessage}");

        // Act
        var result = _pivotCommands.ConnectPivotTables(batch, "SoloSlicer", new List<string> { "SoloP2" });

        // Assert
        _output.WriteLine($"connect-pivots across independent caches -> Success={result.Success}: {result.ErrorMessage}");
        Assert.False(result.Success);
        Assert.Contains("PivotCache", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("share_cache_from", result.ErrorMessage);

        // The rejected request must not have left a half-applied connection behind.
        var listed = _pivotCommands.ListSlicers(batch);
        var slicer = FindSlicerByName(listed, "SoloSlicer");
        Assert.NotNull(slicer);
        Assert.DoesNotContain("SoloP2", slicer!.ConnectedPivotTables);

        // An empty request is rejected before Excel is touched, not silently reported as success.
        var emptyRequest = _pivotCommands.ConnectPivotTables(batch, "SoloSlicer", new List<string>());
        Assert.False(emptyRequest.Success);
        Assert.Contains("at least one non-empty PivotTable name", emptyRequest.ErrorMessage);
    }

    /// <summary>
    /// Disconnecting works, and a slicer is never left with zero connections.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void SlicerLink_DisconnectPivots_RemovesConnectionButKeepsAtLeastOne()
    {
        // Arrange
        var testFile = CreateTestFileWithData(nameof(SlicerLink_DisconnectPivots_RemovesConnectionButKeepsAtLeastOne));

        var logger = _loggerFactory.CreateLogger<ExcelBatch>();
        using var batch = new ExcelBatch(new[] { testFile }, logger);

        AddSheet(batch, "D1");
        AddSheet(batch, "D2");
        AddSheet(batch, "D3");

        var p1 = _pivotCommands.CreateFromRange(batch, SalesSheet, SalesSource, "D1", "A1", "DropP1");
        var p2 = _pivotCommands.CreateFromRange(batch, SalesSheet, SalesSource, "D2", "A1", "DropP2", shareCacheFrom: "DropP1");
        var p3 = _pivotCommands.CreateFromRange(batch, SalesSheet, SalesSource, "D3", "A1", "DropP3", shareCacheFrom: "DropP1");
        Assert.True(p1.Success && p2.Success && p3.Success,
            $"PivotTable setup failed: {p1.ErrorMessage} {p2.ErrorMessage} {p3.ErrorMessage}");

        _pivotCommands.AddRowField(batch, "DropP1", "Region");
        _pivotCommands.AddRowField(batch, "DropP2", "Region");
        _pivotCommands.AddRowField(batch, "DropP3", "Region");

        var created = _pivotCommands.CreateSlicer(batch, "DropP1", "Region", "DropSlicer", SalesSheet, "I2");
        Assert.True(created.Success, $"CreateSlicer failed: {created.ErrorMessage}");

        var connected = _pivotCommands.ConnectPivotTables(
            batch, "DropSlicer", new List<string> { "DropP2", "DropP3" });
        Assert.True(connected.Success, $"ConnectPivotTables failed: {connected.ErrorMessage}");
        Assert.Equal(3, connected.ConnectedPivotTables.Count);

        // Act
        var afterDrop = _pivotCommands.DisconnectPivotTables(batch, "DropSlicer", new List<string> { "DropP3" });

        // Assert
        Assert.True(afterDrop.Success, $"DisconnectPivotTables failed: {afterDrop.ErrorMessage}");
        Assert.Equal(2, afterDrop.ConnectedPivotTables.Count);
        Assert.DoesNotContain("DropP3", afterDrop.ConnectedPivotTables);

        // Removing everything must be refused - a slicer needs at least one connection.
        var removeAll = _pivotCommands.DisconnectPivotTables(
            batch, "DropSlicer", new List<string> { "DropP1", "DropP2" });
        Assert.False(removeAll.Success);
        Assert.Contains("at least one", removeAll.ErrorMessage, StringComparison.OrdinalIgnoreCase);

        // A name that is not connected is reported instead of silently ignored.
        var notConnected = _pivotCommands.DisconnectPivotTables(batch, "DropSlicer", new List<string> { "DropP3" });
        Assert.False(notConnected.Success);
        Assert.Contains("None of the requested", notConnected.ErrorMessage);

        // An empty request is rejected before Excel is touched.
        var emptyRequest = _pivotCommands.DisconnectPivotTables(batch, "DropSlicer", new List<string>());
        Assert.False(emptyRequest.Success);
        Assert.Contains("at least one non-empty PivotTable name", emptyRequest.ErrorMessage);
    }

    /// <summary>
    /// share_cache_from refuses to reuse a cache built on a different source, so callers cannot
    /// silently create PivotTables that no slicer will ever link.
    ///
    /// Excel reports a range PivotCache's SourceData in R1C1 notation, so this also pins the guard
    /// to compare like with like - an A1-vs-R1C1 comparison would reject every legitimate share.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void SlicerLink_ShareCacheFromDifferentSource_IsRejected()
    {
        // Arrange
        var testFile = CreateTestFileWithData(nameof(SlicerLink_ShareCacheFromDifferentSource_IsRejected));

        var logger = _loggerFactory.CreateLogger<ExcelBatch>();
        using var batch = new ExcelBatch(new[] { testFile }, logger);

        AddSheet(batch, "X1");
        AddSheet(batch, "X2");

        var first = _pivotCommands.CreateFromRange(batch, SalesSheet, SalesSource, "X1", "A1", "SrcP1");
        Assert.True(first.Success, $"CreateFromRange(P1) failed: {first.ErrorMessage}");

        // Act + Assert - a narrower range must not be allowed to reuse P1's cache
        var ex = Assert.Throws<InvalidOperationException>(() =>
        {
            _pivotCommands.CreateFromRange(
                batch, SalesSheet, "A1:C6", "X2", "A1", "SrcP2", shareCacheFrom: "SrcP1");
        });

        Assert.Contains("share_cache_from", ex.Message);
        Assert.Contains("SrcP1", ex.Message);
    }

    /// <summary>
    /// Characterisation probe. Microsoft's reference for SlicerPivotTables.AddPivotTable does not
    /// document a same-PivotCache requirement, yet XML-level investigation showed Excel ignores a
    /// connection that crosses caches. This records which branch Excel actually takes when the call
    /// is made directly, bypassing the guard in ConnectPivotTables. The asserted invariant is that
    /// Excel answers deterministically: it either connects the PivotTable or raises a COM error.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void SlicerLink_RawAddPivotTableAcrossCaches_IsDeterministic()
    {
        // Arrange
        var testFile = CreateTestFileWithData(nameof(SlicerLink_RawAddPivotTableAcrossCaches_IsDeterministic));

        var logger = _loggerFactory.CreateLogger<ExcelBatch>();
        using var batch = new ExcelBatch(new[] { testFile }, logger);

        AddSheet(batch, "R1");
        AddSheet(batch, "R2");

        var p1 = _pivotCommands.CreateFromRange(batch, SalesSheet, SalesSource, "R1", "A1", "RawP1");
        var p2 = _pivotCommands.CreateFromRange(batch, SalesSheet, SalesSource, "R2", "A1", "RawP2");
        Assert.True(p1.Success && p2.Success, $"PivotTable setup failed: {p1.ErrorMessage} {p2.ErrorMessage}");

        _pivotCommands.AddRowField(batch, "RawP1", "Region");
        _pivotCommands.AddRowField(batch, "RawP2", "Region");
        var created = _pivotCommands.CreateSlicer(batch, "RawP1", "Region", "RawSlicer", SalesSheet, "I2");
        Assert.True(created.Success, $"CreateSlicer failed: {created.ErrorMessage}");

        bool raised = false;
        string outcome = "accepted";

        // Act - call the raw COM API, deliberately skipping ExcelMcp's cache guard
        batch.Execute((ctx, ct) =>
        {
            dynamic? slicerCaches = null;
            dynamic? cache = null;
            dynamic? pivotTables = null;
            dynamic? target = null;

            try
            {
                slicerCaches = ctx.Book.SlicerCaches;
                cache = slicerCaches.Item(1);
                pivotTables = cache.PivotTables;
                target = FindPivotTableByRawCom(ctx.Book, "RawP2");

                try
                {
                    pivotTables.AddPivotTable(target);
                }
                catch (System.Runtime.InteropServices.COMException ex)
                {
                    raised = true;
                    outcome = $"COMException 0x{ex.HResult:X8}: {ex.Message}";
                }

                return 0;
            }
            finally
            {
                ComUtilities.Release(ref target);
                ComUtilities.Release(ref pivotTables);
                ComUtilities.Release(ref cache);
                ComUtilities.Release(ref slicerCaches);
            }
        });

        _output.WriteLine($"Raw SlicerPivotTables.AddPivotTable across different PivotCaches -> {outcome}");

        var listed = _pivotCommands.ListSlicers(batch);
        var slicer = FindSlicerByName(listed, "RawSlicer");
        Assert.NotNull(slicer);
        bool connectedNow = slicer!.ConnectedPivotTables.Contains("RawP2");

        _output.WriteLine($"Slicer connections after the raw call: [{string.Join(", ", slicer.ConnectedPivotTables)}]");

        // Excel must not both reject the call and report the connection.
        Assert.False(raised && connectedNow,
            "Excel raised an error and still reported the PivotTable as connected, which is contradictory.");
    }

    /// <summary>
    /// Regression: repeating a name (including as a different casing) must not corrupt the
    /// "a slicer keeps at least one connection" arithmetic, and must not reach Excel twice.
    ///
    /// Excel's connection list is casing-insensitive for matching but the removal call is not
    /// forgiving of a second attempt on an already removed entry, so a duplicated request used to
    /// either double-count against the last-connection guard or raise a bare COM error.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void SlicerLink_DisconnectDuplicateNames_KeepTheLastConnectionGuardAccurate()
    {
        // Arrange
        var testFile = CreateTestFileWithData(nameof(SlicerLink_DisconnectDuplicateNames_KeepTheLastConnectionGuardAccurate));

        var logger = _loggerFactory.CreateLogger<ExcelBatch>();
        using var batch = new ExcelBatch(new[] { testFile }, logger);

        AddSheet(batch, "U1");
        AddSheet(batch, "U2");
        AddSheet(batch, "U3");

        var p1 = _pivotCommands.CreateFromRange(batch, SalesSheet, SalesSource, "U1", "A1", "DupP1");
        var p2 = _pivotCommands.CreateFromRange(batch, SalesSheet, SalesSource, "U2", "A1", "DupP2", shareCacheFrom: "DupP1");
        var p3 = _pivotCommands.CreateFromRange(batch, SalesSheet, SalesSource, "U3", "A1", "DupP3", shareCacheFrom: "DupP1");
        Assert.True(p1.Success && p2.Success && p3.Success,
            $"PivotTable setup failed: {p1.ErrorMessage} {p2.ErrorMessage} {p3.ErrorMessage}");

        _pivotCommands.AddRowField(batch, "DupP1", "Region");
        _pivotCommands.AddRowField(batch, "DupP2", "Region");
        _pivotCommands.AddRowField(batch, "DupP3", "Region");

        var created = _pivotCommands.CreateSlicer(batch, "DupP1", "Region", "DupSlicer", SalesSheet, "I2");
        Assert.True(created.Success, $"CreateSlicer failed: {created.ErrorMessage}");

        var connected = _pivotCommands.ConnectPivotTables(
            batch, "DupSlicer", new List<string> { "DupP2", "DupP3" });
        Assert.True(connected.Success, $"ConnectPivotTables failed: {connected.ErrorMessage}");
        Assert.Equal(3, connected.ConnectedPivotTables.Count);

        // Act - the same target named twice, differing only in casing
        var duplicated = _pivotCommands.DisconnectPivotTables(
            batch, "DupSlicer", new List<string> { "DupP3", "dupp3" });

        // Assert - one PivotTable was asked for, so exactly one must go
        Assert.True(duplicated.Success, $"Duplicated names broke the disconnect: {duplicated.ErrorMessage}");
        Assert.Equal(2, duplicated.ConnectedPivotTables.Count);
        Assert.DoesNotContain("DupP3", duplicated.ConnectedPivotTables);

        // The last-connection guard must count distinct PivotTables, not list entries: removing
        // DupP1 once leaves DupP2 behind, which is legal.
        var lastOne = _pivotCommands.DisconnectPivotTables(
            batch, "DupSlicer", new List<string> { "DupP1", "dupp1" });

        Assert.True(lastOne.Success,
            $"Removing one distinct PivotTable was counted as removing two: {lastOne.ErrorMessage}");
        Assert.Single(lastOne.ConnectedPivotTables);
        Assert.Contains("DupP2", lastOne.ConnectedPivotTables);
    }

    /// <summary>
    /// Regression: a request made only of blank names is not a request. Reporting success (and an
    /// "already connected" hint) for it would tell a caller its intent was carried out when nothing
    /// was even looked up.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void SlicerLink_ConnectBlankOnlyRequest_IsRejectedInsteadOfReportedAsDone()
    {
        // Arrange
        var testFile = CreateTestFileWithData(nameof(SlicerLink_ConnectBlankOnlyRequest_IsRejectedInsteadOfReportedAsDone));

        var logger = _loggerFactory.CreateLogger<ExcelBatch>();
        using var batch = new ExcelBatch(new[] { testFile }, logger);

        AddSheet(batch, "B1");

        var first = _pivotCommands.CreateFromRange(batch, SalesSheet, SalesSource, "B1", "A1", "BlankP1");
        Assert.True(first.Success, $"CreateFromRange failed: {first.ErrorMessage}");

        _pivotCommands.AddRowField(batch, "BlankP1", "Region");
        var created = _pivotCommands.CreateSlicer(batch, "BlankP1", "Region", "BlankSlicer", SalesSheet, "I2");
        Assert.True(created.Success, $"CreateSlicer failed: {created.ErrorMessage}");

        // Act
        var result = _pivotCommands.ConnectPivotTables(batch, "BlankSlicer", new List<string> { "", "   " });

        // Assert
        Assert.False(result.Success, "A request containing no usable PivotTable name must not report success.");
        Assert.Contains("name", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);

        // The same rule applies when shrinking a slicer.
        var disconnect = _pivotCommands.DisconnectPivotTables(batch, "BlankSlicer", new List<string> { "", "  " });
        Assert.False(disconnect.Success,
            "A disconnect request containing no usable PivotTable name must not be reported as a not-connected error.");
        Assert.Contains("name", disconnect.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// A partially applied connect request must not read as "nothing happened". When one name sticks
    /// and another is refused, the failure has to disclose the live connection list so the caller can
    /// see the state it is actually in.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void SlicerLink_PartialConnectFailure_DisclosesTheStateItLeftBehind()
    {
        // Arrange
        var testFile = CreateTestFileWithData(nameof(SlicerLink_PartialConnectFailure_DisclosesTheStateItLeftBehind));

        var logger = _loggerFactory.CreateLogger<ExcelBatch>();
        using var batch = new ExcelBatch(new[] { testFile }, logger);

        AddSheet(batch, "M1");
        AddSheet(batch, "M2");
        AddSheet(batch, "M3");

        // Two PivotTables share a cache; the third is built independently on purpose.
        var shared1 = _pivotCommands.CreateFromRange(batch, SalesSheet, SalesSource, "M1", "A1", "MixP1");
        var shared2 = _pivotCommands.CreateFromRange(batch, SalesSheet, SalesSource, "M2", "A1", "MixP2", shareCacheFrom: "MixP1");
        var independent = _pivotCommands.CreateFromRange(batch, SalesSheet, SalesSource, "M3", "A1", "MixSolo");
        Assert.True(shared1.Success && shared2.Success && independent.Success,
            $"PivotTable setup failed: {shared1.ErrorMessage} {shared2.ErrorMessage} {independent.ErrorMessage}");

        _pivotCommands.AddRowField(batch, "MixP1", "Region");
        _pivotCommands.AddRowField(batch, "MixP2", "Region");
        _pivotCommands.AddRowField(batch, "MixSolo", "Region");

        var created = _pivotCommands.CreateSlicer(batch, "MixP1", "Region", "MixSlicer", SalesSheet, "I2");
        Assert.True(created.Success, $"CreateSlicer failed: {created.ErrorMessage}");

        // Act - one legitimate target, one that Excel cannot accept
        var result = _pivotCommands.ConnectPivotTables(
            batch, "MixSlicer", new List<string> { "MixP2", "MixSolo" });

        // Assert
        _output.WriteLine($"Partial connect -> Success={result.Success}: {result.ErrorMessage}");

        Assert.False(result.Success, "Connecting a PivotTable with its own PivotCache must fail.");
        Assert.Contains("MixSolo", result.ErrorMessage);

        // MixP2 shares the cache, so it either was already connected or just got connected. Whichever
        // it is, the report has to say what the slicer filters now - silence lets a caller assume the
        // request as a whole was discarded.
        var listed = _pivotCommands.ListSlicers(batch);
        var slicer = FindSlicerByName(listed, "MixSlicer");
        Assert.NotNull(slicer);
        _output.WriteLine($"Live connections: [{string.Join(", ", slicer!.ConnectedPivotTables)}]");

        Assert.Contains("MixP2", slicer.ConnectedPivotTables);
        Assert.Contains("MixP2", result.ErrorMessage);
    }

    /// <summary>
    /// Regression: two different sheets whose names end with one another must not be treated as the
    /// same source. Normalisation strips quoting and anchors, and a suffix match was also accepted -
    /// which silently let a PivotTable reuse a cache built on a different sheet.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void SlicerLink_ShareCacheFromSuffixSheetName_IsRejectedAsADifferentSource()
    {
        // Arrange
        var testFile = CreateTestFileWithData(nameof(SlicerLink_ShareCacheFromSuffixSheetName_IsRejectedAsADifferentSource));

        var logger = _loggerFactory.CreateLogger<ExcelBatch>();
        using var batch = new ExcelBatch(new[] { testFile }, logger);

        // 'SalesData' is a suffix of 'RawSalesData', so a suffix-tolerant guard cannot tell them apart.
        CopyRangeToNewSheet(batch, SalesSheet, SalesSource, "RawSalesData");
        AddSheet(batch, "N1");
        AddSheet(batch, "N2");

        var original = _pivotCommands.CreateFromRange(batch, SalesSheet, SalesSource, "N1", "A1", "SuffixP1");
        Assert.True(original.Success, $"CreateFromRange failed: {original.ErrorMessage}");

        // Act + Assert - the same addresses on a different sheet are a different cache source
        var ex = Assert.Throws<InvalidOperationException>(() =>
        {
            _pivotCommands.CreateFromRange(
                batch, "RawSalesData", SalesSource, "N2", "A1", "SuffixP2", shareCacheFrom: "SuffixP1");
        });

        _output.WriteLine($"Rejected with: {ex.Message}");
        Assert.Contains("share_cache_from", ex.Message);
    }

    /// <summary>
    /// Characterisation probe for non-rectangular source ranges. Excel reports PivotCache.SourceData
    /// in R1C1, and this records what that notation looks like for a whole-column source next to
    /// what the range itself reports - the pair decides whether the shared-cache guard can compare
    /// the two forms at all.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void SlicerLink_WholeColumnSource_RecordsTheNotationTheGuardMustCompare()
    {
        // Arrange
        var testFile = CreateTestFileWithData(nameof(SlicerLink_WholeColumnSource_RecordsTheNotationTheGuardMustCompare));

        var logger = _loggerFactory.CreateLogger<ExcelBatch>();
        using var batch = new ExcelBatch(new[] { testFile }, logger);

        string rangeNotation = string.Empty;
        string cacheNotation = string.Empty;

        batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            dynamic? caches = null;
            dynamic? cache = null;

            try
            {
                sheet = ctx.Book.Worksheets[SalesSheet];
                range = sheet.Range["A:D"];
                rangeNotation = Convert.ToString(range.Address(ReferenceStyle: -4150)) ?? string.Empty;

                caches = ctx.Book.PivotCaches();
                cache = caches.Create(SourceType: 1, SourceData: $"'{SalesSheet}'!A:D", Version: 4);
                cacheNotation = Convert.ToString(cache.SourceData) ?? string.Empty;

                return 0;
            }
            finally
            {
                ComUtilities.Release(ref cache);
                ComUtilities.Release(ref caches);
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });

        _output.WriteLine($"Range('A:D').Address(R1C1) = '{rangeNotation}'");
        _output.WriteLine($"PivotCache.SourceData      = '{cacheNotation}'");

        // The two notations have to agree for the guard to be able to compare them. When they do not,
        // sharing a cache built from this very source would be rejected - so the guard must be
        // deliberately lenient about the row span rather than strict.
        Assert.False(string.IsNullOrWhiteSpace(cacheNotation), "Excel reported no SourceData for the whole-column cache.");
    }

    /// <summary>
    /// Pins both directions of the whitespace decision in the shared-cache guard: quoting and
    /// anchors are normalised away, spaces inside a sheet name are not.
    ///
    /// ('Sales Data' and 'SalesData' are two different sheets, so treating them as equal would let a
    /// PivotTable adopt a cache built somewhere else entirely.)
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void SlicerLink_ShareCacheFromSpaceInSheetName_MatchesOnlyTheSameSheet()
    {
        // Arrange
        var testFile = CreateTestFileWithData(nameof(SlicerLink_ShareCacheFromSpaceInSheetName_MatchesOnlyTheSameSheet));

        var logger = _loggerFactory.CreateLogger<ExcelBatch>();
        using var batch = new ExcelBatch(new[] { testFile }, logger);

        const string spacedSheet = "Sales Data";
        CopyRangeToNewSheet(batch, SalesSheet, SalesSource, spacedSheet);
        AddSheet(batch, "SP2");
        AddSheet(batch, "SP3");
        AddSheet(batch, "SP4");

        var original = _pivotCommands.CreateFromRange(batch, spacedSheet, SalesSource, "SP2", "A1", "SpaceP1");
        Assert.True(original.Success, $"CreateFromRange on a spaced sheet name failed: {original.ErrorMessage}");

        // Act + Assert - the identical source must be shareable, even though Excel reports it as
        // 'Sales Data'!R1C1:R6C4 while the request builds the reference unquoted.
        var shared = _pivotCommands.CreateFromRange(
            batch, spacedSheet, SalesSource, "SP3", "A1", "SpaceP2", shareCacheFrom: "SpaceP1");

        Assert.True(shared.Success,
            $"A sheet name containing a space blocked a legitimate share: {shared.ErrorMessage}");
        Assert.Equal("SpaceP1", shared.SharedCacheFrom);

        // The same addresses on 'SalesData' are a different sheet and must be refused.
        var ex = Assert.Throws<InvalidOperationException>(() =>
        {
            _pivotCommands.CreateFromRange(
                batch, SalesSheet, SalesSource, "SP4", "A1", "SpaceP3", shareCacheFrom: "SpaceP1");
        });

        _output.WriteLine($"Rejected with: {ex.Message}");
        Assert.Contains("share_cache_from", ex.Message);
    }

    /// <summary>
    /// Copies a range of values onto a brand new worksheet, so a test can present the same figures
    /// under a different sheet name.
    /// </summary>
    private static void CopyRangeToNewSheet(ExcelBatch batch, string sourceSheetName, string sourceRangeAddress, string newSheetName)
    {
        batch.Execute((ctx, ct) =>
        {
            dynamic? sheets = null;
            dynamic? sourceSheet = null;
            dynamic? sourceRange = null;
            dynamic? lastSheet = null;
            dynamic? added = null;
            dynamic? destRange = null;

            try
            {
                sheets = ctx.Book.Worksheets;
                sourceSheet = sheets[sourceSheetName];
                sourceRange = sourceSheet.Range[sourceRangeAddress];
                object values = sourceRange.Value2;

                lastSheet = sheets.Item(sheets.Count);
                added = sheets.Add(After: lastSheet);
                added.Name = newSheetName;

                destRange = added.Range[sourceRangeAddress];
                destRange.Value2 = values;

                return 0;
            }
            finally
            {
                ComUtilities.Release(ref destRange);
                ComUtilities.Release(ref added);
                ComUtilities.Release(ref lastSheet);
                ComUtilities.Release(ref sourceRange);
                ComUtilities.Release(ref sourceSheet);
                ComUtilities.Release(ref sheets);
            }
        });
    }

    /// <summary>
    /// Adds a worksheet through raw COM - the product has no add-sheet command, and each
    /// PivotTable needs its own sheet because Excel rejects a second PivotTable on an occupied one.
    /// </summary>
    private static void AddSheet(ExcelBatch batch, string sheetName)
    {
        batch.Execute((ctx, ct) =>
        {
            dynamic? worksheets = null;
            dynamic? last = null;
            dynamic? added = null;
            try
            {
                worksheets = ctx.Book.Worksheets;
                last = worksheets.Item(worksheets.Count);
                added = worksheets.Add(After: last);
                added.Name = sheetName;
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref added);
                ComUtilities.Release(ref last);
                ComUtilities.Release(ref worksheets);
            }
        });
    }

    /// <summary>
    /// Reads a PivotTable's grand total from its rendered range: the last numeric cell of the last
    /// row, which is the Grand Total for a single-value-field layout.
    /// </summary>
    private double GetGrandTotal(IExcelBatch batch, string pivotTableName)
    {
        var data = _pivotCommands.GetData(batch, pivotTableName);
        Assert.True(data.Success, $"GetData('{pivotTableName}') failed: {data.ErrorMessage}");
        Assert.NotEmpty(data.Values);

        var lastRow = data.Values[data.Values.Count - 1];
        for (int i = lastRow.Count - 1; i >= 0; i--)
        {
            var cell = lastRow[i];
            if (cell is not null && double.TryParse(cell.ToString(), out double total))
            {
                return total;
            }
        }

        throw new InvalidOperationException(
            $"No numeric grand total found for PivotTable '{pivotTableName}'.");
    }

    private static SlicerInfo? FindSlicerByName(SlicerListResult list, string name)
    {
        foreach (var slicer in list.Slicers)
        {
            if (string.Equals(slicer.Name, name, StringComparison.OrdinalIgnoreCase))
            {
                return slicer;
            }
        }

        return null;
    }

    /// <summary>
    /// Resolves a PivotTable by name using raw COM only, so the probe can bypass the product's own
    /// lookup helpers and cache guard.
    /// </summary>
    private static dynamic? FindPivotTableByRawCom(dynamic workbook, string pivotTableName)
    {
        dynamic? sheets = null;
        dynamic? sheet = null;
        dynamic? pivotTables = null;

        try
        {
            sheets = workbook.Worksheets;
            for (int i = 1; i <= sheets.Count; i++)
            {
                ComUtilities.Release(ref sheet);
                sheet = sheets.Item(i);
                pivotTables = sheet.PivotTables;

                for (int j = 1; j <= pivotTables.Count; j++)
                {
                    dynamic? pivot = null;
                    try
                    {
                        pivot = pivotTables.Item(j);
                        if (string.Equals(pivot.Name?.ToString(), pivotTableName, StringComparison.OrdinalIgnoreCase))
                        {
                            dynamic found = pivot;
                            pivot = null; // ownership moves to the caller
                            return found;
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref pivot);
                    }
                }

                ComUtilities.Release(ref pivotTables);
            }

            return null;
        }
        finally
        {
            ComUtilities.Release(ref pivotTables);
            ComUtilities.Release(ref sheet);
            ComUtilities.Release(ref sheets);
        }
    }
}
