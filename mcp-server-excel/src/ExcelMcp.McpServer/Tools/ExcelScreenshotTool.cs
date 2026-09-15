using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands.Screenshot;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Manual MCP tool for screenshot operations.
/// Returns ImageContentBlock for proper MCP image handling.
/// </summary>
[McpServerToolType]
public static class ExcelScreenshotTool
{
    /// <summary>
    /// Capture Excel worksheet content as images for visual verification.
    /// Photographs the live Excel window, so the image shows exactly what Excel displays
    /// (formatting, charts, conditional formatting). Works on protected sheets and leaves the
    /// workbook and clipboard untouched, but requires an interactive desktop session.
    /// capture: specific range (requires rangeAddress).
    /// capture-sheet: entire used area of worksheet.
    /// Returns the image directly as MCP ImageContent.
    /// Use after operations to visually verify results.
    /// quality: Medium (default, JPEG 75% scale, ~4-8x smaller), High (PNG full scale), Low (JPEG 50% scale).
    /// </summary>
    [McpServerTool(Name = "screenshot", Title = "Screenshot", Destructive = false)]
    [McpMeta("category", "visualization")]
    [McpMeta("requiresSession", true)]
    [Description("Capture Excel worksheet content as images for visual verification. " +
        "Photographs the live Excel window, so the image shows exactly what Excel displays " +
        "(formatting, charts, conditional formatting). Works on protected sheets and leaves the " +
        "workbook and clipboard untouched, but requires an interactive desktop session. " +
        "capture: specific range (requires rangeAddress). " +
        "capture-sheet: entire used area of worksheet. " +
        "Returns the image directly as MCP ImageContent. " +
        "Use after operations to visually verify results. " +
        "quality: Medium (default, JPEG 75% scale, ~4-8x smaller than High), High (PNG full scale), Low (JPEG 50% scale).")]
    public static CallToolResult ExcelScreenshot(
        [Description("The action to perform")] ScreenshotAction action,
        [Description("Session ID from file 'open' action")] string session_id,
        [Description("Worksheet to capture. Omit to capture the active sheet. Required when the "
            + "session has never shown a sheet, because the active sheet is whatever Excel last "
            + "focused.")]
        [DefaultValue(null)] string? sheet_name,
        [Description("Range to capture for action='capture', for example 'A1:F20'. Ignored by "
            + "action='capture-sheet', which always captures the whole used area.")]
        [DefaultValue("A1:Z30")] string range_address,
        [Description("Image quality. Medium (default): JPEG at 75% scale, roughly 4-8x smaller than "
            + "High. High: PNG at full scale, use when the caller must read small text or fine "
            + "formatting. Low: JPEG at 50% scale.")]
        [DefaultValue(ScreenshotQuality.Medium)] ScreenshotQuality quality,
        CancellationToken cancellationToken = default)
    {
        using var cancellationScope = ExcelToolsBase.PushCancellationToken(cancellationToken);

        // Forward to service and get JSON response
        var jsonResponse = ExcelToolsBase.ExecuteToolAction(
            "screenshot",
            ServiceRegistry.Screenshot.ToActionString(action),
            () => RouteScreenshotAction(
                action,
                session_id,
                sheet_name,
                range_address,
                quality,
                ExcelToolsBase.ForwardToServiceFunc
            ));

        // Parse the JSON response to extract image data
        try
        {
            var result = JsonSerializer.Deserialize<ScreenshotResult>(jsonResponse, ExcelToolsBase.JsonOptions);

            if (result is null || !result.Success || string.IsNullOrEmpty(result.ImageBase64))
            {
                // Return error as text content
                return new CallToolResult
                {
                    IsError = true,
                    Content = [new TextContentBlock { Text = jsonResponse }]
                };
            }

            // Return image as ImageContentBlock + metadata as TextContentBlock
            var metadata = $"Screenshot: {result.RangeAddress} on '{result.SheetName}' ({result.Width}x{result.Height}px)";

            return new CallToolResult
            {
                Content =
                [
                    ImageContentBlock.FromBytes(Convert.FromBase64String(result.ImageBase64), result.MimeType),
                    new TextContentBlock
                    {
                        Text = metadata
                    }
                ]
            };
        }
        catch (JsonException)
        {
            // If JSON parsing fails, return the raw response as error
            return new CallToolResult
            {
                IsError = true,
                Content = [new TextContentBlock { Text = jsonResponse }]
            };
        }
    }

    internal static string RouteScreenshotAction(
        ScreenshotAction action,
        string sessionId,
        string? sheetName,
        string rangeAddress,
        ScreenshotQuality quality,
        Func<string, string, object?, string> forwardToService)
    {
        return ServiceRegistry.Screenshot.RouteAction(
            action,
            sessionId,
            forwardToService,
            sheetName: sheetName,
            rangeAddress: action == ScreenshotAction.CaptureRange ? rangeAddress : null,
            quality: quality);
    }
}
