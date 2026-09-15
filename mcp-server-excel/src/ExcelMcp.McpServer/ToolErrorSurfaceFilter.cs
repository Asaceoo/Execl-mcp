using System.Text;
using System.Text.Json;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.McpServer.Tools;

namespace Sbroenne.ExcelMcp.McpServer;

/// <summary>
/// Surfaces the real reason a tools/call failed.
///
/// The SDK's parameter binder validates arguments before any tool code runs. When it rejects a
/// call - unknown argument, missing required argument, value that does not fit the parameter type,
/// action outside the declared enum - the client only sees "An error occurred invoking 'x'.",
/// which makes the failure impossible to act on.
///
/// This filter catches those failures and returns the binder's own message plus the arguments the
/// tool actually accepts, so the caller can correct the call without guessing.
/// </summary>
internal static class ToolErrorSurfaceFilter
{
    internal static McpRequestHandler<CallToolRequestParams, CallToolResult> Wrap(
        McpRequestHandler<CallToolRequestParams, CallToolResult> next) =>
        async (request, cancellationToken) =>
        {
            try
            {
                return await next(request, cancellationToken).ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                throw;
            }
#pragma warning disable CA1031 // A tools/call failure must reach the caller as text, not as a broken pipe.
            catch (Exception ex)
#pragma warning restore CA1031
            {
                var tool = (request.MatchedPrimitive as McpServerTool)?.ProtocolTool;
                var detail = DescribeFailure(ex, request, tool);
                Console.Error.WriteLine($"[jyyj-mcp] tools/call rejected: {detail}");

                var payload = ExcelToolsBase.SerializeToolError(
                    tool?.Name ?? "tools/call",
                    null,
                    new ArgumentException(detail, ex));

                return new CallToolResult
                {
                    IsError = true,
                    Content = [new TextContentBlock { Text = payload }]
                };
            }
        };

    private static string DescribeFailure(
        Exception exception,
        RequestContext<CallToolRequestParams> request,
        Tool? tool)
    {
        var builder = new StringBuilder();
        builder.Append(tool is null
            ? "The tool call was rejected before it reached the tool."
            : $"Tool '{tool.Name}' rejected the arguments before executing.");
        builder.Append(' ');
        builder.Append(FlattenMessages(exception));

        // The shape remedy goes FIRST, ahead of the long "accepted arguments" list. It is the only
        // part that lets the caller fix the call in one round trip, and every consumer that caps
        // tool output - clients, terminals, log lines, this repo's own edge-case probe - keeps the
        // head of the string, not the tail. With the remedy last, a 400-character cap silently ate
        // exactly the sentence that mattered, which is how the defect was found.
        //
        // Naming the shape problem is what turns this class of failure into a one-shot fix: the
        // binder's own message reports "Path: $" and never says which argument, or which of the two
        // list wire formats the tool wanted.
        foreach (string remedy in ShapeMismatches(request, tool))
        {
            builder.Append(' ').Append(remedy);
        }

        var accepted = AcceptedArguments(tool);
        var unexpected = UnexpectedArguments(request, tool);
        if (unexpected.Count > 0)
        {
            builder.Append(" Unexpected argument(s): ").Append(string.Join(", ", unexpected)).Append('.');
        }

        if (accepted.Count > 0)
        {
            builder.Append(" Accepted argument(s): ").Append(string.Join(", ", accepted)).Append('.');
        }

        builder.Append(" Re-send the call with the accepted argument names.");
        return builder.ToString();
    }

    private static string FlattenMessages(Exception exception)
    {
        var messages = new List<string>();
        for (Exception? current = exception; current is not null; current = current.InnerException)
        {
            if (!string.IsNullOrWhiteSpace(current.Message) &&
                !messages.Exists(existing => string.Equals(existing, current.Message, StringComparison.Ordinal)))
            {
                messages.Add(current.Message.Trim());
            }
        }

        return messages.Count == 0 ? "No error detail was provided." : string.Join(" -> ", messages);
    }

    /// <summary>Argument names the tool declares in its input schema, in schema order.</summary>
    private static List<string> AcceptedArguments(Tool? tool)
    {
        var names = new List<string>();
        if (tool?.InputSchema.ValueKind != JsonValueKind.Object ||
            !tool.InputSchema.TryGetProperty("properties", out var properties) ||
            properties.ValueKind != JsonValueKind.Object)
        {
            return names;
        }

        foreach (var property in properties.EnumerateObject())
        {
            names.Add(property.Name);
        }

        return names;
    }

    /// <summary>Arguments the caller sent that the tool does not declare.</summary>
    private static List<string> UnexpectedArguments(
        RequestContext<CallToolRequestParams> request,
        Tool? tool)
    {
        var unexpected = new List<string>();
        if (request.Params?.Arguments is not { } arguments || tool is null)
        {
            return unexpected;
        }

        var declared = AcceptedArguments(tool);
        if (declared.Count == 0)
        {
            return unexpected;
        }

        var known = new HashSet<string>(declared, StringComparer.Ordinal);
        foreach (var argument in arguments.Keys)
        {
            if (!known.Contains(argument))
            {
                unexpected.Add(argument);
            }
        }

        return unexpected;
    }

    /// <summary>
    /// Arguments whose JSON shape contradicts the shape their tool declared.
    ///
    /// The surface uses two wire formats for "list of strings": some parameters take a JSON array
    /// inside a string (generated from <c>List&lt;string&gt;</c>) and others take a native JSON array
    /// (generated from <c>string[]</c>). The SDK binder reports a mismatch as "The JSON value could
    /// not be converted to System.String. Path: $" - which names neither the argument nor the fix.
    /// Comparing what arrived against what the schema declares produces a message the caller can act
    /// on without a second round trip.
    /// </summary>
    private static List<string> ShapeMismatches(
        RequestContext<CallToolRequestParams> request,
        Tool? tool)
    {
        var remedies = new List<string>();
        if (request.Params?.Arguments is not { } arguments || tool is null)
        {
            return remedies;
        }

        var declaredTypes = DeclaredTypes(tool);
        if (declaredTypes.Count == 0)
        {
            return remedies;
        }

        foreach (var argument in arguments)
        {
            if (!declaredTypes.TryGetValue(argument.Key, out var declaredType))
            {
                continue;
            }

            if (argument.Value.ValueKind == JsonValueKind.Array && declaredType == "string")
            {
                remedies.Add(
                    $"Argument '{argument.Key}' was sent as a JSON array, but this tool declares it as a "
                    + "string that carries a JSON array - send it quoted, for example "
                    + $"\"{argument.Key}\":\"[\\\"value1\\\",\\\"value2\\\"]\".");
            }
            else if (argument.Value.ValueKind == JsonValueKind.String && declaredType == "array")
            {
                remedies.Add(
                    $"Argument '{argument.Key}' was sent as a quoted string, but this tool declares it "
                    + "as a native JSON array - send it unquoted, for example "
                    + $"\"{argument.Key}\":[\"value1\",\"value2\"].");
            }
        }

        return remedies;
    }

    /// <summary>Declared JSON type of every argument the tool advertises, keyed by argument name.</summary>
    private static Dictionary<string, string> DeclaredTypes(Tool? tool)
    {
        var types = new Dictionary<string, string>(StringComparer.Ordinal);
        if (tool?.InputSchema.ValueKind != JsonValueKind.Object ||
            !tool.InputSchema.TryGetProperty("properties", out var properties) ||
            properties.ValueKind != JsonValueKind.Object)
        {
            return types;
        }

        foreach (var property in properties.EnumerateObject())
        {
            if (property.Value.ValueKind == JsonValueKind.Object &&
                property.Value.TryGetProperty("type", out var declaredType) &&
                declaredType.ValueKind == JsonValueKind.String)
            {
                types[property.Name] = declaredType.GetString() ?? string.Empty;
            }
        }

        return types;
    }
}
