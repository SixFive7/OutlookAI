using System.Text.Json;
using Xunit;

namespace OutlookAI.McpServer.Tests.T3;

/// <summary>
/// What the stdio suites read off the real server's tool surface (Q154): one tool's advertised entry, that a
/// parameter is optional, a tool call's raw result and the JSON payload inside it. The one copy of what the
/// tool-shape and supervision suites each kept, imported through <c>using static</c>. Nothing here names a
/// tool or declares mailbox contact: the tool names and the client's contact token stay in the calling
/// classes, where <c>T1/LiveTierInventoryTests</c> reads them out of the compiled code.
/// </summary>
internal static class ToolShapes
{
    /// <summary>The <c>tools/list</c> entry of the tool named <paramref name="name"/>, independent of the response it came in.</summary>
    public static async Task<JsonElement> GetToolAsync(McpStdioClient client, string name)
    {
        JsonElement list = await client.RoundTripAsync("tools/list", new { });
        foreach (JsonElement tool in list.GetProperty("result").GetProperty("tools").EnumerateArray())
        {
            if (tool.GetProperty("name").GetString() == name)
            {
                return tool.Clone();
            }
        }

        throw new InvalidOperationException($"tool '{name}' is not advertised");
    }

    /// <summary>That none of <paramref name="parameterNames"/> is in the schema's <c>required</c> list.</summary>
    public static void AssertOptional(JsonElement schema, string toolName, params string[] parameterNames)
    {
        if (!schema.TryGetProperty("required", out JsonElement required))
        {
            return;
        }

        string[] requiredNames = required.EnumerateArray().Select(r => r.GetString()!).ToArray();
        foreach (string parameterName in parameterNames)
        {
            Assert.DoesNotContain(parameterName, requiredNames, StringComparer.Ordinal);
        }
    }

    /// <summary>A tool call's <c>result</c>, an error result included, as the server sent it.</summary>
    public static async Task<JsonElement> CallRawAsync(McpStdioClient client, string tool, object arguments)
    {
        JsonElement envelope = await client.RoundTripAsync("tools/call", new { name = tool, arguments });
        return envelope.GetProperty("result");
    }

    /// <summary>The JSON payload a tool result carries as its first text content.</summary>
    public static JsonElement PayloadOf(JsonElement result)
    {
        string text = result.GetProperty("content")[0].GetProperty("text").GetString()!;
        return JsonDocument.Parse(text).RootElement;
    }
}
