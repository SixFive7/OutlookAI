using System.Diagnostics;
using System.Text.Json;

namespace OutlookAI.McpServer.Tests;

/// <summary>
/// The raw MCP JSON-RPC wire over a server's stdio, hand-rolled rather than the SDK's so that what the stdio tests prove
/// is the protocol itself: one message a line, and a response found by its id among the notifications a server may send
/// before it. <c>T3/McpStdioClient</c> and <c>T3/McpStdioConformanceTests</c> each wrote it out (the code-sharing
/// refactor's S5b, the survey's T13).
/// </summary>
internal static class McpWire
{
    /// <summary>Writes <paramref name="message"/> as one line of JSON and flushes it.</summary>
    internal static async Task SendAsync(StreamWriter stdin, object message, CancellationToken ct)
    {
        string json = JsonSerializer.Serialize(message);
        await stdin.WriteLineAsync(json.AsMemory(), ct);
        await stdin.FlushAsync(ct);
    }

    /// <summary>
    /// The response to request <paramref name="id"/>, read a line at a time from <paramref name="stdout"/>: blank lines,
    /// notifications and messages with another id are skipped, as MCP allows. <paramref name="checkEnvelope"/>, when
    /// given, sees the response before its error is judged. Throws when stdout closes first - with what
    /// <paramref name="whenClosed"/> returns appended as the server's stderr, when given - and on a JSON-RPC error.
    /// </summary>
    internal static async Task<JsonElement> ReadResponseAsync(
        StreamReader stdout,
        int id,
        string method,
        CancellationToken ct,
        Func<Task<string>>? whenClosed = null,
        Action<JsonElement>? checkEnvelope = null)
    {
        while (true)
        {
            string? line = await stdout.ReadLineAsync(ct);
            if (line is null)
            {
                string stderr = whenClosed == null ? string.Empty : "\n--- server stderr ---\n" + await whenClosed();
                throw new InvalidOperationException($"Server closed stdout before answering '{method}' (id {id})." + stderr);
            }

            if (string.IsNullOrWhiteSpace(line))
            {
                continue;
            }

            JsonElement root;
            using (JsonDocument doc = JsonDocument.Parse(line))
            {
                root = doc.RootElement.Clone();
            }

            if (!root.TryGetProperty("id", out JsonElement idProp)
                || idProp.ValueKind != JsonValueKind.Number
                || idProp.GetInt32() != id)
            {
                continue; // notification or unrelated message
            }

            checkEnvelope?.Invoke(root);
            if (root.TryGetProperty("error", out JsonElement error))
            {
                throw new InvalidOperationException($"JSON-RPC error response for '{method}' (id {id}): {error.GetRawText()}");
            }

            return root;
        }
    }

    /// <summary>
    /// What the server wrote to stderr: the process tree is ended first when it still runs, and the read is waited for
    /// up to ten seconds. A drain that fails says so instead of throwing over a diagnostic.
    /// </summary>
    internal static async Task<string> DrainStderrAsync(Process server, Task<string> stderrTask)
    {
        try
        {
            if (!server.HasExited)
            {
                server.Kill(entireProcessTree: true);
            }

            return await stderrTask.WaitAsync(TimeSpan.FromSeconds(10));
        }
        catch (Exception drainEx)
        {
            return $"<stderr unavailable: {drainEx.GetType().Name}>";
        }
    }
}
