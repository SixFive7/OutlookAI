using System.Diagnostics;
using System.Text;
using System.Text.Json;

namespace OutlookAI.RemediationTools;

/// <summary>
/// Minimal raw-stdio MCP client for the telefonie refile (step 1): spawns the
/// REGISTERED server exe and speaks newline-delimited JSON-RPC (initialize,
/// notifications/initialized, tools/list, tools/call) - the T3 tier's proven wire
/// pattern. Every move therefore goes through the product's move_mail tool and gets
/// its load-bearing per-item audit line; this tool never moves mail via raw COM.
/// </summary>
public sealed class McpMoveClient : IDisposable
{
    private readonly Process _server;
    private readonly CancellationTokenSource _cts;
    private int _nextId;

    private McpMoveClient(Process server, CancellationTokenSource cts)
    {
        _server = server;
        _cts = cts;
    }

    /// <summary>Spawns the server exe and completes the MCP handshake.</summary>
    public static McpMoveClient StartAndInitialize(string serverExePath, TimeSpan? timeout = null)
    {
        if (!File.Exists(serverExePath))
        {
            throw new InvalidOperationException($"Server exe not found: {serverExePath}");
        }

        var cts = new CancellationTokenSource(timeout ?? TimeSpan.FromMinutes(5));
        var utf8NoBom = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
        var psi = new ProcessStartInfo(serverExePath)
        {
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true,
            WorkingDirectory = Path.GetDirectoryName(serverExePath)!,
            StandardInputEncoding = utf8NoBom,
            StandardOutputEncoding = utf8NoBom,
            StandardErrorEncoding = utf8NoBom,
        };

        Process server = Process.Start(psi)
            ?? throw new InvalidOperationException("Failed to start the MCP server process.");
        var client = new McpMoveClient(server, cts);
        JsonElement init = client.RoundTrip("initialize", new
        {
            protocolVersion = "2025-06-18",
            capabilities = new { },
            clientInfo = new { name = "OutlookAI.RemediationTools", version = "1.0.0" },
        });
        _ = init.GetProperty("result").GetProperty("protocolVersion");
        client.Notify("notifications/initialized");
        return client;
    }

    /// <summary>tools/list names (sanity: move_mail advertised, roster size).</summary>
    public IReadOnlyList<string> ListToolNames()
    {
        JsonElement list = RoundTrip("tools/list", new { });
        return list.GetProperty("result").GetProperty("tools").EnumerateArray()
            .Select(t => t.GetProperty("name").GetString()!)
            .ToList();
    }

    /// <summary>
    /// Calls a tool and returns its single JSON text content parsed. Throws on
    /// protocol isError; domain errors come back inside the JSON ({"error": ...}).
    /// </summary>
    public JsonElement CallTool(string name, object arguments)
    {
        JsonElement call = RoundTrip("tools/call", new { name, arguments });
        JsonElement result = call.GetProperty("result");
        if (result.TryGetProperty("isError", out JsonElement isError) && isError.ValueKind == JsonValueKind.True)
        {
            throw new InvalidOperationException($"tools/call {name} reported isError=true: {result.GetRawText()}");
        }

        string text = result.GetProperty("content").EnumerateArray()
            .First(c => c.GetProperty("type").GetString() == "text")
            .GetProperty("text").GetString()
            ?? throw new InvalidOperationException($"tools/call {name} returned no text content.");
        using JsonDocument doc = JsonDocument.Parse(text);
        return doc.RootElement.Clone();
    }

    /// <summary>Closes stdin and waits for the clean EOF exit (leak guard).</summary>
    public bool CloseAndAwaitExit(TimeSpan timeout)
    {
        _server.StandardInput.Close();
        return _server.WaitForExit((int)timeout.TotalMilliseconds);
    }

    private void Notify(string method)
    {
        Send(new { jsonrpc = "2.0", method });
    }

    private JsonElement RoundTrip(string method, object parameters)
    {
        int id = Interlocked.Increment(ref _nextId);
        Send(new { jsonrpc = "2.0", id, method, @params = parameters });
        while (true)
        {
            _cts.Token.ThrowIfCancellationRequested();
            string? line = _server.StandardOutput.ReadLine();
            if (line is null)
            {
                throw new InvalidOperationException($"Server closed stdout before answering '{method}' (id {id}).");
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
                continue;
            }

            if (root.TryGetProperty("error", out JsonElement error))
            {
                throw new InvalidOperationException($"JSON-RPC error for '{method}' (id {id}): {error.GetRawText()}");
            }

            return root;
        }
    }

    private void Send(object message)
    {
        string json = JsonSerializer.Serialize(message);
        _server.StandardInput.WriteLine(json);
        _server.StandardInput.Flush();
    }

    public void Dispose()
    {
        try
        {
            if (!_server.HasExited)
            {
                _server.Kill(entireProcessTree: true);
                _server.WaitForExit(10_000);
            }
        }
        catch (Exception)
        {
            // Teardown must not mask the operation's own outcome.
        }
        finally
        {
            _server.Dispose();
            _cts.Dispose();
        }
    }
}
