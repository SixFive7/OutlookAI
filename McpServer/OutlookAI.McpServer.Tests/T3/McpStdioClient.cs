using System.Diagnostics;
using System.Reflection;
using System.Text;
using System.Text.Json;

namespace OutlookAI.McpServer.Tests.T3;

/// <summary>
/// Hand-rolled stdio MCP client for the T3 tier (v3.MD section 0.6): spawns the built
/// server exe and speaks raw newline-delimited JSON-RPC - initialize,
/// notifications/initialized, tools/list, tools/call - proving the wire protocol
/// independent of the ModelContextProtocol package. Also the documented fallback
/// skeleton should the SDK ever be replaced (D31).
/// </summary>
public sealed class McpStdioClient : IAsyncDisposable
{
    private readonly Process _server;
    private readonly CancellationTokenSource _cts;
    private readonly Task<string> _stderrTask;
    private int _nextId;

    private McpStdioClient(Process server, CancellationTokenSource cts, Task<string> stderrTask)
    {
        _server = server;
        _cts = cts;
        _stderrTask = stderrTask;
    }

    /// <summary>Path of the built server exe (baked into the test assembly at build time).</summary>
    public static string ServerExePath =>
        typeof(McpStdioClient).Assembly
            .GetCustomAttributes<AssemblyMetadataAttribute>()
            .FirstOrDefault(a => a.Key == "McpServerExePath")?.Value
        ?? throw new InvalidOperationException("AssemblyMetadata 'McpServerExePath' is missing.");

    /// <summary>Spawns the server and completes the initialize handshake.</summary>
    /// <param name="timeout">Overall client timeout.</param>
    /// <param name="environment">
    /// Extra environment variables for the server process. Used by the supervision tests
    /// to inject a COM-host fault and shorten the deadline: the timeout/kill/respawn path
    /// is only observable by actually exceeding a budget, and waiting out the real
    /// two-minute one in every such test would make the suite unusable.
    /// </param>
    public static async Task<McpStdioClient> StartAndInitializeAsync(
        TimeSpan? timeout = null,
        IReadOnlyDictionary<string, string>? environment = null)
    {
        string exePath = ServerExePath;
        if (!File.Exists(exePath))
        {
            throw new InvalidOperationException($"Server exe not found at '{exePath}' - build OutlookAI.McpServer first.");
        }

        var cts = new CancellationTokenSource(timeout ?? TimeSpan.FromSeconds(180));
        var utf8NoBom = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
        var psi = new ProcessStartInfo(exePath)
        {
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true,
            WorkingDirectory = Path.GetDirectoryName(exePath)!,
            StandardInputEncoding = utf8NoBom,
            StandardOutputEncoding = utf8NoBom,
            StandardErrorEncoding = utf8NoBom,
        };

        if (environment != null)
        {
            foreach (KeyValuePair<string, string> entry in environment)
            {
                psi.Environment[entry.Key] = entry.Value;
            }
        }

        Process server = Process.Start(psi)
            ?? throw new InvalidOperationException("Failed to start the server process.");
        Task<string> stderrTask = server.StandardError.ReadToEndAsync(cts.Token);
        var client = new McpStdioClient(server, cts, stderrTask);

        JsonElement init = await client.RoundTripAsync("initialize", new
        {
            protocolVersion = "2025-06-18",
            capabilities = new { },
            clientInfo = new { name = "OutlookAI.T3Client", version = "0.0.2" },
        });
        _ = init.GetProperty("result").GetProperty("protocolVersion");
        await client.NotifyAsync("notifications/initialized");
        return client;
    }

    /// <summary>The raw initialize result is re-fetchable via tools/list etc.; expose the process for asserts.</summary>
    public Process ServerProcess => _server;

    /// <summary>Sends a JSON-RPC notification (no response expected).</summary>
    public async Task NotifyAsync(string method)
    {
        await SendAsync(new { jsonrpc = "2.0", method });
    }

    /// <summary>Sends a request and returns the matching response envelope (throws on JSON-RPC error).</summary>
    public async Task<JsonElement> RoundTripAsync(string method, object parameters)
    {
        int id = Interlocked.Increment(ref _nextId);
        await SendAsync(new { jsonrpc = "2.0", id, method, @params = parameters });

        while (true)
        {
            string? line = await _server.StandardOutput.ReadLineAsync(_cts.Token);
            if (line is null)
            {
                string stderr = await DrainStderrAsync();
                throw new InvalidOperationException(
                    $"Server closed stdout before answering '{method}' (id {id}).\n--- server stderr ---\n{stderr}");
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

            if (root.TryGetProperty("error", out JsonElement error))
            {
                throw new InvalidOperationException($"JSON-RPC error response for '{method}' (id {id}): {error.GetRawText()}");
            }

            return root;
        }
    }

    /// <summary>Lists tool names via tools/list.</summary>
    public async Task<IReadOnlyList<string>> ListToolNamesAsync()
    {
        JsonElement list = await RoundTripAsync("tools/list", new { });
        return list.GetProperty("result").GetProperty("tools").EnumerateArray()
            .Select(t => t.GetProperty("name").GetString()!)
            .ToList();
    }

    /// <summary>
    /// Calls a tool and returns the first text content parsed as JSON (all OutlookAI
    /// tools return a single JSON text block). Asserts isError is not set.
    /// </summary>
    /// <remarks>
    /// A tool error is NOT a transport failure and is not thrown here. Domain failures now
    /// set MCP <c>isError</c> - they used to be protocol-level successes whose text merely
    /// contained an error object - and the error payload is precisely what most of these
    /// tests assert on. Use <see cref="CallToolWithIsErrorAsync"/> when the flag itself
    /// matters; a genuine protocol fault still throws from <see cref="RoundTripAsync"/>.
    /// </remarks>
    public async Task<JsonElement> CallToolAsync(string name, object arguments)
    {
        (JsonElement payload, _) = await CallToolWithIsErrorAsync(name, arguments);
        return payload;
    }

    /// <summary>Calls a tool and returns both its payload and whether MCP flagged it as an error.</summary>
    public async Task<(JsonElement Payload, bool IsError)> CallToolWithIsErrorAsync(string name, object arguments)
    {
        JsonElement call = await RoundTripAsync("tools/call", new { name, arguments });
        JsonElement result = call.GetProperty("result");
        bool isError = result.TryGetProperty("isError", out JsonElement flag) && flag.ValueKind == JsonValueKind.True;

        string? text = result.GetProperty("content").EnumerateArray()
            .First(c => c.GetProperty("type").GetString() == "text")
            .GetProperty("text").GetString();
        if (text == null)
        {
            throw new InvalidOperationException($"tools/call {name} returned no text content.");
        }

        using JsonDocument doc = JsonDocument.Parse(text);
        return (doc.RootElement.Clone(), isError);
    }

    /// <summary>Calls a tool whose result is plain text (echo), returning the raw text.</summary>
    public async Task<string> CallToolRawTextAsync(string name, object arguments)
    {
        JsonElement call = await RoundTripAsync("tools/call", new { name, arguments });
        return call.GetProperty("result").GetProperty("content").EnumerateArray()
            .First(c => c.GetProperty("type").GetString() == "text")
            .GetProperty("text").GetString()!;
    }

    /// <summary>Closes stdin and waits for the server to exit (leak guard, Phase-0 fact).</summary>
    public async Task<bool> CloseAndAwaitExitAsync(TimeSpan timeout)
    {
        _server.StandardInput.Close();
        using var exitCts = CancellationTokenSource.CreateLinkedTokenSource(_cts.Token);
        exitCts.CancelAfter(timeout);
        try
        {
            await _server.WaitForExitAsync(exitCts.Token);
            return true;
        }
        catch (OperationCanceledException)
        {
            return false;
        }
    }

    /// <summary>Kills the server (if still alive) and returns captured stderr.</summary>
    public async Task<string> DrainStderrAsync()
    {
        try
        {
            if (!_server.HasExited)
            {
                _server.Kill(entireProcessTree: true);
            }

            return await _stderrTask.WaitAsync(TimeSpan.FromSeconds(10));
        }
        catch (Exception drainEx)
        {
            return $"<stderr unavailable: {drainEx.GetType().Name}>";
        }
    }

    private async Task SendAsync(object message)
    {
        string json = JsonSerializer.Serialize(message);
        await _server.StandardInput.WriteLineAsync(json.AsMemory(), _cts.Token);
        await _server.StandardInput.FlushAsync(_cts.Token);
    }

    public async ValueTask DisposeAsync()
    {
        try
        {
            if (!_server.HasExited)
            {
                _server.Kill(entireProcessTree: true);
                await _server.WaitForExitAsync(new CancellationTokenSource(TimeSpan.FromSeconds(10)).Token);
            }
        }
        catch (Exception)
        {
            // Teardown must not mask test failures.
        }
        finally
        {
            _server.Dispose();
            _cts.Dispose();
        }
    }
}
