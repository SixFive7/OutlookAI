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
    /// <summary>
    /// The token a test passes to <see cref="StartAndInitializeAsync"/> before this client
    /// will send a <c>tools/call</c> for a tool that reaches the machine's own Outlook.
    /// <para>
    /// It exists because "CI-safe" was a claim in a class name rather than a fact: sixteen
    /// T3 files ran under <c>Category!=Live</c> and eleven of their tests attached to
    /// whatever Outlook was on the machine, which on the maintainer's box is a production
    /// mailbox. A comment cannot stop the twelfth from being written; a refusal can.
    /// </para>
    /// <para>
    /// A deliberate literal rather than a bool because it has to be VISIBLE to the pin that
    /// enforces the classification: <c>T1.LiveTierInventoryTests</c> reads the IL of every
    /// T3 class looking for exactly this string, and a unique string cannot be confused with
    /// a tool name that some other test merely asserts the schema of. A bool argument
    /// compiles to <c>ldc.i4.1</c>, which is indistinguishable from every other true.
    /// </para>
    /// </summary>
    public const string OutlookReachingToolsAllowed = "outlook-reaching-tools-allowed";

    /// <summary>
    /// The tools that reach Outlook - and the Windows Search index - for EVERY argument
    /// shape, so no test can call one and stay off the mailbox by validating badly.
    /// <para>
    /// The rest of the surface is judged by its arguments instead, and is therefore NOT
    /// listed here: <c>search</c>, <c>read</c>, <c>thread</c>, the draft tools,
    /// <c>move_mail</c> and the show-me tools all have a refusal that fires before any COM
    /// work, which is what the CI-safe half of this tier is built on. This guard reads the
    /// tool NAME only; it cannot tell a bounded exhaustive <c>search</c> (which does reach
    /// Outlook) from an unbounded one (which is refused), so that judgement stays with the
    /// class-level classification and with review.
    /// </para>
    /// </summary>
    private static readonly string[] ToolsThatAlwaysReachOutlook =
    {
        "outlook_health",
        "list_accounts",
        "list_folders",
    };

    private readonly Process _server;
    private readonly CancellationTokenSource _cts;
    private readonly Task<string> _stderrTask;
    private readonly bool _outlookReachingToolsAllowed;
    private int _nextId;
    private JsonElement _initializeResult;

    private McpStdioClient(
        Process server,
        CancellationTokenSource cts,
        Task<string> stderrTask,
        bool outlookReachingToolsAllowed)
    {
        _server = server;
        _cts = cts;
        _stderrTask = stderrTask;
        _outlookReachingToolsAllowed = outlookReachingToolsAllowed;
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
    /// <param name="outlookReachingTools">
    /// Pass <see cref="OutlookReachingToolsAllowed"/> to permit the tools listed in
    /// <c>ToolsThatAlwaysReachOutlook</c>. Anything else - null included - makes this client
    /// refuse them, so a test that has not thought about mailbox contact cannot make it by
    /// accident.
    /// </param>
    /// <summary>
    /// The MCP revision this client speaks, and the one every T3 assertion is written
    /// against. Sent in <c>initialize</c> and compared against what the server answers.
    /// </summary>
    public const string ProtocolVersion = "2025-06-18";

    public static async Task<McpStdioClient> StartAndInitializeAsync(
        TimeSpan? timeout = null,
        IReadOnlyDictionary<string, string>? environment = null,
        string? outlookReachingTools = null)
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
        var client = new McpStdioClient(
            server,
            cts,
            stderrTask,
            string.Equals(outlookReachingTools, OutlookReachingToolsAllowed, StringComparison.Ordinal));

        JsonElement init = await client.RoundTripAsync("initialize", new
        {
            protocolVersion = ProtocolVersion,
            capabilities = new { },
            clientInfo = new { name = "OutlookAI.T3Client", version = "0.0.2" },
        });

        // COMPARED, not merely fetched. This used to be `_ = init...GetProperty(...)`, which
        // asserted the field EXISTS and nothing more: a server that negotiated a different
        // protocol version passed every T3 test silently, while a real client would have
        // disconnected. The whole T3 tier is a conformance suite, so the one field that says
        // which conformance is being claimed has to be checked.
        string negotiated = init.GetProperty("result").GetProperty("protocolVersion").GetString()
            ?? throw new InvalidOperationException("The server's initialize result carried no protocolVersion string.");
        if (!string.Equals(negotiated, ProtocolVersion, StringComparison.Ordinal))
        {
            throw new InvalidOperationException(
                $"MCP protocol version mismatch: this client speaks '{ProtocolVersion}' and the server negotiated "
                + $"'{negotiated}'. Every T3 assertion below is written against '{ProtocolVersion}'. Either the SDK was "
                + "upgraded to a newer revision (update this client and re-read the T3 expectations against it), or the "
                + "server is answering with a version it does not implement.");
        }
        client._initializeResult = init.GetProperty("result").Clone();
        await client.NotifyAsync("notifications/initialized");
        return client;
    }

    /// <summary>
    /// The <c>result</c> object of the initialize handshake, kept because the handshake
    /// happens once and cannot be replayed on a live session - <c>instructions</c> and the
    /// advertised capabilities are only ever on the wire here.
    /// </summary>
    public JsonElement InitializeResult => _initializeResult;

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
        // Guarded HERE rather than in CallToolAsync because this is the only choke point:
        // the supervision and availability tests build their own tools/call envelopes and
        // send them straight through, so a check on the convenience helpers would leave the
        // two files that reach Outlook hardest unguarded.
        if (string.Equals(method, "tools/call", StringComparison.Ordinal))
        {
            RefuseUndeclaredOutlookContact(parameters);
        }

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

    /// <summary>
    /// Spends this server process's one writing-rules rejection, so a test whose subject is
    /// something else can assert on its own subject.
    /// <para>
    /// The server refuses the FIRST drafting call of every process on purpose and answers with
    /// the user's writing rules attached (<c>WritingRulesGate</c>); the retry is what gets
    /// through. Each of these tests spawns a fresh server, so each would otherwise meet that
    /// rejection instead of the validation error, refusal or draft it came to check.
    /// </para>
    /// <para>
    /// The priming call supplies BOTH body and body_html, which the tool layer rejects as
    /// mutually exclusive before any COM work. So it costs one round trip and can never reach
    /// Outlook - not even if the gate does not fire, which is a legitimate state (a user who
    /// cleared their rules has none to deliver). Nothing is asserted here for that reason.
    /// </para>
    /// </summary>
    public async Task PrimeWritingRulesGateAsync()
    {
        await CallToolWithIsErrorAsync("new_draft", new
        {
            account = "nobody@example.invalid",
            to = "nobody@example.invalid",
            subject = "writing-rules gate priming",
            body = "priming",
            body_html = "<p>priming</p>",
            display = false,
        });
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

    /// <summary>
    /// Fails the test rather than the mailbox: a <c>tools/call</c> naming a tool that always
    /// reaches Outlook is refused unless the test declared it.
    /// </summary>
    /// <remarks>
    /// The parameters are serialised to read the tool name because the callers pass
    /// anonymous types, and reflecting over those would break the moment one of them named
    /// the property differently. Serialising is what goes on the wire anyway, so this reads
    /// exactly what the server would have been asked for.
    /// </remarks>
    private void RefuseUndeclaredOutlookContact(object parameters)
    {
        if (_outlookReachingToolsAllowed)
        {
            return;
        }

        JsonElement envelope = JsonSerializer.SerializeToElement(parameters);
        if (envelope.ValueKind != JsonValueKind.Object
            || !envelope.TryGetProperty("name", out JsonElement nameProperty))
        {
            return;
        }

        string? tool = nameProperty.GetString();
        if (tool == null || Array.IndexOf(ToolsThatAlwaysReachOutlook, tool) < 0)
        {
            return;
        }

        throw new InvalidOperationException(
            $"'{tool}' reaches the machine's own Outlook and the Windows Search index for every argument shape, "
            + "so it may not be called from a test that has not declared mailbox contact. Either call a tool whose "
            + "arguments are refused before any COM work, or move this test into a class carrying "
            + "Category=Live / LiveTier / Requires=OutlookInstance and start the client with "
            + $"outlookReachingTools: {nameof(McpStdioClient)}.{nameof(OutlookReachingToolsAllowed)}. "
            + "T1 LiveTierInventoryTests pins that pairing.");
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
