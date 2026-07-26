using System;
using System.Diagnostics;
using System.Reflection;
using System.Text;
using System.Text.Json;
using Xunit;

namespace OutlookAI.McpServer.Tests.T3;

/// <summary>
/// T3 conformance tier (v3.MD section 0.6): spawns the built server exe and speaks raw
/// MCP JSON-RPC over stdio - initialize, notifications/initialized, tools/list, tools/call.
/// The client is deliberately hand-rolled (not the SDK client) so the wire protocol itself
/// is proven, independent of the ModelContextProtocol package on both ends. This test needs
/// no Outlook and no index, so it runs in CI as well.
/// </summary>
public sealed class McpStdioConformanceTests
{
    private static string ServerExePath =>
        typeof(McpStdioConformanceTests).Assembly
            .GetCustomAttributes<AssemblyMetadataAttribute>()
            .FirstOrDefault(a => a.Key == "McpServerExePath")?.Value
        ?? throw new InvalidOperationException("AssemblyMetadata 'McpServerExePath' is missing.");

    [Fact]
    public async Task Handshake_ListsTools_And_CallsOutlookHealth()
    {
        string exePath = ServerExePath;
        Assert.True(File.Exists(exePath), $"Server exe not found at '{exePath}' - build OutlookAI.McpServer first.");

        using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(120));
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

        using Process server = Process.Start(psi)
            ?? throw new InvalidOperationException("Failed to start the server process.");
        Task<string> stderrTask = server.StandardError.ReadToEndAsync(cts.Token);

        try
        {
            // 1. initialize - the server must answer with protocol version + server info.
            JsonElement init = await RoundTripAsync(server, id: 1, method: "initialize", parameters: new
            {
                protocolVersion = "2025-06-18",
                capabilities = new { },
                clientInfo = new { name = "OutlookAI.T3Client", version = "0.0.1" },
            }, cts.Token);
            JsonElement initResult = init.GetProperty("result");
            Assert.False(string.IsNullOrWhiteSpace(initResult.GetProperty("protocolVersion").GetString()));
            Assert.False(string.IsNullOrWhiteSpace(initResult.GetProperty("serverInfo").GetProperty("name").GetString()));
            Assert.True(initResult.GetProperty("capabilities").TryGetProperty("tools", out _),
                "Server must advertise the tools capability.");
            // D36 (soak fix 6): the initialize result must carry the server instructions
            // EXACTLY as authored - Claude Code injects the field passively into every
            // session at start, even when the tool roster is deferred by tool search.
            Assert.Equal(ServerMetadata.Instructions, initResult.GetProperty("instructions").GetString());

            // 2. initialized notification (no response expected).
            await SendAsync(server, new { jsonrpc = "2.0", method = "notifications/initialized" }, cts.Token);

            // 3. tools/list - the full L1-L5 surface plus the merged diagnostics tool
            // (echo/index_status/health were deleted in soak fix D37).
            JsonElement list = await RoundTripAsync(server, id: 2, method: "tools/list", parameters: new { }, cts.Token);
            var names = list.GetProperty("result").GetProperty("tools").EnumerateArray()
                .Select(t => t.GetProperty("name").GetString())
                .ToList();
            foreach (string expected in new[]
                     {
                         "search", "thread", "read", "save_attachment", "list_accounts", "list_folders", "list_signatures",
                         "manage_signature", "move_mail", "archive_mail", "open_in_outlook", "goto_folder", "show_search_results",
                         "new_draft", "reply_draft", "replyall_draft", "forward_draft",
                         "send", "outlook_health",
                     })
            {
                Assert.Contains(expected, names);
            }

            foreach (string deleted in new[] { "echo", "index_status", "health" })
            {
                Assert.DoesNotContain(deleted, names);
            }

            // 4. tools/call outlook_health - the smoke call (echo's old role; callable
            // on any machine because health degrades instead of throwing, and it proves
            // the server -> OutlookAI.Core chain since the report is built in Core).
            JsonElement call = await RoundTripAsync(server, id: 3, method: "tools/call", parameters: new
            {
                name = "outlook_health",
                arguments = new { },
            }, cts.Token);
            JsonElement callResult = call.GetProperty("result");
            if (callResult.TryGetProperty("isError", out JsonElement isError))
            {
                Assert.False(isError.ValueKind == JsonValueKind.True, "outlook_health tool call reported isError=true.");
            }

            string? text = callResult.GetProperty("content").EnumerateArray()
                .First(c => c.GetProperty("type").GetString() == "text")
                .GetProperty("text").GetString();
            Assert.NotNull(text);
            using (JsonDocument healthDoc = JsonDocument.Parse(text!))
            {
                string? status = healthDoc.RootElement.GetProperty("status").GetString();
                Assert.True(status is "ok" or "degraded", $"unexpected outlook_health status '{status}'");
            }

            // 5. Closing stdin must terminate the server - agent sessions must not leak processes.
            server.StandardInput.Close();
            using var exitCts = CancellationTokenSource.CreateLinkedTokenSource(cts.Token);
            exitCts.CancelAfter(TimeSpan.FromSeconds(30));
            try
            {
                await server.WaitForExitAsync(exitCts.Token);
            }
            catch (OperationCanceledException)
            {
                Assert.Fail("Server did not exit within 30 s after stdin was closed.");
            }
        }
        catch (Exception ex) when (ex is not Xunit.Sdk.XunitException)
        {
            string stderr = await DrainStderrAsync(server, stderrTask);
            throw new InvalidOperationException(
                $"T3 conformance run failed: {ex.Message}{Environment.NewLine}--- server stderr ---{Environment.NewLine}{stderr}", ex);
        }
        finally
        {
            if (!server.HasExited)
            {
                server.Kill(entireProcessTree: true);
            }
        }
    }

    private static async Task SendAsync(Process server, object message, CancellationToken ct)
    {
        string json = JsonSerializer.Serialize(message);
        await server.StandardInput.WriteLineAsync(json.AsMemory(), ct);
        await server.StandardInput.FlushAsync(ct);
    }

    /// <summary>
    /// Sends a JSON-RPC request and reads newline-delimited JSON messages from stdout until
    /// the response with the matching id arrives. Server-initiated notifications in between
    /// are legal per MCP and are skipped.
    /// </summary>
    private static async Task<JsonElement> RoundTripAsync(
        Process server, int id, string method, object parameters, CancellationToken ct)
    {
        await SendAsync(server, new { jsonrpc = "2.0", id, method, @params = parameters }, ct);

        while (true)
        {
            string? line = await server.StandardOutput.ReadLineAsync(ct);
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
                continue; // notification or unrelated message
            }

            Assert.Equal("2.0", root.GetProperty("jsonrpc").GetString());
            if (root.TryGetProperty("error", out JsonElement error))
            {
                throw new InvalidOperationException(
                    $"JSON-RPC error response for '{method}' (id {id}): {error.GetRawText()}");
            }

            return root;
        }
    }

    private static async Task<string> DrainStderrAsync(Process server, Task<string> stderrTask)
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
