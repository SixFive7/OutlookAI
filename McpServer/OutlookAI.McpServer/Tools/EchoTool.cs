using System.ComponentModel;
using ModelContextProtocol.Server;
using OutlookAI.Core.Diagnostics;

namespace OutlookAI.McpServer.Tools;

/// <summary>
/// Phase-0 scaffold tool proving the stdio MCP pipeline end-to-end (T3 handshake:
/// initialize + tools/list + tools/call). The real L1-L5 tool surface arrives in
/// Phases 1-5 (v3.MD section 0.5).
/// </summary>
[McpServerToolType]
public static class EchoTool
{
    [McpServerTool(Name = "echo")]
    [Description("Connectivity check: echoes the message back through OutlookAI.Core and reports which Core target framework is loaded.")]
    public static string Echo([Description("Text to echo back.")] string message)
        => $"{Ping.Echo(message)} [Core: {Ping.TargetFramework}]";
}
