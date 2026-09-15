using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using OutlookAI.McpServer;
using OutlookAI.McpServer.Tools;

// OutlookAI.McpServer - stdio MCP host (v3.MD section 0.5, Option A).
// stdout carries the MCP JSON-RPC stream; ALL logging must go to stderr.
var builder = Host.CreateApplicationBuilder(args);
builder.Logging.AddConsole(options => options.LogToStandardErrorThreshold = LogLevel.Trace);

builder.Services
    .AddMcpServer(options => options.ServerInstructions = ServerMetadata.Instructions)
    .WithStdioServerTransport()
    .WithToolsFromAssembly();

// WHY THE try/finally, AND WHAT IT DOES NOT COVER.
//
// The COM host child holds every Outlook reference this product takes out, and its only
// release path runs when it exits of its own accord. Something has to ask it to. Nothing
// did: the gateway is a static Lazy in ServerRuntime rather than a DI singleton, so the
// host's own container never disposed it either, and every shutdown reached the child as a
// termination. ServerRuntime.ReleaseOutlookResources is that ask - see its own comment for
// what it buys and what it does not.
//
// COVERED, because all three end with RunAsync returning and this frame unwinding:
//   - stdin closing, which is how an MCP stdio server normally ends. The SDK's hosted
//     service calls IHostApplicationLifetime.StopApplication on EOF; it does not call
//     Environment.Exit, which was checked against the shipped ModelContextProtocol 2.2.0
//     assembly rather than assumed.
//   - Ctrl-C and SIGTERM, via the default ConsoleLifetime.
//   - an exception escaping the host, which then keeps propagating after this runs.
//
// NOT COVERED, and no hook placed here could cover them:
//   - TerminateProcess. A client that kills its server rather than closing its stdin runs
//     no managed code at all; the child dies with it via the kill-on-close job object,
//     holding whatever it held. Measured on the test guest on 2026-09-15, this costs less
//     than it reads: a client terminated while holding Application, NameSpace and a Folder
//     left the next client's connect at 0.37 s, and so did one that also held an unclosed
//     pin Explorer. What survives is the pin Explorer itself, which accumulates.
//   - a stack overflow, a FailFast, or an unhandled exception on a background thread. The
//     runtime tears the process down without unwinding.
//
// AppDomain.ProcessExit was considered and rejected: it adds only the Environment.Exit case,
// which nothing in this process or in the SDK does, and it would bring a shutdown budget
// shorter than the grace the COM host needs - a hook that reads as covering the hard exits
// while covering none of them.
try
{
    await builder.Build().RunAsync().ConfigureAwait(false);
}
finally
{
    ServerRuntime.ReleaseOutlookResources();
}
