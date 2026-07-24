using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using OutlookAI.McpServer;

// OutlookAI.McpServer - stdio MCP host (v3.MD section 0.5, Option A).
// stdout carries the MCP JSON-RPC stream; ALL logging must go to stderr.
var builder = Host.CreateApplicationBuilder(args);
builder.Logging.AddConsole(options => options.LogToStandardErrorThreshold = LogLevel.Trace);

builder.Services
    .AddMcpServer(options => options.ServerInstructions = ServerMetadata.Instructions)
    .WithStdioServerTransport()
    .WithToolsFromAssembly();

await builder.Build().RunAsync().ConfigureAwait(false);
