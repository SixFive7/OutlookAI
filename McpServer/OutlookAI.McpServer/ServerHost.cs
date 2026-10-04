using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

namespace OutlookAI.McpServer;

/// <summary>
/// The server's generic host, built so that nothing in it watches a file.
/// <para>
/// <c>Host.CreateApplicationBuilder(args)</c> makes the CURRENT DIRECTORY the content root and
/// reads <c>appsettings.json</c> and <c>appsettings.{environment}.json</c> from it with
/// reload-on-change on - and watching those two files costs one recursive
/// <see cref="FileSystemWatcher"/> over the whole content root. Every Claude Code session starts
/// this server with the current directory set to the session's project folder, so every file
/// created, changed or deleted anywhere in a project cost each server a <c>FileInfo.Exists</c>,
/// and every renamed directory a walk of the whole renamed tree. Measured on the maintainer's
/// machine on 2026-10-04, against 3.1.0.325: up to 0.9 of a core per server, about 95% of it
/// kernel time, 3,000 to 4,900 CPU-seconds per copy, and 2.39 cores across all of them in two
/// minutes while he saw lag. No configuration file ships with the server and it reads none, so
/// the watcher bought nothing.
/// </para>
/// <para>
/// Two settings, either of which ends that, and both because each covers a case the other does
/// not:
/// </para>
/// <list type="bullet">
/// <item><description>
/// <c>ContentRootPath</c> is the server's own folder, <see cref="AppContext.BaseDirectory"/>,
/// whatever directory it is started in. A builder setting outranks every other source -
/// <c>DOTNET_CONTENTROOT</c> and <c>--contentRoot</c> included - so nothing in the environment can
/// point the content root back at a project folder, and the server no longer reads a project's
/// own <c>appsettings.json</c> as if it were its configuration.
/// </description></item>
/// <item><description>
/// <c>hostBuilder:reloadConfigOnChange</c> is false, so no configuration file is watched and
/// no watcher is created at all. It has to be in the configuration the builder starts with: the
/// builder adds the JSON files, and starts their watcher, inside its own constructor. It is the
/// lowest-priority source, so <c>DOTNET_hostBuilder__reloadConfigOnChange=true</c> still turns
/// reloading back on - and the watcher it brings then watches the install folder, which nothing
/// writes to while the server runs.
/// </description></item>
/// </list>
/// <para>
/// <c>T1.ServerHostWatchesNothingTests</c> builds the host through this method and pins both;
/// <c>T3.ServerWorkingDirectoryCiTests</c> starts the shipped exe in a scratch folder and holds
/// its CPU time to a burst of file activity there.
/// </para>
/// </summary>
internal static class ServerHost
{
    /// <summary>The key the generic host reads to decide whether its configuration files reload on change.</summary>
    internal const string ReloadConfigOnChangeKey = "hostBuilder:reloadConfigOnChange";

    /// <summary>Everything <c>Program</c> runs, short of running it.</summary>
    internal static HostApplicationBuilder CreateBuilder(string[] args)
    {
        var startingConfiguration = new ConfigurationManager();
        startingConfiguration.AddInMemoryCollection(new Dictionary<string, string?>
        {
            [ReloadConfigOnChangeKey] = "false",
        });

        HostApplicationBuilder builder = Host.CreateApplicationBuilder(new HostApplicationBuilderSettings
        {
            Args = args,
            Configuration = startingConfiguration,
            ContentRootPath = AppContext.BaseDirectory,
        });

        // stdout carries the MCP JSON-RPC stream; ALL logging must go to stderr.
        builder.Logging.AddConsole(options => options.LogToStandardErrorThreshold = LogLevel.Trace);

        builder.Services
            .AddMcpServer(options => options.ServerInstructions = ServerMetadata.Instructions)
            .WithStdioServerTransport()
            .WithToolsFromAssembly();

        return builder;
    }
}
