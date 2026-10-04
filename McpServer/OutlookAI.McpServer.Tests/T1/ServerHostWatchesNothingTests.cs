using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.FileProviders;
using Microsoft.Extensions.Hosting;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The server's host watches no file, and its content root is its own folder - never the
/// directory it was started in.
/// <para>
/// Every Claude Code session starts its own server with the working directory set to the
/// session's project folder. <c>Host.CreateApplicationBuilder(args)</c> made that folder the
/// content root and watched it, recursively, for changes to <c>appsettings.json</c>, so every
/// file event anywhere in a project cost every server started there - up to 0.9 of a core per
/// copy on the maintainer's machine on 2026-10-04. <see cref="ServerHost"/> has the whole story;
/// <c>T3.ServerWorkingDirectoryCiTests</c> holds the shipped exe to it from the outside.
/// </para>
/// <para>
/// Each host is built with the working directory moved to a scratch folder, because under the
/// test runner the working directory is usually the test assembly's own folder: a content root
/// taken from it would equal <see cref="AppContext.BaseDirectory"/>, and the defect would pass.
/// Moving it is safe here because collections run one at a time (<c>xunit.runner.json</c>), and
/// it is put back before the build returns.
/// </para>
/// </summary>
public sealed class ServerHostWatchesNothingTests
{
    [Fact]
    public void TheContentRoot_IsTheServersOwnFolder_NotTheWorkingDirectory()
    {
        HostApplicationBuilder builder = BuildIn(out string workingDirectory, _ => ServerHost.CreateBuilder([]));

        Assert.NotEqual(Normalize(AppContext.BaseDirectory), Normalize(workingDirectory), StringComparer.OrdinalIgnoreCase);
        Assert.Equal(Normalize(AppContext.BaseDirectory), Normalize(builder.Environment.ContentRootPath), ignoreCase: true);
        PhysicalFileProvider contentRootFiles = Assert.IsType<PhysicalFileProvider>(builder.Environment.ContentRootFileProvider);
        Assert.Equal(Normalize(AppContext.BaseDirectory), Normalize(contentRootFiles.Root), ignoreCase: true);
    }

    [Fact]
    public void TheContentRoot_CannotBePointedAtTheWorkingDirectory()
    {
        // The same precedence covers DOTNET_CONTENTROOT; the command line is the source a test
        // can set without changing the environment of every test after it.
        HostApplicationBuilder builder = BuildIn(out _, dir => ServerHost.CreateBuilder(["--contentRoot", dir]));

        Assert.Equal(Normalize(AppContext.BaseDirectory), Normalize(builder.Environment.ContentRootPath), ignoreCase: true);
    }

    [Fact]
    public void NoConfigurationFile_IsWatched()
    {
        using IHost host = BuildIn(out _, _ => ServerHost.CreateBuilder([])).Build();
        FileConfigurationProvider[] files = FileProvidersOf(host);

        // The defaults still read appsettings.json and appsettings.Production.json, optionally, from
        // the content root. Present, so the check below is about the real sources; none watched.
        Assert.Contains(files, f => f.Source.Path == "appsettings.json");
        Assert.All(files, f => Assert.False(
            f.Source.ReloadOnChange,
            $"'{f.Source.Path}' reloads on change - which starts a recursive FileSystemWatcher over the whole content root"));
    }

    [Fact]
    public void TheStockBuilder_FailsBothChecks_SoTheyCanCatchARevert()
    {
        // Host.CreateApplicationBuilder(args) is what Program called before 2026-10-04. If this ever
        // passes the checks above, they no longer tell the fixed host from the broken one.
        using IHost stock = BuildIn(out string workingDirectory, _ => Host.CreateApplicationBuilder(Array.Empty<string>())).Build();

        Assert.Equal(Normalize(workingDirectory), Normalize(stock.Services.GetRequiredService<IHostEnvironment>().ContentRootPath), ignoreCase: true);
        Assert.Contains(FileProvidersOf(stock), f => f.Source.ReloadOnChange);
    }

    private static T BuildIn<T>(out string workingDirectory, Func<string, T> build)
    {
        string scratch = Path.Combine(Path.GetTempPath(), "OutlookAI-ServerHostTests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(scratch);
        string original = Environment.CurrentDirectory;
        try
        {
            Environment.CurrentDirectory = scratch;
            workingDirectory = scratch;
            return build(scratch);
        }
        finally
        {
            Environment.CurrentDirectory = original;

            // Nothing is ever written there. A watcher the stock builder put on it is gone with its
            // host; until then Windows only marks the folder for deletion.
            Directory.Delete(scratch, recursive: true);
        }
    }

    private static FileConfigurationProvider[] FileProvidersOf(IHost host) =>
        ((IConfigurationRoot)host.Services.GetRequiredService<IConfiguration>()).Providers
            .OfType<FileConfigurationProvider>()
            .ToArray();

    private static string Normalize(string path) => Path.TrimEndingDirectorySeparator(Path.GetFullPath(path));
}
