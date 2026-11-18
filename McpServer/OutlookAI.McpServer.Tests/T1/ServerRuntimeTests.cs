using OutlookAI.Core.AddIn;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Where the mail server's private .NET runtime lives and when it is complete (Q172) - the one copy
/// of the answer, which setup's helper checks its work with and the add-in asks before it registers
/// the server. The machine-wide probe reads this machine's real .NET, so it is not pinned here; the
/// build machine has an SDK of the same major, which it finds (<see cref="TheBuildMachinesOwnNetCountsAsMachineWide"/>).
/// </summary>
public sealed class ServerRuntimeTests : IDisposable
{
    private readonly string _root = Path.Combine(Path.GetTempPath(), "OutlookAI-ServerRuntimeTests-" + Guid.NewGuid().ToString("N"), "dotnet");

    public void Dispose() => ScratchFiles.TryDeleteTree(Path.GetDirectoryName(_root)!);

    private void Place(string relative)
    {
        string path = Path.Combine(_root, relative);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path, "x");
    }

    [Fact]
    public void TheRuntimeIsElevenInOutlookAIsOwnFolder()
    {
        Assert.Equal("11.0.0", ServerRuntime.Version);
        Assert.Equal("11", ServerRuntime.Major);
        Assert.Equal("dotnet", ServerRuntime.FolderName);
        Assert.Equal(
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OutlookAI", "dotnet"),
            ServerRuntime.DefaultRoot());
    }

    [Fact]
    public void AFolderWithTheHostResolverAndTheFrameworkIsComplete()
    {
        Place(Path.Combine("host", "fxr", ServerRuntime.Version, "hostfxr.dll"));
        Place(Path.Combine("shared", "Microsoft.NETCore.App", ServerRuntime.Version, "System.Private.CoreLib.dll"));

        Assert.True(ServerRuntime.IsComplete(_root));
    }

    [Fact]
    public void AFolderWithoutTheFrameworkIsNot()
    {
        Place(Path.Combine("host", "fxr", ServerRuntime.Version, "hostfxr.dll"));

        Assert.False(ServerRuntime.IsComplete(_root));
    }

    [Fact]
    public void AFolderWithoutTheHostResolverIsNot()
    {
        Place(Path.Combine("shared", "Microsoft.NETCore.App", ServerRuntime.Version, "System.Private.CoreLib.dll"));

        Assert.False(ServerRuntime.IsComplete(_root));
    }

    [Fact]
    public void AnotherVersionIsNot()
    {
        Place(Path.Combine("host", "fxr", "10.0.15", "hostfxr.dll"));
        Place(Path.Combine("shared", "Microsoft.NETCore.App", "10.0.15", "System.Private.CoreLib.dll"));

        Assert.False(ServerRuntime.IsComplete(_root));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("C:\\no\\such\\folder")]
    [InlineData("\0")]
    public void NoFolderIsNot(string? root)
    {
        Assert.False(ServerRuntime.IsComplete(root));
    }

    [Fact]
    public void TheBuildMachinesOwnNetCountsAsMachineWide()
    {
        // The suite runs on the .NET of the major it targets, and that one is machine-wide wherever
        // the suite runs: the build VM's SDK, a developer's. So the probe must find it.
        Assert.Equal(ServerRuntime.Major, Environment.Version.Major.ToString(System.Globalization.CultureInfo.InvariantCulture));
        Assert.True(ServerRuntime.IsMachineWideInstalled());
        Assert.True(ServerRuntime.IsAvailable());
    }
}
