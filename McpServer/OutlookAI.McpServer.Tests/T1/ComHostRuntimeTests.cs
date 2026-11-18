using OutlookAI.ComHost.Supervision;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The COM host runs on the server's own .NET (Q172). The installed server's apphost finds the private
/// runtime setup puts in %LOCALAPPDATA%\OutlookAI\dotnet; the COM host's keeps the default search, which
/// would not find it, so the server starts it with DOTNET_ROOT_X64 and DOTNET_ROOT naming the runtime
/// the server itself runs on - three folders above System.Private.CoreLib's.
/// </summary>
public sealed class ComHostRuntimeTests : IDisposable
{
    private readonly string _root = Path.Combine(Path.GetTempPath(), "OutlookAI-ComHostRuntimeTests-" + Guid.NewGuid().ToString("N"), "dotnet");

    public void Dispose() => ScratchFiles.TryDeleteTree(Path.GetDirectoryName(_root)!);

    private string CoreLibDirectory() => Path.Combine(_root, "shared", "Microsoft.NETCore.App", "11.0.0");

    [Fact]
    public void TheChildIsPointedAtTheRuntimeTheServerRunsOn()
    {
        Directory.CreateDirectory(CoreLibDirectory());
        Directory.CreateDirectory(Path.Combine(_root, "host", "fxr", "11.0.0"));

        IReadOnlyList<KeyValuePair<string, string>> variables = ComHostSupervisor.OwnRuntimeVariables(CoreLibDirectory());

        Assert.Equal(
            new[] { new KeyValuePair<string, string>("DOTNET_ROOT_X64", _root), new KeyValuePair<string, string>("DOTNET_ROOT", _root) },
            variables);
    }

    [Fact]
    public void ARuntimeFolderWithoutAHostResolverNamesNothing()
    {
        Directory.CreateDirectory(CoreLibDirectory());

        Assert.Empty(ComHostSupervisor.OwnRuntimeVariables(CoreLibDirectory()));
    }

    [Fact]
    public void ThisProcesssOwnRuntimeIsOneTheChildCanStartOn()
    {
        IReadOnlyList<KeyValuePair<string, string>> variables = ComHostSupervisor.OwnRuntimeVariables();

        Assert.Equal(2, variables.Count);
        string root = variables[0].Value;
        Assert.True(Directory.Exists(Path.Combine(root, "host", "fxr")), root + " has no host resolver");
        Assert.True(File.Exists(Path.Combine(root, "shared", "Microsoft.NETCore.App", Environment.Version.ToString(), "System.Private.CoreLib.dll")),
            root + " does not hold the runtime this test runs on, " + Environment.Version);
    }
}
