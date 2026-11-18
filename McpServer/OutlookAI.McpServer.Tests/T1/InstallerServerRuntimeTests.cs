using System.Text.RegularExpressions;

using OutlookAI.Core.AddIn;
using OutlookAI.RuntimeSetup;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Setup's half of the mail server's private .NET runtime (Q172), read out of <c>Installer.iss</c>, whose
/// Pascal runs nowhere in this suite: it runs the helper on every install and every silent update,
/// without asking for administrator rights, reads the helper's exit codes as the helper writes them,
/// and its uninstaller deletes exactly what the helper made. The helper itself is
/// <see cref="PrivateRuntimeTests"/>; the compiled installer's run on a test machine is the phase-4
/// proof recorded in Docs/live-tier-on-the-vm.md.
/// </summary>
public sealed class InstallerServerRuntimeTests
{
    [Fact]
    public void SetupRunsTheHelperOnEveryInstallAndEverySilentUpdate()
    {
        string iss = InstallerScript.Text();
        string steps = InstallerScript.Routine(iss, "procedure", "CurStepChanged");

        // After the interactive-only block closes, so a silent update runs it too: the call sits at
        // the ssPostInstall block's own indentation, not inside "if not WizardSilent then begin".
        Assert.Matches(@"(?m)^    SetUpServerRuntime;\r?$", steps);
        int interactive = steps.IndexOf("if not WizardSilent then", StringComparison.Ordinal);
        int interactiveEnd = steps.IndexOf("\n    end;", interactive, StringComparison.Ordinal);
        Assert.True(interactive > 0 && interactiveEnd > interactive, "CurStepChanged's interactive block was not found");
        Assert.True(steps.IndexOf("SetUpServerRuntime;", StringComparison.Ordinal) > interactiveEnd,
            "SetUpServerRuntime runs inside the interactive-only block, so a silent update would never set the runtime up");
    }

    [Fact]
    public void TheHelperRunsAsTheUser_FromTheFolderTheReleasePublishesItTo()
    {
        string routine = InstallerScript.Routine(InstallerScript.Text(), "procedure", "SetUpServerRuntime");

        Assert.Contains("ExpandConstant('{app}\\RuntimeSetup\\OutlookAI.RuntimeSetup.exe')", routine, StringComparison.Ordinal);
        Assert.Contains("Exec(Helper, '--from \"' + AddBackslash(ExpandConstant('{src}')) + '.\"', '', SW_HIDE, ewWaitUntilTerminated, ResultCode)", routine, StringComparison.Ordinal);
        // Exec runs it with setup's own rights; nothing here asks Windows for more.
        Assert.DoesNotContain("runas", routine, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("ShellExec", routine, StringComparison.Ordinal);
    }

    [Fact]
    public void SetupReadsTheHelpersExitCodesAsTheHelperWritesThem()
    {
        string routine = InstallerScript.Routine(InstallerScript.Text(), "procedure", "SetUpServerRuntime");

        Assert.Contains("if ResultCode = 0 then", routine, StringComparison.Ordinal);
        Assert.Equal(0, (int)PrivateRuntime.ExitCode.Ready);
        foreach (PrivateRuntime.ExitCode code in new[] { PrivateRuntime.ExitCode.DownloadFailed, PrivateRuntime.ExitCode.PackageRejected, PrivateRuntime.ExitCode.UnpackFailed, PrivateRuntime.ExitCode.SwapFailed })
            Assert.Matches(@"(?m)^    " + (int)code + @": Reason := '", routine);
    }

    [Fact]
    public void TheUninstallerDeletesTheRuntimeItsWorkFoldersAndItsLog_AndNothingElseOfOutlookAIs()
    {
        string root = @"{localappdata}\OutlookAI\" + ServerRuntime.FolderName;
        string[] deletes = InstallerScript.Section(InstallerScript.Text(), "UninstallDelete");

        Assert.Equal(
            new[]
            {
                "Type: filesandordirs; Name: \"" + root + "\"",
                "Type: filesandordirs; Name: \"" + @"{localappdata}\OutlookAI\" + PrivateRuntime.WorkFolderPrefix(root) + "*\"",
                "Type: files; Name: \"" + PrivateRuntime.LogPath(root) + "\"",
            },
            deletes);
    }

    [Fact]
    public void NothingOfTheMachineWideNet10RuntimeIsLeft()
    {
        string iss = InstallerScript.Text();

        Assert.DoesNotMatch(new Regex("NetRuntime10|IsNetRuntime10Installed|DownloadAndInstallNetRuntime|dotnet-runtime-10"), iss);
        Assert.Equal("https://dotnet.microsoft.com/download/dotnet/" + ServerRuntime.Major + ".0", InstallerScript.Define(iss, "ServerRuntimeManualUrl"));
    }
}
