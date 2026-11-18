using System.Diagnostics;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Makes a directory junction for a test - a link Windows follows like a folder, and one any account
/// may make, unlike a symbolic link. Through <c>cmd /c mklink /J</c>: one line, where the
/// reparse-point API would take forty. Used by <see cref="AttachmentFileWriterDiskTests"/> and
/// <c>T2.LiveSaveAttachmentFolderTests</c>, which point junctions out of the download folder.
/// </summary>
internal static class TestJunction
{
    /// <summary>Makes <paramref name="link"/>, a junction to the folder <paramref name="target"/>, or fails the test.</summary>
    internal static void Make(string link, string target)
    {
        ProcessStartInfo start = new ProcessStartInfo("cmd.exe", "/d /c mklink /J \"" + link + "\" \"" + target + "\"")
        {
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
        };
        using Process process = Process.Start(start)!;
        string said = process.StandardOutput.ReadToEnd() + process.StandardError.ReadToEnd();
        process.WaitForExit();
        Assert.True(process.ExitCode == 0 && Directory.Exists(link), "mklink /J failed (exit " + process.ExitCode + "): " + said);
    }

    /// <summary>
    /// Removes the junction <paramref name="link"/> itself - never what it leads to - when it is there.
    /// Best-effort, for clean-up.
    /// </summary>
    internal static void Remove(string link)
    {
        try
        {
            if (Directory.Exists(link))
            {
                Directory.Delete(link, recursive: false);
            }
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }
}
