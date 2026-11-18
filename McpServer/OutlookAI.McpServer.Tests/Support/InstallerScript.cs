using System.Text.RegularExpressions;

using Xunit;

namespace OutlookAI.McpServer.Tests;

/// <summary>
/// <c>Installer.iss</c> read as text, for the tests that pin what setup does: Inno Setup's Pascal runs
/// nowhere in this suite, so its text is what a test can hold. The whole script, one <c>#define</c>'s
/// value, one <c>[Code]</c> routine, and one section. The one copy for a new test (Q154);
/// <c>InstallerCompromisedCertificateTests</c> and <c>InstallerCompromisedTrustEntryTests</c> still carry
/// private copies of the first three, which the shared-code refactor in TODO.md folds in.
/// </summary>
internal static class InstallerScript
{
    /// <summary>The script as it is on disk now.</summary>
    public static string Text() => File.ReadAllText(Path.Combine(RepoPaths.RepoRoot(), "Installer.iss"));

    /// <summary>The value of <c>#define <paramref name="name"/> "..."</c>.</summary>
    public static string Define(string iss, string name)
    {
        Match m = Regex.Match(iss, "^#define " + Regex.Escape(name) + " \"([^\"]*)\"\\r?$", RegexOptions.Multiline);
        Assert.True(m.Success, "Installer.iss defines no " + name);
        return m.Groups[1].Value;
    }

    /// <summary>A [Code] routine, from its heading to the "end;" that closes it at the start of a line.</summary>
    public static string Routine(string iss, string kind, string name)
    {
        Match m = Regex.Match(iss, "^" + kind + " " + Regex.Escape(name) + @"\b.*?^end;", RegexOptions.Multiline | RegexOptions.Singleline);
        Assert.True(m.Success, "Installer.iss has no " + kind + " " + name);
        return m.Value;
    }

    /// <summary>A section's lines, from its <c>[Name]</c> heading to the next heading, comments left out.</summary>
    public static string[] Section(string iss, string name)
    {
        string[] lines = SourceText.Lines(iss);
        int start = Array.IndexOf(lines, "[" + name + "]");
        Assert.True(start >= 0, "Installer.iss has no [" + name + "] section");
        return lines.Skip(start + 1)
            .TakeWhile(l => !l.StartsWith('['))
            .Where(l => l.Length > 0 && !l.StartsWith(';'))
            .ToArray();
    }
}
