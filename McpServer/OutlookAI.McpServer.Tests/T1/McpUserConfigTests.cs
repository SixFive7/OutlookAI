using OutlookAI.Core.AddIn;
using Xunit;

using EntryState = OutlookAI.Core.AddIn.McpRegistrationDecision.EntryState;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The user-scope registration in Claude Code's <c>~/.claude.json</c>: what the add-in's reconcile
/// judges the <c>outlookai</c> entry to be, and how it rewrites it. Until stage S3 of the code-sharing
/// refactor this was the add-in's own <c>McpRegistrationService</c> over a
/// <c>JavaScriptSerializer</c> dictionary, which no test could load. It is OutlookAI.Core's
/// <see cref="McpConfigEditor"/> now (Q163 (a)), reading with the scanner <c>outlook_health</c> uses,
/// and every judgement below is pinned here.
/// </summary>
public sealed class McpUserConfigTests
{
    private const string Installed = @"C:\Users\x\AppData\Local\OutlookAI\Setup\McpServer\OutlookAI.McpServer.exe";
    private const string Portable = "${LOCALAPPDATA}/OutlookAI/Setup/McpServer/OutlookAI.McpServer.exe";
    private const string Exe = "OutlookAI.McpServer.exe";

    private static string? Env(string name) => name == "LOCALAPPDATA" ? @"C:\Users\x\AppData\Local" : null;

    private static EntryState Classify(string raw, string installed = Installed) =>
        McpConfigEditor.ClassifyServerEntry(McpConfigEditor.IsReadableConfig(true, raw), raw, installed, Exe, Env);

    private static string Entry(string body) => "{\"mcpServers\":{\"outlookai\":" + body + "}}";

    // ===== Which files may be read at all =====

    [Theory]
    [InlineData(false, "", true)]          // absent: readable, holds nothing
    [InlineData(true, "", false)]          // there, but read back empty: a failed read
    [InlineData(true, "  \n", false)]
    [InlineData(true, "{\"a\":1", false)]  // cut off mid-write
    [InlineData(true, "[1]", false)]       // not an object
    [InlineData(true, "{}", true)]
    public void OnlyAnAbsentFileOrOneWellFormedObjectIsReadable(bool exists, string raw, bool expected)
    {
        Assert.Equal(expected, McpConfigEditor.IsReadableConfig(exists, raw));
    }

    // ===== What the entry is =====

    [Fact]
    public void AnUnreadableFile_IsUnreadable()
    {
        Assert.Equal(EntryState.Unreadable, McpConfigEditor.ClassifyServerEntry(false, "{", Installed, Exe, Env));
    }

    [Theory]
    [InlineData("{}")]
    [InlineData("{\"mcpServers\":{}}")]
    [InlineData("{\"mcpServers\":{\"other\":{\"command\":\"x.exe\"}}}")]
    [InlineData("{\"mcpServers\":\"not-an-object\"}")]
    [InlineData("{\"projects\":{\"C:/p\":{\"mcpServers\":{\"outlookai\":{\"command\":\"C:/x.exe\"}}}}}")]
    public void NoTopLevelEntry_IsAbsent(string raw)
    {
        Assert.Equal(EntryState.Absent, Classify(raw));
    }

    [Fact]
    public void AnAbsentFile_IsAbsent()
    {
        Assert.Equal(EntryState.Absent, McpConfigEditor.ClassifyServerEntry(true, "", Installed, Exe, Env));
    }

    [Theory]
    [InlineData("\"C:/x.exe\"")]                                                              // not an object
    [InlineData("{\"url\":\"https://example.com/mcp\"}")]                                     // remote
    [InlineData("{\"httpUrl\":\"https://example.com/mcp\",\"command\":\"OutlookAI.McpServer.exe\"}")]
    [InlineData("{\"type\":\"sse\",\"command\":\"OutlookAI.McpServer.exe\"}")]               // not stdio
    [InlineData("{\"type\":5,\"command\":\"OutlookAI.McpServer.exe\"}")]                      // a type that is no string
    [InlineData("{\"type\":\"stdio\"}")]                                                      // no command
    [InlineData("{\"command\":\"\"}")]
    [InlineData("{\"command\":42}")]
    [InlineData("{\"command\":\"${UNSET}\"}")]                                                // expands to nothing
    [InlineData("{\"command\":\"C:\\\\my\\\\wrapper.cmd\"}")]                                  // the user's own choice
    public void AnythingButOurInstalledServer_IsForeign(string body)
    {
        Assert.Equal(EntryState.Foreign, Classify(Entry(body)));
    }

    [Theory]
    [InlineData("{\"command\":\"C:\\\\Users\\\\x\\\\AppData\\\\Local\\\\OutlookAI\\\\Setup\\\\McpServer\\\\OutlookAI.McpServer.exe\"}")]
    [InlineData("{\"type\":\"STDIO\",\"command\":\"c:/users/x/appdata/local/outlookai/setup/mcpserver/outlookai.mcpserver.exe\"}")]
    [InlineData("{\"type\":\"stdio\",\"command\":\"" + Portable + "\",\"args\":[],\"env\":{}}")]
    public void TheInstalledServer_IsOurs(string body)
    {
        Assert.Equal(EntryState.Ours, Classify(Entry(body)));
    }

    [Theory]
    [InlineData("{\"command\":\"D:/builds/OutlookAI.McpServer.exe\"}", EntryState.Ours)]
    [InlineData("{\"command\":\"D:/builds/Other.exe\"}", EntryState.Foreign)]
    public void WithNoInstalledServer_OurExecutablesNameIsWhatCounts(string body, EntryState expected)
    {
        Assert.Equal(expected, Classify(Entry(body), installed: ""));
    }

    [Theory]
    [InlineData("{\"command\":\"x.exe\"}", true, false)]
    [InlineData("{\"type\":\"stdio\",\"command\":\"x.exe\"}", true, false)]
    [InlineData("{\"type\":\"sse\",\"url\":\"https://example.com\"}", false, true)]
    [InlineData("{\"command\":\"x.exe\",\"url\":\"https://example.com\"}", false, true)]
    [InlineData("{\"type\":\"stdio\",\"command\":\"x.exe\",\"headers\":{}}", true, true)]
    [InlineData("{}", false, false)]
    public void StdioAndRemoteKeys_AreReadOffTheEntry(string body, bool stdio, bool remote)
    {
        Assert.Equal(stdio, McpConfigEditor.IsStdioEntry(Entry(body)));
        Assert.Equal(remote, McpConfigEditor.HasRemoteKeys(Entry(body)));
    }

    // ===== The rewrite =====

    [Fact]
    public void AnAbsentFile_GetsANewDocument()
    {
        Assert.True(McpConfigEditor.TryBuildUserConfig("", false, Portable, false, out string updated, out _));

        Assert.True(McpConfigEditor.TryReadServerCommand(updated, out string command));
        Assert.Equal(Portable, command);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void AFileThatReadsBackEmpty_IsNeverWritten(string raw)
    {
        Assert.False(McpConfigEditor.TryBuildUserConfig(raw, true, Portable, false, out _, out string error));
        Assert.Contains("read back empty", error, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("{\"a\":1")]
    [InlineData("{\"mcpServers\":[]}")]
    public void AFileItCannotSafelySplice_IsRefused(string raw)
    {
        Assert.False(McpConfigEditor.TryBuildUserConfig(raw, true, Portable, false, out _, out _));
    }

    [Fact]
    public void EveryOtherServerAndSetting_KeepsItsBytes()
    {
        string other = "{\n      \"command\": \"node\",\n      \"args\": [ \"x.js\" ]\n    }";
        string raw = "{\n  \"numStartups\": 7,\n  \"mcpServers\": {\n    \"other\": " + other
                     + ",\n    \"outlookai\": { \"command\": \"C:/old/OutlookAI.McpServer.exe\" }\n  },\n  \"projects\": {}\n}\n";

        Assert.True(McpConfigEditor.TryBuildUserConfig(raw, true, Portable, true, out string updated, out _));

        Assert.Contains("\"other\": " + other, updated, StringComparison.Ordinal);
        Assert.StartsWith("{\n  \"numStartups\": 7,\n  \"mcpServers\": {\n", updated, StringComparison.Ordinal);
        Assert.EndsWith("\n  },\n  \"projects\": {}\n}\n", updated, StringComparison.Ordinal);
        Assert.True(McpConfigEditor.TryReadServerCommand(updated, out string command));
        Assert.Equal(Portable, command);
    }

    [Fact]
    public void NoServersYet_TheSectionIsAdded()
    {
        Assert.True(McpConfigEditor.TryBuildUserConfig("{\"numStartups\":7}", true, Portable, false, out string updated, out _));

        Assert.Equal(new[] { "mcpServers", "numStartups" }, McpConfigEditor.ListTopLevelKeys(updated).OrderBy(k => k, StringComparer.Ordinal));
        Assert.True(McpConfigEditor.IsStdioEntry(updated));
    }

    [Fact]
    public void OurOwnEntry_KeepsItsOtherKeys_AndLosesTheRemoteOnes()
    {
        string raw = Entry("{\"command\":\"C:/old/OutlookAI.McpServer.exe\",\"env\":{\"LOG\":\"1\"},\"cwd\":\"C:/w\","
                           + "\"url\":\"https://example.com\",\"headers\":{\"a\":\"b\"}}");

        Assert.True(McpConfigEditor.TryBuildUserConfig(raw, true, Portable, true, out string updated, out _));

        Assert.True(McpConfigEditor.IsStdioEntry(updated));
        Assert.False(McpConfigEditor.HasRemoteKeys(updated));
        Assert.Contains("\"env\": {\"LOG\":\"1\"}", updated, StringComparison.Ordinal);
        Assert.Contains("\"cwd\": \"C:/w\"", updated, StringComparison.Ordinal);
        Assert.Contains("\"args\": []", updated, StringComparison.Ordinal);
        Assert.True(McpConfigEditor.TryReadServerCommand(updated, out string command));
        Assert.Equal(Portable, command);
    }

    [Fact]
    public void SomebodyElsesEntry_IsReplacedWhole()
    {
        string raw = Entry("{\"type\":\"sse\",\"url\":\"https://example.com\",\"env\":{\"LOG\":\"1\"}}");

        Assert.True(McpConfigEditor.TryBuildUserConfig(raw, true, Portable, false, out string updated, out _));

        Assert.DoesNotContain("LOG", updated, StringComparison.Ordinal);
        Assert.DoesNotContain("example.com", updated, StringComparison.Ordinal);
        Assert.Contains(McpConfigEditor.RenderEntryInline(Portable), updated, StringComparison.Ordinal);
    }

    [Fact]
    public void AnEntryThatAlreadyNamesTheCommand_IsStillWritten_SoAHybridIsRepaired()
    {
        // The project scope leaves such an entry alone; the user scope writes what it is asked to,
        // because its caller only asks when a rewrite is due - here, a url beside the right command.
        string raw = Entry("{\"command\":\"" + Portable + "\",\"url\":\"https://example.com\"}");

        Assert.True(McpConfigEditor.TryBuildUserConfig(raw, true, Portable, true, out string updated, out _));

        Assert.NotEqual(raw, updated);
        Assert.False(McpConfigEditor.HasRemoteKeys(updated));
    }

    [Fact]
    public void AKeyWrittenTwice_IsKeptOnce_WithItsFirstValue()
    {
        string raw = Entry("{\"command\":\"C:/old.exe\",\"env\":{\"A\":\"1\"},\"env\":{\"A\":\"2\"}}");

        Assert.True(McpConfigEditor.TryBuildUserConfig(raw, true, Portable, true, out string updated, out _));

        Assert.Contains("\"env\": {\"A\":\"1\"}", updated, StringComparison.Ordinal);
        Assert.DoesNotContain("\"2\"", updated, StringComparison.Ordinal);
    }
}
