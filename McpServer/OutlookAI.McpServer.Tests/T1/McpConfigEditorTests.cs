using OutlookAI.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The add-in's MCP configuration surgery (<c>Services/McpConfigEditor.cs</c>, LINKED into
/// this test project — see the tests csproj). It edits two files the add-in does not own:
/// Claude Code's <c>~/.claude.json</c> and a project's <c>.mcp.json</c>. A bad splice there
/// silently destroys settings the user cannot get back, so every branch that decides to
/// write — and every branch that decides to refuse — is pinned here.
///
/// Line endings are normalized on both sides of every comparison: the source file's own
/// endings must not decide whether these tests pass.
/// </summary>
public sealed class McpConfigEditorTests
{
    private const string Command = "${LOCALAPPDATA}/OutlookAI/Setup/McpServer/OutlookAI.McpServer.exe";

    private static string Lf(string s) => s.Replace("\r\n", "\n");

    private static void AssertJson(string expected, string actual) => Assert.Equal(Lf(expected), Lf(actual));

    /// <summary>An environment with just the variables a test names.</summary>
    private static Func<string, string> Env(params (string Name, string Value)[] values)
    {
        return name =>
        {
            foreach ((string n, string v) in values)
            {
                if (n == name)
                {
                    return v;
                }
            }

            return "";
        };
    }

    // ===== Environment-variable references =====

    [Fact]
    public void ExpandsASimpleReference()
    {
        Assert.Equal(
            @"C:\Users\x\AppData\Local/OutlookAI/Setup/McpServer/OutlookAI.McpServer.exe",
            McpConfigEditor.ExpandEnvironmentReferences(Command, Env((@"LOCALAPPDATA", @"C:\Users\x\AppData\Local"))));
    }

    [Fact]
    public void UsesTheDefaultWhenTheVariableIsUnset()
    {
        Assert.Equal(
            "fallback/x",
            McpConfigEditor.ExpandEnvironmentReferences("${NOPE:-fallback}/x", Env()));
    }

    [Fact]
    public void PrefersTheVariableOverItsDefault()
    {
        Assert.Equal(
            "real/x",
            McpConfigEditor.ExpandEnvironmentReferences("${SET:-fallback}/x", Env(("SET", "real"))));
    }

    [Fact]
    public void UnsetWithoutADefaultExpandsToNothing()
    {
        Assert.Equal("a/b", McpConfigEditor.ExpandEnvironmentReferences("a/${NOPE}b", Env()));
    }

    [Theory]
    // A bare $, an unterminated ${ and an empty name are literal text, not errors: this
    // value is about to be compared against a real file, so guessing would be worse.
    [InlineData("plain/path.exe")]
    [InlineData("costs $5")]
    [InlineData("${unterminated")]
    [InlineData("${}")]
    public void NonReferencesAreCopiedThrough(string value)
    {
        Assert.Equal(value, McpConfigEditor.ExpandEnvironmentReferences(value, Env(("unterminated", "X"))));
    }

    [Fact]
    public void ExpandsSeveralReferencesInOneValue()
    {
        Assert.Equal(
            "A-B",
            McpConfigEditor.ExpandEnvironmentReferences("${ONE}-${TWO}", Env(("ONE", "A"), ("TWO", "B"))));
    }

    [Theory]
    [InlineData("${LOCALAPPDATA}/x", true)]
    [InlineData(@"C:\x\y.exe", false)]
    [InlineData("", false)]
    public void RecognizesValuesThatNeedExpanding(string value, bool expected)
    {
        Assert.Equal(expected, McpConfigEditor.ContainsEnvironmentReference(value));
    }

    // ===== Which spelling gets registered =====

    [Fact]
    public void DefaultInstallRegistersThePortableSpelling()
    {
        // The whole point: the entry keeps working after a roaming profile or a rename.
        Assert.Equal(
            McpConfigEditor.PortableInstalledCommand,
            McpConfigEditor.PreferredCommand(
                @"C:\Users\x\AppData\Local\OutlookAI\Setup\McpServer\OutlookAI.McpServer.exe",
                Env((@"LOCALAPPDATA", @"C:\Users\x\AppData\Local"))));
    }

    [Fact]
    public void ADeveloperBuildRegistersItsRealPath()
    {
        // A build output is not under %LOCALAPPDATA%\OutlookAI\Setup, so the portable form
        // would name a different file - register the truth instead.
        const string build = @"C:\Source\OutlookAI\McpServer\OutlookAI.McpServer\bin\Release\net10.0-windows\OutlookAI.McpServer.exe";

        Assert.Equal(build, McpConfigEditor.PreferredCommand(build, Env((@"LOCALAPPDATA", @"C:\Users\x\AppData\Local"))));
    }

    [Fact]
    public void WithoutLocalAppDataTheResolvedPathIsRegistered()
    {
        const string installed = @"C:\Users\x\AppData\Local\OutlookAI\Setup\McpServer\OutlookAI.McpServer.exe";

        Assert.Equal(installed, McpConfigEditor.PreferredCommand(installed, Env()));
    }

    [Fact]
    public void NoResolvedServerYieldsNoCommand()
    {
        Assert.Equal("", McpConfigEditor.PreferredCommand("", Env()));
    }

    // The opt-in toggle, the migration rule and every "may Outlook change this by itself?"
    // question moved to McpRegistrationDecision — see McpRegistrationDecisionTests.

    // ===== Project scope: creating .mcp.json =====

    [Theory]
    // Only for a file that is genuinely ABSENT — the caller reads nothing because there was
    // nothing to read. A file that is on disk never reaches this branch; see the
    // "exists but reads empty" section below for why that distinction is load-bearing.
    [InlineData("")]
    [InlineData("   \n\t ")]
    public void CreatesAWholeNewProjectFileWhenThereIsNoneOnDisk(string raw)
    {
        Assert.True(McpConfigEditor.TryBuildProjectConfig(raw, fileExists: false, Command, out string updated, out string error));
        Assert.Equal("", error);
        AssertJson(
            """
            {
              "mcpServers": {
                "outlookai": {
                  "type": "stdio",
                  "command": "${LOCALAPPDATA}/OutlookAI/Setup/McpServer/OutlookAI.McpServer.exe",
                  "args": [],
                  "env": {}
                }
              }
            }

            """,
            updated);
    }

    [Fact]
    public void ACreatedFileReadsBackAsWhatWasAskedFor()
    {
        McpConfigEditor.TryBuildProjectConfig("", fileExists: false, Command, out string updated, out _);

        Assert.True(McpConfigEditor.TryValidateJsonObject(updated, out _));
        Assert.True(McpConfigEditor.TryReadServerCommand(updated, out string command));
        Assert.Equal(Command, command);
        Assert.Equal(new[] { "outlookai" }, McpConfigEditor.ListServerNames(updated));
    }

    [Fact]
    public void EscapesABackslashPathIntoTheFile()
    {
        // The fallback spelling is a Windows path; unescaped it would not parse back.
        McpConfigEditor.TryBuildProjectConfig("", fileExists: false, @"C:\Program Files\OutlookAI\srv.exe", out string updated, out _);

        Assert.Contains(@"C:\\Program Files\\OutlookAI\\srv.exe", updated);
        Assert.True(McpConfigEditor.TryReadServerCommand(updated, out string command));
        Assert.Equal(@"C:\Program Files\OutlookAI\srv.exe", command);
    }

    [Fact]
    public void RefusesToWriteWithoutAServer()
    {
        Assert.False(McpConfigEditor.TryBuildProjectConfig("", fileExists: false, "", out string updated, out string error));
        Assert.Equal("", updated);
        Assert.NotEqual("", error);
    }

    // ===== Project scope: merging into an existing .mcp.json =====

    [Fact]
    public void MergesBesideAnotherServerAndKeepsItsLayout()
    {
        const string raw = """
        {
          "mcpServers": {
            "other": { "command": "other.exe" }
          }
        }
        """;

        Assert.True(McpConfigEditor.TryBuildProjectConfig(raw, fileExists: true, Command, out string updated, out string error));
        Assert.Equal("", error);
        AssertJson(
            """
            {
              "mcpServers": {
                "outlookai": { "type": "stdio", "command": "${LOCALAPPDATA}/OutlookAI/Setup/McpServer/OutlookAI.McpServer.exe", "args": [], "env": {} },
                "other": { "command": "other.exe" }
              }
            }
            """,
            updated);
    }

    [Fact]
    public void KeepsEveryOtherServerAndSettingWhenMerging()
    {
        const string raw = """
        {
          "$schema": "https://example.invalid/mcp.json",
          "mcpServers": {
            "postgres": { "command": "npx", "args": ["-y", "@modelcontextprotocol/server-postgres"] },
            "github": { "type": "http", "url": "https://api.example.invalid/mcp" }
          }
        }
        """;

        Assert.True(McpConfigEditor.TryBuildProjectConfig(raw, fileExists: true, Command, out string updated, out _));

        Assert.Contains(@"""-y"", ""@modelcontextprotocol/server-postgres""", updated);
        Assert.Contains(@"https://api.example.invalid/mcp", updated);
        Assert.Contains(@"""$schema"": ""https://example.invalid/mcp.json""", updated);
        Assert.Equal(
            new[] { "outlookai", "postgres", "github" },
            McpConfigEditor.ListServerNames(updated));
        Assert.Equal(new[] { "$schema", "mcpServers" }, McpConfigEditor.ListTopLevelKeys(updated));
    }

    [Fact]
    public void AddsTheServersBlockWhenTheFileHasNone()
    {
        const string raw = """
        {
          "somethingElse": true
        }
        """;

        Assert.True(McpConfigEditor.TryBuildProjectConfig(raw, fileExists: true, Command, out string updated, out _));
        AssertJson(
            """
            {
              "mcpServers": { "outlookai": { "type": "stdio", "command": "${LOCALAPPDATA}/OutlookAI/Setup/McpServer/OutlookAI.McpServer.exe", "args": [], "env": {} } },
              "somethingElse": true
            }
            """,
            updated);
    }

    [Fact]
    public void HandlesAnEmptyServersBlock()
    {
        Assert.True(McpConfigEditor.TryBuildProjectConfig("""{"mcpServers":{}}""", fileExists: true, Command, out string updated, out _));

        Assert.Equal(
            """{"mcpServers":{"outlookai": { "type": "stdio", "command": "${LOCALAPPDATA}/OutlookAI/Setup/McpServer/OutlookAI.McpServer.exe", "args": [], "env": {} }}}""",
            updated);
    }

    [Fact]
    public void HandlesAnEmptyObject()
    {
        Assert.True(McpConfigEditor.TryBuildProjectConfig("{}", fileExists: true, Command, out string updated, out _));

        Assert.True(McpConfigEditor.TryReadServerCommand(updated, out string command));
        Assert.Equal(Command, command);
        Assert.Equal(new[] { "mcpServers" }, McpConfigEditor.ListTopLevelKeys(updated));
    }

    [Fact]
    public void ReplacesAnEntryThatIsAlreadyThere()
    {
        const string raw = """
        {
          "mcpServers": {
            "outlookai": { "type": "stdio", "command": "C:\\stale\\OutlookAI.McpServer.exe", "args": [], "env": {} }
          }
        }
        """;

        Assert.True(McpConfigEditor.TryBuildProjectConfig(raw, fileExists: true, Command, out string updated, out _));

        Assert.DoesNotContain("stale", updated);
        Assert.Equal(new[] { "outlookai" }, McpConfigEditor.ListServerNames(updated));
        Assert.True(McpConfigEditor.TryReadServerCommand(updated, out string command));
        Assert.Equal(Command, command);
    }

    [Fact]
    public void LeavesAnAlreadyCorrectEntryCompletelyAlone()
    {
        // Not one byte, not the formatting, not the extra keys someone added by hand.
        const string raw = """
        {
          "mcpServers": {
            "outlookai": {
              "type": "stdio",
              "command": "${LOCALAPPDATA}/OutlookAI/Setup/McpServer/OutlookAI.McpServer.exe",
              "args": [],
              "env": { "OUTLOOKAI_SOMETHING": "kept" }
            }
          }
        }
        """;

        Assert.True(McpConfigEditor.TryBuildProjectConfig(raw, fileExists: true, Command, out string updated, out _));
        Assert.Equal(raw, updated);
    }

    [Fact]
    public void RunningTheButtonTwiceChangesNothingTheSecondTime()
    {
        McpConfigEditor.TryBuildProjectConfig("", fileExists: false, Command, out string once, out _);
        McpConfigEditor.TryBuildProjectConfig(once, fileExists: true, Command, out string twice, out _);

        // Byte-identical: the caller uses exactly this to decide not to write, so the
        // project's source control sees no pointless diff.
        Assert.Equal(once, twice);
    }

    [Fact]
    public void DoesNotMistakeANestedServersBlockForTheTopLevelOne()
    {
        const string raw = """
        {
          "wrapper": { "mcpServers": { "outlookai": { "command": "NESTED.exe" } } },
          "mcpServers": { "other": { "command": "other.exe" } }
        }
        """;

        Assert.True(McpConfigEditor.TryBuildProjectConfig(raw, fileExists: true, Command, out string updated, out _));

        Assert.Contains("NESTED.exe", updated);
        Assert.Equal(new[] { "outlookai", "other" }, McpConfigEditor.ListServerNames(updated));
    }

    // ===== Project scope: refusing =====

    [Theory]
    // Anything we cannot read is left alone: rewriting it would cost the user whatever is
    // in it, and "it looked close enough" is exactly how that happens.
    [InlineData("this is not json")]
    [InlineData("{ \"mcpServers\": { } ")]                        // truncated
    [InlineData("{ \"mcpServers\": {}, }")]                       // trailing comma
    [InlineData("{ mcpServers: {} }")]                            // unquoted key
    [InlineData("{ 'mcpServers': {} }")]                          // single quotes
    [InlineData("{ \"mcpServers\": {} } trailing junk")]
    [InlineData("{ \"a\": 01 }")]                                 // leading zero
    [InlineData("{ \"a\": \"bad \\x escape\" }")]
    [InlineData("[ { \"mcpServers\": {} } ]")]                    // not an object
    [InlineData("// a comment\n{ \"mcpServers\": {} }")]
    public void RefusesAFileItCannotRead(string raw)
    {
        Assert.False(McpConfigEditor.TryBuildProjectConfig(raw, fileExists: true, Command, out string updated, out string error));
        Assert.Equal("", updated);
        Assert.NotEqual("", error);
    }

    [Fact]
    public void RefusesWhenTheServersSettingIsNotAnObject()
    {
        Assert.False(McpConfigEditor.TryBuildProjectConfig(
            """{"mcpServers": "nonsense"}""", fileExists: true, Command, out string updated, out string error));
        Assert.Equal("", updated);
        Assert.Contains("mcpServers", error);
    }

    // ===== A config that EXISTS but reads back empty =====
    //
    // The data-loss case both scopes have to survive. Reads are opened with
    // FileShare.ReadWrite so the Claude Code CLI holding its own config open does not fail
    // us — and that is exactly what makes the CLI's rewrite window observable: it truncates
    // the file, then flushes the new content, and a read landing in between comes back with
    // zero bytes. Scoring that as "nothing configured yet" would have us write a fresh
    // one-property document, and the atomic replace would move the user's real configuration
    // aside into a backup the CLI's own open handle then overwrites. Nothing recoverable is
    // left. So: absent means create, empty-but-present means REFUSE.

    [Theory]
    [InlineData("")]
    [InlineData(" ")]
    [InlineData("   \r\n\t ")]
    [InlineData(null)]
    public void AFileOnDiskThatReadsEmptyIsAFailedRead(string? raw)
    {
        Assert.True(McpConfigEditor.ExistsButReadsEmpty(fileExists: true, raw!));
    }

    [Theory]
    // No file at all: nothing was read because there is nothing there. The only state that
    // may be created from scratch, in either scope.
    [InlineData(false, "")]
    [InlineData(false, "   ")]
    [InlineData(false, null)]
    // A file with any content is simply a file to splice into - including one that is
    // unreadable for some other reason, which the parse gates refuse on their own terms.
    [InlineData(true, "{}")]
    [InlineData(true, "not json at all")]
    [InlineData(true, " {} ")]
    public void EverythingElseIsNotAFailedRead(bool fileExists, string? raw)
    {
        Assert.False(McpConfigEditor.ExistsButReadsEmpty(fileExists, raw!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("   \r\n\t ")]
    public void RefusesToRebuildAProjectFileThatExistsButReadsEmpty(string raw)
    {
        // Nothing is produced for the caller to write, and the reason says so.
        Assert.False(McpConfigEditor.TryBuildProjectConfig(
            raw, fileExists: true, Command, out string updated, out string error));

        Assert.Equal("", updated);
        Assert.NotEqual("", error);
        Assert.Contains("empty", error);
    }

    [Fact]
    public void TheSameTextCreatesAFileWhenThereIsNoneOnDisk()
    {
        // The other half of the rule, side by side with it: identical text, opposite verdict,
        // and the ONLY thing separating them is whether the file is there.
        Assert.False(McpConfigEditor.TryBuildProjectConfig("", fileExists: true, Command, out _, out _));
        Assert.True(McpConfigEditor.TryBuildProjectConfig("", fileExists: false, Command, out string created, out _));

        Assert.True(McpConfigEditor.TryReadServerCommand(created, out string command));
        Assert.Equal(Command, command);
    }

    // ===== Turning the toggle off: removal =====

    [Fact]
    public void RemovesOurEntryAndLeavesTheOthersAlone()
    {
        const string raw = """
        {
          "numStartups": 42,
          "projects": {
            "C:\\some\\project": { "mcpServers": { "outlookai": { "command": "NESTED.exe" } } }
          },
          "mcpServers": {
            "outlookai": { "type": "stdio", "command": "TOP.exe", "args": [], "env": {} },
            "other": { "command": "other.exe" }
          }
        }
        """;

        Assert.True(McpConfigEditor.TryBuildConfigWithoutServer(raw, out string updated, out bool changed, out string error));
        Assert.True(changed);
        Assert.Equal("", error);
        AssertJson(
            """
            {
              "numStartups": 42,
              "projects": {
                "C:\\some\\project": { "mcpServers": { "outlookai": { "command": "NESTED.exe" } } }
              },
              "mcpServers": {
                "other": { "command": "other.exe" }
              }
            }
            """,
            updated);
    }

    [Fact]
    public void RemovesTheLastEntryTogetherWithThePrecedingComma()
    {
        const string raw = """
        {
          "mcpServers": {
            "other": { "command": "other.exe" },
            "outlookai": { "type": "stdio", "command": "TOP.exe", "args": [], "env": {} }
          }
        }
        """;

        Assert.True(McpConfigEditor.TryBuildConfigWithoutServer(raw, out string updated, out bool changed, out _));
        Assert.True(changed);
        AssertJson(
            """
            {
              "mcpServers": {
                "other": { "command": "other.exe" }
              }
            }
            """,
            updated);
    }

    [Fact]
    public void RemovingTheOnlyEntryLeavesTheBlockEmptyNotBroken()
    {
        Assert.True(McpConfigEditor.TryBuildConfigWithoutServer(
            """{"numStartups":1,"mcpServers":{"outlookai":{"command":"x.exe"}}}""",
            out string updated, out bool changed, out _));

        Assert.True(changed);
        Assert.Equal("""{"numStartups":1,"mcpServers":{}}""", updated);
        Assert.Empty(McpConfigEditor.ListServerNames(updated));
        Assert.Equal(new[] { "numStartups", "mcpServers" }, McpConfigEditor.ListTopLevelKeys(updated));
    }

    [Theory]
    // Nothing of ours to remove: success, no change, and therefore no write at all.
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("{}")]
    [InlineData("""{"mcpServers":{}}""")]
    [InlineData("""{"mcpServers":{"other":{"command":"other.exe"}}}""")]
    [InlineData("""{"numStartups":1}""")]
    [InlineData("""{"mcpServers":"not-an-object"}""")]
    public void RemovingWhatIsNotThereChangesNothing(string raw)
    {
        Assert.True(McpConfigEditor.TryBuildConfigWithoutServer(raw, out string updated, out bool changed, out string error));

        Assert.False(changed);
        Assert.Equal(raw, updated);
        Assert.Equal("", error);
    }

    [Theory]
    // A member named outlookai whose value is not an OBJECT is not something we could have
    // written, and the presence detector does not count it either - so on a fresh install
    // (opted out by default, which is precisely when this removal path runs) it must survive
    // untouched. Matching by name alone here would silently delete a user's hand-written
    // entry the moment Outlook started.
    [InlineData("""{"mcpServers":{"outlookai":"x"}}""")]
    [InlineData("""{"mcpServers":{"outlookai":42}}""")]
    [InlineData("""{"mcpServers":{"outlookai":null}}""")]
    [InlineData("""{"mcpServers":{"outlookai":true}}""")]
    [InlineData("""{"mcpServers":{"outlookai":["x"]}}""")]
    [InlineData("""{"mcpServers":{"outlookai":"x","other":{"command":"o.exe"}}}""")]
    public void RemovalLeavesAnEntryThePresenceDetectorDoesNotClaim(string raw)
    {
        Assert.True(McpConfigEditor.TryBuildConfigWithoutServer(raw, out string updated, out bool changed, out string error));

        Assert.False(changed);
        Assert.Equal(raw, updated);
        Assert.Equal("", error);
    }

    [Fact]
    public void RemovalStillTakesARealEntryBesideAMalformedNeighbour()
    {
        // The refusal above is about the value's SHAPE, not about giving up on the file:
        // a proper object entry is still removed however odd its neighbours are.
        Assert.True(McpConfigEditor.TryBuildConfigWithoutServer(
            """{"mcpServers":{"other":"not-an-object","outlookai":{"command":"x.exe"}}}""",
            out string updated, out bool changed, out _));

        Assert.True(changed);
        Assert.Equal("""{"mcpServers":{"other":"not-an-object"}}""", updated);
    }

    [Fact]
    public void RemovalRefusesAFileItCannotRead()
    {
        const string raw = "{ \"mcpServers\": { \"outlookai\": { \"command\": \"x.exe\" } }, }";

        Assert.False(McpConfigEditor.TryBuildConfigWithoutServer(raw, out string updated, out bool changed, out string error));
        Assert.False(changed);
        Assert.Equal(raw, updated);
        Assert.NotEqual("", error);
    }

    [Fact]
    public void RemovalIsIdempotent()
    {
        McpConfigEditor.TryBuildConfigWithoutServer(
            """{"mcpServers":{"outlookai":{"command":"x.exe"},"other":{"command":"o.exe"}}}""",
            out string once, out _, out _);
        McpConfigEditor.TryBuildConfigWithoutServer(once, out string twice, out bool changedAgain, out _);

        Assert.False(changedAgain);
        Assert.Equal(once, twice);
    }

    // ===== The comparison the verification gates lean on =====

    [Fact]
    public void SameNamesInADifferentOrderCompareEqual()
    {
        // Order must NOT matter: a splice inserts our member at the front of the block, so
        // the names legitimately come back in a different order than they went in.
        Assert.True(McpConfigEditor.SameMultiset(
            new List<string> { "outlookai", "postgres", "github" },
            new List<string> { "github", "outlookai", "postgres" }));
    }

    [Fact]
    public void DifferentMultiplicitiesOfTheSameNamesCompareUnequal()
    {
        // THE reason this is a multiset comparison. A count-plus-Contains test passes this:
        // same length, and every name on the left does appear on the right. This is
        // verification code standing in front of a write, where a check that the wrong file
        // can satisfy is worse than no check at all.
        Assert.False(McpConfigEditor.SameMultiset(
            new List<string> { "a", "a", "b" },
            new List<string> { "a", "b", "b" }));
    }

    [Fact]
    public void RepeatedNamesAreMatchedOneForOne()
    {
        Assert.True(McpConfigEditor.SameMultiset(
            new List<string> { "a", "a", "b" },
            new List<string> { "b", "a", "a" }));
        Assert.False(McpConfigEditor.SameMultiset(
            new List<string> { "a", "a" },
            new List<string> { "a", "b" }));
    }

    [Fact]
    public void LengthAndMembershipAreBothRequired()
    {
        Assert.True(McpConfigEditor.SameMultiset(new List<string>(), new List<string>()));
        Assert.False(McpConfigEditor.SameMultiset(new List<string> { "a" }, new List<string>()));
        Assert.False(McpConfigEditor.SameMultiset(new List<string>(), new List<string> { "a" }));
        Assert.False(McpConfigEditor.SameMultiset(new List<string> { "a" }, new List<string> { "A" }));
    }

    // ===== Readers the verification gates depend on =====

    [Fact]
    public void ReadsTheRegisteredCommandFromARealisticConfig()
    {
        const string json = """
        {
          "numStartups": 42,
          "projects": { "C:\\p": { "mcpServers": { "outlookai": { "command": "WRONG.exe" } } } },
          "mcpServers": {
            "outlookai": { "type": "stdio", "command": "C:\\Program Files\\OutlookAI\\srv.exe", "args": [], "env": {} }
          }
        }
        """;

        Assert.True(McpConfigEditor.TryReadServerCommand(json, out string command));
        Assert.Equal(@"C:\Program Files\OutlookAI\srv.exe", command);
    }

    [Theory]
    [InlineData("")]
    [InlineData("{}")]
    [InlineData("""{"mcpServers":{}}""")]
    [InlineData("""{"mcpServers":{"outlookai":{}}}""")]
    [InlineData("""{"mcpServers":{"outlookai":{"command":123}}}""")]
    [InlineData("""{"mcpServers":{"outlookai":"not-an-object"}}""")]
    public void ReadsNoCommandWhenThereIsNone(string json)
    {
        Assert.False(McpConfigEditor.TryReadServerCommand(json, out string command));
        Assert.Equal("", command);
    }

    [Fact]
    public void FindsOnlyTheTopLevelSpanOfARepeatedKeyName()
    {
        const string json = """{"projects":{"p":{"mcpServers":{"a":1}}},"mcpServers":{"b":2}}""";

        Assert.True(McpConfigEditor.TryFindTopLevelValueSpan(json, "mcpServers", out int start, out int end));
        Assert.Equal("""{"b":2}""", json.Substring(start, end - start));
    }

    // ===== The validator in front of every write =====

    [Fact]
    public void AcceptsARealisticConfig()
    {
        const string json = """
        {
          "numStartups": 42,
          "autoUpdates": true,
          "ratio": -12.5e3,
          "nothing": null,
          "list": [1, 2, {"nested": "}"}],
          "quoted": "he said \"{\" and left \u0041",
          "mcpServers": { "outlookai": { "command": "x.exe", "args": [], "env": {} } }
        }
        """;

        Assert.True(McpConfigEditor.TryValidateJsonObject(json, out string error));
        Assert.Equal("", error);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("[1,2,3]")]
    [InlineData("\"a string\"")]
    [InlineData("{")]
    [InlineData("{\"a\":}")]
    [InlineData("{\"a\":1,}")]
    [InlineData("{\"a\" 1}")]
    [InlineData("{\"a\":1}{\"b\":2}")]
    [InlineData("{\"a\":\"line\nbreak\"}")]
    [InlineData("{\"a\":+1}")]
    [InlineData("{\"a\":1.}")]
    [InlineData("{\"a\":tru}")]
    public void RejectsAnythingThatIsNotOneWholeJsonObject(string json)
    {
        Assert.False(McpConfigEditor.TryValidateJsonObject(json, out string error));
        Assert.NotEqual("", error);
    }

    [Fact]
    public void DeeplyNestedInputTerminatesInsteadOfBlowingTheStack()
    {
        string json = "{\"a\":" + new string('[', 5000) + new string(']', 5000) + "}";

        Assert.False(McpConfigEditor.TryValidateJsonObject(json, out _));
    }
}
