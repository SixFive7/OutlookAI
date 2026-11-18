using OutlookAI.Core.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pure logic behind outlook_health's <c>registration</c> block (Phase 8 / D6 v2 / R11):
/// the drift verdict, and the reading of mcpServers -> outlookai -> command out of Claude
/// Code's config text - through the add-in's own <c>McpConfigEditor</c> since stage S3 of the
/// code-sharing refactor (Q154), so the file the add-in writes is read here by the same
/// scanner and judged by the same rule. Its ${VAR} expansion is pinned in McpConfigEditorTests.
/// </summary>
public sealed class McpRegistrationHealthTests
{
    private const string ServerPath = @"C:\Users\x\AppData\Local\OutlookAI\Setup\McpServer\OutlookAI.McpServer.exe";

    // ===== DescribeMcpRegistration =====

    [Fact]
    public void SamePathRegistered_IsOk()
    {
        Assert.Equal(
            HealthReporting.RegistrationOk,
            HealthReporting.DescribeMcpRegistration(ServerPath, ServerPath));
    }

    [Theory]
    // Casing and separator differences still name the same file, so they are NOT drift.
    [InlineData(@"c:\users\x\appdata\local\outlookai\setup\mcpserver\outlookai.mcpserver.exe")]
    [InlineData(@"C:\Users\x\AppData\Local\OutlookAI\Setup\McpServer\..\McpServer\OutlookAI.McpServer.exe")]
    [InlineData(@"C:/Users/x/AppData/Local/OutlookAI/Setup/McpServer/OutlookAI.McpServer.exe")]
    public void EquivalentPathSpellings_AreOk(string registered)
    {
        Assert.Equal(
            HealthReporting.RegistrationOk,
            HealthReporting.DescribeMcpRegistration(registered, ServerPath));
    }

    [Fact]
    public void PortableRegistration_IsOkNotDrifted()
    {
        // What the add-in now registers: Claude Code expands ${VAR} when it reads the
        // config, so comparing the raw text would call a correct registration drift.
        Environment.SetEnvironmentVariable("OUTLOOKAI_TEST_ROOT", @"C:\Users\x\AppData\Local");
        try
        {
            Assert.Equal(
                HealthReporting.RegistrationOk,
                HealthReporting.DescribeMcpRegistration(
                    "${OUTLOOKAI_TEST_ROOT}/OutlookAI/Setup/McpServer/OutlookAI.McpServer.exe",
                    ServerPath));
        }
        finally
        {
            Environment.SetEnvironmentVariable("OUTLOOKAI_TEST_ROOT", null);
        }
    }

    [Fact]
    public void PortableRegistrationOfSomethingElse_IsStillDrifted()
    {
        Assert.Equal(
            HealthReporting.RegistrationDrifted,
            HealthReporting.DescribeMcpRegistration(
                "${OUTLOOKAI_TEST_UNSET}/OutlookAI.McpServer.exe",
                ServerPath));
    }

    [Fact]
    public void ARegistrationThatExpandsToNothing_IsDriftedNotOk()
    {
        // Never let an empty expansion be compared as a path; it names no file.
        Assert.Equal(
            HealthReporting.RegistrationDrifted,
            HealthReporting.DescribeMcpRegistration("${OUTLOOKAI_TEST_UNSET}", ServerPath));
    }

    [Fact]
    public void DifferentPathRegistered_IsDrifted()
    {
        // The exact drift this exists to catch: a registration left pointing at a build output.
        Assert.Equal(
            HealthReporting.RegistrationDrifted,
            HealthReporting.DescribeMcpRegistration(
                @"C:\Source\OutlookAI-v3\McpServer\OutlookAI.McpServer\bin\Release\net10.0-windows\OutlookAI.McpServer.exe",
                ServerPath));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void NoRegisteredCommand_IsAbsent(string? registered)
    {
        Assert.Equal(
            HealthReporting.RegistrationAbsent,
            HealthReporting.DescribeMcpRegistration(registered, ServerPath));
    }

    [Fact]
    public void UnknownRunningPath_IsUnknownNotDrifted()
    {
        // Failing to read our own process path is not evidence the registration is wrong.
        Assert.Equal(
            HealthReporting.RegistrationUnknown,
            HealthReporting.DescribeMcpRegistration(ServerPath, null));
    }

    // ===== ExtractRegisteredCommand =====

    [Fact]
    public void ExtractsCommandFromRealisticConfig()
    {
        const string json = """
        {
          "numStartups": 42,
          "mcpServers": {
            "outlookai": {
              "type": "stdio",
              "command": "C:\\Program Files\\OutlookAI\\McpServer\\OutlookAI.McpServer.exe",
              "args": [],
              "env": {}
            }
          },
          "projects": {}
        }
        """;

        string? command = HealthReporting.ExtractRegisteredCommand(json, out bool readable);

        Assert.True(readable);
        Assert.Equal(@"C:\Program Files\OutlookAI\McpServer\OutlookAI.McpServer.exe", command);
    }

    [Fact]
    public void IgnoresNestedKeysOfTheSameName()
    {
        // A project entry may carry its own mcpServers block. Only the top-level one is
        // the user-global registration, and a scanner that matched by name alone would
        // report the wrong path here.
        const string json = """
        {
          "projects": {
            "C:\\some\\project": {
              "mcpServers": { "outlookai": { "command": "WRONG.exe" } }
            }
          },
          "mcpServers": { "outlookai": { "command": "RIGHT.exe" } }
        }
        """;

        Assert.Equal("RIGHT.exe", HealthReporting.ExtractRegisteredCommand(json, out _));
    }

    [Fact]
    public void SkipsPrecedingMembersOfEveryShape()
    {
        // Arrays, nested objects, strings with braces, escaped quotes, numbers, literals -
        // each must be skipped whole before the member walk reaches mcpServers.
        const string json = """
        {
          "a": [1, 2, {"b": "}"}],
          "c": "he said \"{\" and left",
          "d": null,
          "e": true,
          "f": -12.5e3,
          "mcpServers": { "outlookai": { "command": "OK.exe" } }
        }
        """;

        Assert.Equal("OK.exe", HealthReporting.ExtractRegisteredCommand(json, out _));
    }

    [Theory]
    // Well-formed JSON objects that simply have no registration.
    [InlineData("{}")]
    [InlineData("""{"mcpServers":{}}""")]
    [InlineData("""{"mcpServers":{"other":{"command":"x.exe"}}}""")]
    [InlineData("""{"mcpServers":{"outlookai":{}}}""")]
    [InlineData("""{"mcpServers":{"outlookai":{"command":123}}}""")]
    [InlineData("""{"mcpServers":"not-an-object"}""")]
    public void WellFormedButUnregistered_IsReadableWithNoCommand(string json)
    {
        string? command = HealthReporting.ExtractRegisteredCommand(json, out bool readable);

        Assert.True(readable);
        Assert.Null(command);
    }

    [Theory]
    [InlineData("not json at all")]
    [InlineData("[1,2,3]")]
    [InlineData("\"a string\"")]
    public void NotAJsonObject_IsUnreadable(string json)
    {
        _ = HealthReporting.ExtractRegisteredCommand(json, out bool readable);

        Assert.False(readable);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void EmptyInput_IsReadableWithNoCommand(string? json)
    {
        // An absent/empty config is not a corrupt one - it simply has no entry, and must
        // never be reported as unreadable (which is the state that blocks repair).
        string? command = HealthReporting.ExtractRegisteredCommand(json, out bool readable);

        Assert.True(readable);
        Assert.Null(command);
    }

    [Theory]
    // Truncated input must terminate and refuse, never hang or throw.
    [InlineData("""{"mcpServers":{"outlookai":{"command":"unterminated""")]
    [InlineData("""{"mcpServers":{"outlookai":""")]
    [InlineData("{")]
    public void TruncatedInput_TerminatesWithoutThrowing(string json)
    {
        Assert.Null(HealthReporting.ExtractRegisteredCommand(json, out _));
    }

    [Theory]
    // A file caught mid-rewrite, or broken: the add-in will not touch it, and the report says
    // it could not read it instead of calling it unregistered. Until stage S3 the first two of
    // these read as registered - the command came before the break - and the last as "absent".
    [InlineData("""{"mcpServers":{"outlookai":{"command":"OK.exe"}}""")]
    [InlineData("""{"mcpServers":{"outlookai":{"command":"OK.exe"}}},""")]
    [InlineData("{")]
    public void AnObjectThatDoesNotParse_IsUnreadable(string json)
    {
        string? command = HealthReporting.ExtractRegisteredCommand(json, out bool readable);

        Assert.False(readable);
        Assert.Null(command);
    }

    // ===== ReadRegisteredCommand: the file, by the add-in's rule =====

    [Fact]
    public void AnAbsentFile_IsReadableWithNoCommand()
    {
        string? command = HealthReporting.ReadRegisteredCommand(fileExists: false, text: null, out bool readable);

        Assert.True(readable);
        Assert.Null(command);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData(null)]
    public void AFileThatIsThereButReadsEmpty_IsUnreadable(string? text)
    {
        // The Claude Code CLI truncates the file before it flushes, and the read is shared with
        // it, so an empty read of a file that exists is a failed read: the add-in never writes
        // over it, and outlook_health says unreadable. Until stage S3 it said "absent" here.
        string? command = HealthReporting.ReadRegisteredCommand(fileExists: true, text, out bool readable);

        Assert.False(readable);
        Assert.Null(command);
    }

    [Fact]
    public void AFileThatHoldsTheRegistration_IsRead()
    {
        string? command = HealthReporting.ReadRegisteredCommand(
            fileExists: true, """{"mcpServers":{"outlookai":{"command":"OK.exe"}}}""", out bool readable);

        Assert.True(readable);
        Assert.Equal("OK.exe", command);
    }

    [Fact]
    public void DecodesEscapesInTheCommand()
    {
        const string json = """{"mcpServers":{"outlookai":{"command":"C:\\a b\\\u0073rv.exe"}}}""";

        Assert.Equal(@"C:\a b\srv.exe", HealthReporting.ExtractRegisteredCommand(json, out _));
    }

    // ===== Probes =====

    [Fact]
    public void MachineProbes_NeverThrow()
    {
        // Impure probes: values are machine-dependent, absence of exceptions is the contract.
        McpRegistrationHealthView view = HealthReporting.ReadMcpRegistration(HealthReporting.CurrentProcessPath());

        Assert.NotNull(view);
        Assert.Contains(view.Status, new[]
        {
            HealthReporting.RegistrationOk,
            HealthReporting.RegistrationDrifted,
            HealthReporting.RegistrationAbsent,
            HealthReporting.RegistrationUnreadable,
            HealthReporting.RegistrationUnknown,
        });
    }

    [Fact]
    public void CurrentProcessPath_NamesAnExecutable()
    {
        string? path = HealthReporting.CurrentProcessPath();

        Assert.False(string.IsNullOrWhiteSpace(path));
        Assert.EndsWith(".exe", path, System.StringComparison.OrdinalIgnoreCase);
    }
}
