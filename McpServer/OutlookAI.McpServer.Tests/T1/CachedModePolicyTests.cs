using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml.Linq;

using OutlookAI.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The five Cached Mode policy values - <c>Services\CachedModePolicy.cs</c>, LINKED - and how the
/// elevated helper that writes them ships (Q128).
///
/// <para>
/// The list is the whole of what a UAC approval of <c>OutlookAI.PolicyWriter.exe</c> can be used
/// for, so it is pinned value by value: a name or a number added here is an elevated write added
/// to the product, and that should be a visible change to a test, not a quiet one.
/// </para>
///
/// <para>
/// The packaging half reads the repository's own files, because the add-in and the helper are
/// MSBuild-only and nothing in this suite can build them: the helper must be built by the add-in's
/// build, ship beside it, be named what the add-in launches, run only elevated, and never touch
/// HKEY_CURRENT_USER - which, under an administrator's credentials, is the administrator's hive.
/// </para>
/// </summary>
public sealed class CachedModePolicyTests
{
    [Fact]
    public void TheFiveNames_InCatalogOrder()
    {
        Assert.Equal(
            new[] { "SyncWindowSetting", "SyncWindowSettingDays", "DownloadSharedFolders", "CacheOthersMail", "DisableSyncSliderForSharedMailbox" },
            CachedModePolicy.ValueNames);
    }

    [Theory]
    [InlineData("SyncWindowSetting", 0, new[] { 0, 1, 3, 6, 12, 24 })]
    [InlineData("SyncWindowSettingDays", 0, new[] { 0, 3, 7, 14 })]
    [InlineData("DownloadSharedFolders", 1, new[] { 1, 0 })]
    [InlineData("CacheOthersMail", 1, new[] { 1, 0 })]
    [InlineData("DisableSyncSliderForSharedMailbox", 1, new[] { 1, 0 })]
    public void EachValue_ItsShippedDefault_AndExactlyTheChoicesOffered(string name, int shippedDefault, int[] offered)
    {
        Assert.Equal(shippedDefault, CachedModePolicy.ShippedDefault(name));
        Assert.True(CachedModePolicy.IsAllowed(name, shippedDefault), "the shipped default must be a value the helper accepts");
        Assert.Equal(offered, CachedModePolicy.Choices(name).Select(c => c.Value));
        Assert.False(string.IsNullOrWhiteSpace(CachedModePolicy.Meaning(name)));
        Assert.All(CachedModePolicy.Choices(name), c => Assert.False(string.IsNullOrWhiteSpace(c.Label)));
    }

    [Theory]
    [InlineData("syncwindowsetting")]
    [InlineData("SYNCWINDOWSETTING")]
    [InlineData(" SyncWindowSetting")]
    [InlineData("SyncWindowSetting​")]
    [InlineData("Cached Mode")]
    [InlineData("")]
    public void NamesAreOrdinal_ACaseVariantIsNotOneOfTheFive(string name)
    {
        Assert.False(CachedModePolicy.IsKnownName(name));
        Assert.False(CachedModePolicy.IsAllowed(name, 0));
    }

    [Fact]
    public void AnUnknownName_Throws_WhereAnswerWouldBeMadeUp()
    {
        Assert.Throws<ArgumentException>(() => CachedModePolicy.ShippedDefault("NoOST"));
        Assert.Throws<ArgumentException>(() => CachedModePolicy.Choices("NoOST"));
    }

    [Theory]
    [InlineData("16.0", @"Software\Policies\Microsoft\Office\16.0\Outlook\Cached Mode")]
    [InlineData("15.0", @"Software\Policies\Microsoft\Office\15.0\Outlook\Cached Mode")]
    [InlineData("17.0", @"Software\Policies\Microsoft\Office\17.0\Outlook\Cached Mode")]
    public void TheKey_IsThePoliciesMirrorOfOutlooksCachedModeKey(string version, string expected)
    {
        Assert.True(CachedModePolicy.IsSupportedOfficeVersion(version));
        Assert.Equal(expected, CachedModePolicy.KeyPath(version));
    }

    [Theory]
    [InlineData("16")]
    [InlineData("14.0")]
    [InlineData("18.0")]
    [InlineData(" 16.0")]
    [InlineData("")]
    public void OnlyTheSupportedOfficeMajors(string version)
    {
        Assert.False(CachedModePolicy.IsSupportedOfficeVersion(version));
    }

    [Fact]
    public void Describe_SaysWhatANumberMeans_AndFlagsOneOutlookAIDoesNotOffer()
    {
        Assert.Equal("(not set)", CachedModePolicy.Describe("SyncWindowSetting", null));
        Assert.Equal("All (0)", CachedModePolicy.Describe("SyncWindowSetting", 0));
        Assert.Equal("2 weeks (14)", CachedModePolicy.Describe("SyncWindowSettingDays", 14));
        Assert.Equal("On (1)", CachedModePolicy.Describe("CacheOthersMail", 1));
        Assert.Equal("36 (not a value OutlookAI offers)", CachedModePolicy.Describe("SyncWindowSetting", 36));
    }

    // ===== How the helper ships =====

    [Fact]
    public void TheAddInLaunches_TheExecutableTheHelperProjectBuilds()
    {
        XDocument helper = XDocument.Load(Path.Combine(RepoRoot(), "PolicyWriter", "OutlookAI.PolicyWriter.csproj"));
        XNamespace ns = helper.Root!.Name.Namespace;

        Assert.Equal(PolicyWriterCommandLine.ExecutableName, helper.Descendants(ns + "AssemblyName").Single().Value + ".exe");
        Assert.Equal("WinExe", helper.Descendants(ns + "OutputType").Single().Value);
        Assert.Equal("v4.8", helper.Descendants(ns + "TargetFrameworkVersion").Single().Value);
        Assert.Equal("app.manifest", helper.Descendants(ns + "ApplicationManifest").Single().Value);
    }

    [Fact]
    public void TheHelperRunsOnlyElevated()
    {
        XDocument manifest = XDocument.Load(Path.Combine(RepoRoot(), "PolicyWriter", "app.manifest"));
        XElement level = manifest.Descendants().Single(e => e.Name.LocalName == "requestedExecutionLevel");

        Assert.Equal("requireAdministrator", (string?)level.Attribute("level"));
        Assert.Equal("false", (string?)level.Attribute("uiAccess"));
    }

    [Fact]
    public void TheAddInBuildsTheHelper_AndShipsItBesideItself()
    {
        XDocument addIn = XDocument.Load(Path.Combine(RepoRoot(), "OutlookAI.csproj"));
        XNamespace ns = addIn.Root!.Name.Namespace;

        XElement reference = addIn.Descendants(ns + "ProjectReference")
            .Single(r => (string?)r.Attribute("Include") == @"PolicyWriter\OutlookAI.PolicyWriter.csproj");
        Assert.Equal("false", reference.Element(ns + "ReferenceOutputAssembly")?.Value);
        // Without it, the payload and release builds' global PrepareForRunDependsOn reaches the
        // helper, where VisualStudioForApplicationsBuild does not exist (MSB4057).
        Assert.Equal("PrepareForRunDependsOn", reference.Element(ns + "GlobalPropertiesToRemove")?.Value);

        XElement content = addIn.Descendants(ns + "Content")
            .Single(c => ((string?)c.Attribute("Include") ?? "").EndsWith(PolicyWriterCommandLine.ExecutableName, StringComparison.Ordinal));
        Assert.Equal(PolicyWriterCommandLine.ExecutableName, content.Element(ns + "Link")?.Value);
        Assert.Equal("PreserveNewest", content.Element(ns + "CopyToOutputDirectory")?.Value);
    }

    [Theory]
    [InlineData("PolicyWriterRun.cs")]
    [InlineData("WindowsPolicyHost.cs")]
    [InlineData("Program.cs")]
    public void TheHelperNeverTouchesHkeyCurrentUser(string file)
    {
        string source = File.ReadAllText(Path.Combine(RepoRoot(), "PolicyWriter", file));
        string code = string.Join("\n", source.Split('\n').Where(l => !l.TrimStart().StartsWith("//", StringComparison.Ordinal)));

        Assert.DoesNotContain("Registry.CurrentUser", code, StringComparison.Ordinal);
        Assert.DoesNotContain("RegistryHive.CurrentUser", code, StringComparison.Ordinal);
        Assert.DoesNotContain("Registry.LocalMachine", code, StringComparison.Ordinal);
        Assert.DoesNotContain("RegLoadKey", code, StringComparison.Ordinal);
    }

    private static string RepoRoot()
    {
        string testProjectDir =
            typeof(CachedModePolicyTests).Assembly
                .GetCustomAttributes<AssemblyMetadataAttribute>()
                .FirstOrDefault(a => a.Key == "TestProjectDir")?.Value
            ?? throw new InvalidOperationException("AssemblyMetadata 'TestProjectDir' is missing.");
        return Path.GetFullPath(Path.Combine(testProjectDir, "..", ".."));
    }
}
