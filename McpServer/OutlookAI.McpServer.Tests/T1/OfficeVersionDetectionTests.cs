using OutlookAI.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The Office major version every registry-backed answer in this server is read out of.
/// <para>
/// It used to be a hardcoded 16.0 in four places - the Outlook Search user key, that key's
/// Policies mirror, and the profile registry root behind accounts and signature defaults - so
/// on Outlook 2013 (15.0) or a future 17.0 the server read hives Outlook never writes and
/// reported empty accounts, empty signature assignments and a default-looking search setting.
/// Every one of those is a perfectly plausible answer on a healthy machine, which is exactly
/// what made the defect invisible.
/// </para>
/// <para>
/// THE POINT OF THE SEAM. This developer machine has ONE Office version really installed
/// (16.0 - though all three Outlook KEYS exist here, see below), so the 15.0, 17.0 and
/// nothing-found branches cannot be reached against the real registry without installing a
/// second Outlook. They are exercised here through the injected
/// "is there a real Outlook hive here?" predicate - the same shape as
/// <see cref="HealthReporting.ReadTuningState"/>'s value reader, and for the same reason.
/// The live-registry path is asserted separately, as a never-throws + agrees-with-itself
/// contract, because its ANSWER is machine-dependent.
/// </para>
/// </summary>
public sealed class OfficeVersionDetectionTests
{
    private readonly ITestOutputHelper _output;

    public OfficeVersionDetectionTests(ITestOutputHelper output)
    {
        _output = output;
    }

    /// <summary>A predicate that reports exactly the listed HKCU key paths as existing.</summary>
    private static Func<string, bool> HivesPresent(params string[] officeVersions)
    {
        HashSet<string> present = new HashSet<string>(
            officeVersions.Select(OutlookProfileRegistry.BuildOutlookRootKeyPath),
            StringComparer.OrdinalIgnoreCase);
        return path => present.Contains(path);
    }

    /// <summary>
    /// THE TRAP THIS DETECTION NEARLY WALKED INTO, pinned with the shapes measured on this
    /// machine on 2026-08-17.
    /// <para>
    /// <c>Installer.iss</c> writes
    /// <c>HKCU\Software\Microsoft\Office\{16.0,15.0,17.0}\Outlook\Resiliency\DoNotDisableAddinList</c>
    /// on every install, and writing that value creates the whole path above it. So on every
    /// machine this product is installed on, ALL THREE Outlook keys exist - verified here: 15.0
    /// and 17.0 each held exactly one subkey (<c>Resiliency</c>) and no values, while the real
    /// 16.0 held 31 subkeys and 4 values. A probe that asked only "does the key exist?" would
    /// have answered 16.0 on an Outlook 2013 machine just as confidently, which is the exact
    /// defect the detection was introduced to remove.
    /// </para>
    /// </summary>
    [Fact]
    public void IsOutlookHive_TellsARealHiveFromOurOwnInstallersFootprint()
    {
        // Measured 15.0 / 17.0 on this machine: our resiliency exemption and nothing else.
        Assert.False(OutlookProfileRegistry.IsOutlookHive(
            Array.Empty<string>(),
            new[] { OutlookProfileRegistry.InstallerFootprintSubKeyName }));

        // Nothing at all (a key some other installer touched and emptied).
        Assert.False(OutlookProfileRegistry.IsOutlookHive(Array.Empty<string>(), Array.Empty<string>()));

        // Measured 16.0 on this machine, trimmed to the load-bearing members: Outlook's own
        // profile state. Either a value or a non-Resiliency subkey is enough.
        Assert.True(OutlookProfileRegistry.IsOutlookHive(
            new[] { "DefaultProfile", "OutlookName" },
            new[] { "Profiles", "Options", "Search", "Setup", "Resiliency" }));
        Assert.True(OutlookProfileRegistry.IsOutlookHive(Array.Empty<string>(), new[] { "Profiles", "Resiliency" }));
        Assert.True(OutlookProfileRegistry.IsOutlookHive(new[] { "DefaultProfile" }, new[] { "Resiliency" }));

        // Case-insensitive, because the registry is.
        Assert.False(OutlookProfileRegistry.IsOutlookHive(Array.Empty<string>(), new[] { "RESILIENCY" }));

        // Null-tolerant: an unreadable enumeration must not decide "yes".
        Assert.False(OutlookProfileRegistry.IsOutlookHive(null!, null!));
    }

    [Fact]
    public void SupportedVersions_AreTheAddInsList_NewestFirst()
    {
        // Shared FILE, not a mirror: Services\OfficeVersions.cs is linked into OutlookAI.Core,
        // so this asserts the add-in's own list. The ORDER is load-bearing - it is the probe
        // order - which is why this is a sequence comparison and not a set comparison.
        Assert.Equal(new[] { "16.0", "17.0", "15.0" }, OutlookProfileRegistry.SupportedOfficeVersions);
        Assert.Equal("16.0", OutlookProfileRegistry.FallbackOfficeVersion);
    }

    [Theory]
    [InlineData("16.0")]
    [InlineData("15.0")]
    [InlineData("17.0")]
    public void DetectOfficeVersion_FindsWhicheverSingleVersionIsInstalled(string installed)
    {
        Assert.Equal(installed, HealthReporting.DetectOfficeVersion(HivesPresent(installed)));
    }

    [Fact]
    public void DetectOfficeVersion_WithSeveralInstalled_TakesThemInProbeOrder()
    {
        // A machine carrying both Outlook 2013 and 2016 is treated as the newer of the two,
        // and a future 17.0 beside a 15.0 likewise - the order in Supported decides, so this
        // pins the whole ranking rather than one pair.
        Assert.Equal("16.0", HealthReporting.DetectOfficeVersion(HivesPresent("15.0", "16.0", "17.0")));
        Assert.Equal("16.0", HealthReporting.DetectOfficeVersion(HivesPresent("15.0", "16.0")));
        Assert.Equal("17.0", HealthReporting.DetectOfficeVersion(HivesPresent("15.0", "17.0")));
    }

    [Fact]
    public void DetectOfficeVersion_WithNoneInstalled_IsNullRatherThanAnException()
    {
        // The reportable state. Null (not a throw, not a silent 16.0) is what lets the health
        // report distinguish "this hive is empty" from "there is no such hive" - the whole
        // reason the field exists.
        Assert.Null(HealthReporting.DetectOfficeVersion(_ => false));
    }

    [Fact]
    public void DetectOfficeVersion_ProbesEveryVersion_ByItsFullHkcuPath()
    {
        List<string> probed = new List<string>();
        HealthReporting.DetectOfficeVersion(path =>
        {
            probed.Add(path);
            return false;
        });

        Assert.Equal(
            OutlookProfileRegistry.SupportedOfficeVersions
                .Select(OutlookProfileRegistry.BuildOutlookRootKeyPath)
                .ToArray(),
            probed);
        Assert.Contains(@"Software\Microsoft\Office\16.0\Outlook", probed);
        Assert.Contains(@"Software\Microsoft\Office\15.0\Outlook", probed);
        Assert.Contains(@"Software\Microsoft\Office\17.0\Outlook", probed);
    }

    [Fact]
    public void DetectOfficeVersion_WhenAProbeThrows_CarriesOnToTheNextVersion()
    {
        // A registry key we are not allowed to open tells us nothing about the NEXT one, and a
        // detection that gave up there would take the whole health report down with it.
        string? detected = HealthReporting.DetectOfficeVersion(path =>
            path.Contains("16.0", StringComparison.Ordinal)
                ? throw new System.Security.SecurityException("denied")
                : path.Contains("17.0", StringComparison.Ordinal));

        Assert.Equal("17.0", detected);
    }

    [Theory]
    [InlineData("15.0", @"Software\Microsoft\Office\15.0\Outlook")]
    [InlineData("16.0", @"Software\Microsoft\Office\16.0\Outlook")]
    [InlineData("17.0", @"Software\Microsoft\Office\17.0\Outlook")]
    public void OutlookRootKeyPath_IsBuiltFromTheDetectedVersion(string officeVersion, string expected)
    {
        Assert.Equal(expected, OutlookProfileRegistry.BuildOutlookRootKeyPath(officeVersion));
    }

    [Theory]
    [InlineData("15.0")]
    [InlineData("16.0")]
    [InlineData("17.0")]
    public void SearchKeyPaths_BothHives_FollowTheDetectedVersion(string officeVersion)
    {
        Assert.Equal(
            @"Software\Microsoft\Office\" + officeVersion + @"\Outlook\Search",
            HealthReporting.BuildOutlookSearchUserKeyPath(officeVersion));

        // The policy mirror lives under a DIFFERENT root (Software\Policies\...), which is the
        // server's own asymmetry - the add-in never touches the Policies hive - so it is built
        // separately and has to be pinned separately.
        Assert.Equal(
            @"Software\Policies\Microsoft\Office\" + officeVersion + @"\Outlook\Search",
            HealthReporting.BuildOutlookSearchPolicyKeyPath(officeVersion));
    }

    [Fact]
    public void LiveKeyPaths_AllAgreeOnOneOfficeVersion()
    {
        // Whatever this machine turns out to have, the four paths that used to say 16.0
        // independently must now be the same major. Machine-agnostic by construction: the
        // expectation is derived from the detection rather than written out.
        string version = OutlookProfileRegistry.OfficeVersion;

        Assert.Contains(version, OutlookProfileRegistry.SupportedOfficeVersions);
        Assert.Equal(
            OutlookProfileRegistry.BuildOutlookRootKeyPath(version),
            OutlookProfileRegistry.OutlookRootKeyPath);
        Assert.Equal(
            HealthReporting.BuildOutlookSearchUserKeyPath(version),
            HealthReporting.OutlookSearchUserKeyPath);
        Assert.Equal(
            HealthReporting.BuildOutlookSearchPolicyKeyPath(version),
            HealthReporting.OutlookSearchPolicyKeyPath);
    }

    [Fact]
    public void LiveDetection_NeverThrows_AndFallsBackVisibly()
    {
        // Machine-state agnostic: a developer box with Outlook reports its major; a CI runner
        // with no Office at all reports null and the paths fall back to 16.0. Either way the
        // read is exception-free (health must always produce a report) and the two answers
        // agree with each other.
        string? detected = HealthReporting.DetectedOfficeVersion();

        // Written out so a run on an unfamiliar machine SHOWS which hive it settled on, rather
        // than leaving it to be inferred. Visible with `--logger "console;verbosity=detailed"`.
        _output.WriteLine("detected Office major: " + (detected ?? "(none - falling back)"));
        _output.WriteLine("Outlook root in use  : " + OutlookProfileRegistry.OutlookRootKeyPath);
        _output.WriteLine("search user key      : " + HealthReporting.OutlookSearchUserKeyPath);
        _output.WriteLine("search policy key    : " + HealthReporting.OutlookSearchPolicyKeyPath);

        if (detected == null)
        {
            Assert.Equal(OutlookProfileRegistry.FallbackOfficeVersion, OutlookProfileRegistry.OfficeVersion);
        }
        else
        {
            Assert.Contains(detected, OutlookProfileRegistry.SupportedOfficeVersions);
            Assert.Equal(detected, OutlookProfileRegistry.OfficeVersion);
        }
    }

    [Fact]
    public void NoOfficeVersionProblem_SaysWhatWasLookedFor_AndWhyTheAnswersLookEmpty()
    {
        string problem = HealthReporting.NoOfficeVersionProblem;

        // Names every version it probed, so a reader can tell "unsupported Office" from
        // "supported Office, key not written yet" without reading the source.
        foreach (string version in OutlookProfileRegistry.SupportedOfficeVersions)
        {
            Assert.Contains(version, problem, StringComparison.Ordinal);
        }

        // Names the hive it fell back to, so the reader can go and look at it.
        Assert.Contains(OutlookProfileRegistry.OutlookRootKeyPath, problem, StringComparison.Ordinal);

        // And states the consequence in plain words - this is the sentence that turns
        // "everything is empty and I do not know why" into a diagnosis.
        Assert.Contains("EMPTY", problem, StringComparison.Ordinal);
        Assert.Contains("accounts", problem, StringComparison.Ordinal);
        Assert.Contains("signature", problem, StringComparison.OrdinalIgnoreCase);

        // House style for every user-visible string in this product: " - ", never an em dash.
        // Spelled as a code point so this file does not contain the character it forbids.
        Assert.DoesNotContain((char)0x2014, problem);
    }
}
