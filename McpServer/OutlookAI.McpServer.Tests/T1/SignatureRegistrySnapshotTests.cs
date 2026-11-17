using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The registry half of Mailbox Safety rule 6 (Q147): the snapshot the signature live suites take
/// before they run fails the run on ANY signature setting that is not back afterwards - changed,
/// added or removed - and names it by path only. Fed fake values; the live capture is the tier's.
/// </summary>
public sealed class SignatureRegistrySnapshotTests
{
    private const string Account = @"Software\Microsoft\Office\16.0\Outlook\Profiles\Tier\9375CFF0413111d3B88A00104B2A6676\00000004";

    [Fact]
    public void PutBackExactly_IsNoDifference()
    {
        SignatureRegistrySnapshot before = SignatureRegistrySnapshot.Capture(Values("Home", "Office", null));
        SignatureRegistrySnapshot after = SignatureRegistrySnapshot.Capture(Values("Home", "Office", null));

        Assert.Empty(before.Diff(after));
    }

    [Fact]
    public void ChangedAddedAndRemoved_AreEachNamed_ByPathOnly()
    {
        SignatureRegistrySnapshot before = SignatureRegistrySnapshot.Capture(Values("Home", "Office", null));
        SignatureRegistrySnapshot after = SignatureRegistrySnapshot.Capture(Values("OutlookAI-McpTest-Sig", null, "Tier"));

        IReadOnlyList<string> diff = before.Diff(after);

        Assert.Equal(
            new[]
            {
                @"ADDED: HKCU\Software\Microsoft\Office\16.0\Outlook|DefaultProfile",
                @"CHANGED: HKCU\" + Account + "|New Signature",
                @"REMOVED: HKCU\" + Account + "|Reply-Forward Signature",
            },
            diff);
        Assert.DoesNotContain(diff, d => d.Contains("OutlookAI-McpTest-Sig", StringComparison.Ordinal));
    }

    [Fact]
    public void TheSameTextWrittenAsAnotherKind_IsAChange()
    {
        // Outlook writes REG_BINARY UTF-16; a REG_SZ put back in its place is not "as it was".
        SignatureRegistrySnapshot binary = SignatureRegistrySnapshot.Capture(new (string, string, object?)[]
        {
            (Account, "New Signature", System.Text.Encoding.Unicode.GetBytes("Home\0")),
        });
        SignatureRegistrySnapshot text = SignatureRegistrySnapshot.Capture(new (string, string, object?)[]
        {
            (Account, "New Signature", "Home"),
        });

        Assert.Single(binary.Diff(text));
    }

    [Fact]
    public void ItCoversEverySettingThatDecidesASignature()
    {
        IReadOnlyList<(string Key, string Value)> locations = SignatureRegistrySnapshot.FixedLocations("16.0");

        Assert.Contains((@"Software\Microsoft\Office\16.0\Outlook", "DefaultProfile"), locations);
        Assert.Contains((@"Software\Microsoft\Office\16.0\Common\General", "Signatures"), locations);
        Assert.Contains((@"Software\Microsoft\Office\16.0\Common\MailSettings", "NewSignature"), locations);
        Assert.Contains((@"Software\Policies\Microsoft\Office\16.0\Common\MailSettings", "DisableSignatures"), locations);
        Assert.Contains((@"Software\Policies\Microsoft\Cloud\Office\16.0\Common\MailSettings", "ReplySignature"), locations);
        Assert.Contains((@"Software\Microsoft\Office\16.0\Outlook\Setup", "DisableRoamingSignatures"), locations);
        Assert.Contains((@"Software\Policies\Microsoft\Office\16.0\Outlook\Setup", "DisableRoamingSignaturesTemporaryToggle"), locations);
    }

    private static IEnumerable<(string, string, object?)> Values(string? newSignature, string? reply, string? defaultProfile)
    {
        return new (string, string, object?)[]
        {
            (@"Software\Microsoft\Office\16.0\Outlook", "DefaultProfile", defaultProfile),
            (Account, "New Signature", newSignature),
            (Account, "Reply-Forward Signature", reply),
        };
    }
}
