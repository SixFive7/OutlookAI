using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The pure half of the live runner's <c>-SignatureSettings</c> step as the live tier reads it (Q200 (a), Q201):
/// which scenario a run's variable names, and how a compose window is told apart against the signatures the step
/// names. The live test that uses them needs a guest whose Outlook started under the setting; these need nothing.
/// </summary>
public sealed class LiveSignatureSettingsStepTests
{
    private const string PolicyNew = "OutlookAI-McpTest-Q200-Forced-New";
    private const string PolicyReply = "OutlookAI-McpTest-Q200-Forced-Reply";
    private const string OwnNew = "OutlookAI-McpTest-Q200-Own-New";

    private static readonly IReadOnlyList<string> Candidates = new[] { PolicyNew, PolicyReply, OwnNew };

    [Theory]
    [InlineData(null, null)]
    [InlineData("", null)]
    [InlineData("   ", null)]
    [InlineData("GroupPolicy", "GroupPolicy")]
    [InlineData(" Preference ", "Preference")]
    [InlineData("PreferenceNoOwnNew", "PreferenceNoOwnNew")]
    [InlineData("NoSetting", "NoSetting")]
    public void TheVariable_NamesAScenario_OrNone(string? value, string? expected)
    {
        Assert.Equal(expected, LiveSignatureSettingsStep.Parse(value));
    }

    [Theory]
    [InlineData("groupPolicy")]
    [InlineData("CloudPolicy")]
    [InlineData("1")]
    public void AValueTheRunnerNeverSets_IsRefused_NotReadAsNoStep(string value)
    {
        InvalidOperationException refusal = Assert.Throws<InvalidOperationException>(() => LiveSignatureSettingsStep.Parse(value));
        Assert.Contains(LiveSignatureSettingsStep.VariableName, refusal.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void TheScenarios_AreTheOnesTheRunnerTakes()
    {
        Assert.Equal(new[] { "GroupPolicy", "Preference", "PreferenceNoOwnNew", "NoSetting" }, LiveSignatureSettingsStep.Scenarios);
        Assert.Equal("OUTLOOKAI_LIVE_SIGNATURE_SETTINGS", LiveSignatureSettingsStep.VariableName);
        Assert.Equal("OUTLOOKAI_LIVE_SIGNATURE_DIALOG", LiveSignatureSettingsStep.DialogVariableName);
    }

    [Theory]
    [InlineData("GroupPolicy", false, true, true)]
    [InlineData("Preference", true, false, true)]
    [InlineData("PreferenceNoOwnNew", false, true, true)]
    [InlineData("NoSetting", true, true, false)]
    public void TheHubsOwnDefaults_AndWhetherASettingIsMade_FollowTheScenario(string scenario, bool newMail, bool replies, bool writesSetting)
    {
        Assert.Equal((newMail, replies), LiveSignatureSettingsStep.HubOwnDefaults(scenario));
        Assert.Equal(writesSetting, LiveSignatureSettingsStep.WritesSetting(scenario));
    }

    [Fact]
    public void AnUnknownScenario_HasNoPlan()
    {
        Assert.Throws<ArgumentException>(() => LiveSignatureSettingsStep.HubOwnDefaults("CloudPolicy"));
        Assert.Throws<ArgumentException>(() => LiveSignatureSettingsStep.WritesSetting("CloudPolicy"));
    }

    [Fact]
    public void TheDialogsReading_IsReadWhole()
    {
        const string Json = """
            {
              "probe": "Measure-SignatureDialog",
              "refused": null,
              "dialogFound": true,
              "dialogClass": "bosa_sdm_Mso96",
              "closed": true,
              "account": { "type": "ComboBox", "name": "E-mail account:", "enabled": true, "value": "hub@vm.invalid" },
              "newMessages": { "type": "ComboBox", "name": "New messages:", "enabled": false, "value": "OutlookAI-McpTest-Q200-Forced-New" },
              "repliesForwards": { "type": "ComboBox", "name": "Replies/forwards:", "enabled": false, "value": "OutlookAI-McpTest-Q200-Forced-Reply" },
              "delete": { "type": "Button", "name": "Delete", "enabled": true, "value": null },
              "signatures": [
                { "name": "OutlookAI-McpTest-Q200-Forced-New", "selected": true, "deleteEnabled": false, "error": null },
                { "name": "OutlookAI-McpTest-Q200-Own-Reply", "selected": true, "deleteEnabled": true, "error": null }
              ],
              "otherSignatureItems": 1,
              "settledAfterSeconds": 1.5,
              "folderUnchanged": true,
              "folderChanges": [],
              "defaultsUnchanged": true,
              "defaultsChanges": [],
              "controls": [ { "type": "Button", "name": "Cancel", "enabled": true, "value": null } ],
              "errors": []
            }
            """;
        SignatureDialogReading reading = LiveSignatureSettingsStep.ParseDialogReading(Json);

        Assert.Null(reading.Refused);
        Assert.True(reading.DialogFound && reading.Closed);
        Assert.Equal("bosa_sdm_Mso96", reading.DialogClass);
        Assert.Equal("hub@vm.invalid", reading.Account!.Value);
        Assert.Equal(false, reading.NewMessages!.Enabled);
        Assert.Equal("OutlookAI-McpTest-Q200-Forced-Reply", reading.RepliesForwards!.Value);
        Assert.Equal(true, reading.Delete!.Enabled);
        Assert.Equal(new[] { false, true }, reading.Signatures.Select(s => s.DeleteEnabled!.Value));
        Assert.Equal(1, reading.OtherSignatureItems);
        Assert.Equal(1.5, reading.SettledAfterSeconds);
        Assert.Equal(true, reading.FolderUnchanged);
        Assert.Empty(reading.FolderChanges);
        Assert.Equal(true, reading.DefaultsUnchanged);
        Assert.Empty(reading.DefaultsChanges);
        Assert.Equal("Cancel", Assert.Single(reading.Controls).Name);
        Assert.Empty(reading.Errors);
    }

    [Fact]
    public void AListOfOne_WrittenAsItsItem_IsStillAList_AndWhatIsMissingIsUnread()
    {
        // Windows PowerShell 5.1's ConvertTo-Json can write a list of one as the item itself.
        const string Json = """
            {
              "refused": "Outlook is not running in this session",
              "dialogFound": false,
              "signatures": { "name": "OutlookAI-McpTest-X", "selected": false, "deleteEnabled": null, "error": "no pattern" },
              "folderChanges": "changed OutlookAI-McpTest-X.htm",
              "settledAfterSeconds": null,
              "defaultsUnchanged": false,
              "defaultsChanges": "changed OutlookAI-Tier\\00000003\\Reply-Forward Signature",
              "errors": "the dialog did not open"
            }
            """;
        SignatureDialogReading reading = LiveSignatureSettingsStep.ParseDialogReading(Json);

        Assert.Equal("Outlook is not running in this session", reading.Refused);
        Assert.False(reading.DialogFound || reading.Closed);
        Assert.Null(reading.NewMessages);
        Assert.Null(reading.FolderUnchanged);
        SignatureDialogEntry entry = Assert.Single(reading.Signatures);
        Assert.Null(entry.DeleteEnabled);
        Assert.Equal("no pattern", entry.Error);
        Assert.Equal("changed OutlookAI-McpTest-X.htm", Assert.Single(reading.FolderChanges));
        Assert.Null(reading.SettledAfterSeconds);
        Assert.Equal(false, reading.DefaultsUnchanged);
        Assert.Equal(@"changed OutlookAI-Tier\00000003\Reply-Forward Signature", Assert.Single(reading.DefaultsChanges));
        Assert.Equal("the dialog did not open", Assert.Single(reading.Errors));
        Assert.Empty(reading.Controls);
    }

    [Fact]
    public void TheSignatureRegion_SaysWhichWentIn()
    {
        Assert.Equal(PolicyNew, LiveSignatureSettingsStep.Classify(
            "Signature step marker: " + PolicyNew + "\r", "Hello\r" + PolicyNew + "\r", Candidates));
        Assert.Equal(OwnNew, LiveSignatureSettingsStep.Classify("x " + OwnNew, null, Candidates));
    }

    [Fact]
    public void WithNoRegion_TheDocumentIsRead_AndAnEmptyOneHasNoSignature()
    {
        Assert.Equal(PolicyReply, LiveSignatureSettingsStep.Classify(null, "Reply text\r" + PolicyReply + "\r> quoted", Candidates));
        Assert.Equal(PolicyReply, LiveSignatureSettingsStep.Classify("  \r", "Reply text\r" + PolicyReply, Candidates));
        Assert.Equal(LiveSignatureSettingsStep.NoSignature, LiveSignatureSettingsStep.Classify(null, "Just the body\r", Candidates));
        Assert.Equal(LiveSignatureSettingsStep.NoSignature, LiveSignatureSettingsStep.Classify(string.Empty, "Just the body", Candidates));
    }

    [Fact]
    public void ARegionHoldingNoneOfThem_IsAnotherSignature_AndTwoOfThemAreSeveral()
    {
        Assert.Equal(LiveSignatureSettingsStep.AnotherSignature, LiveSignatureSettingsStep.Classify("Kind regards, someone", "x", Candidates));
        Assert.Equal(LiveSignatureSettingsStep.SeveralSignatures, LiveSignatureSettingsStep.Classify(PolicyNew + " " + OwnNew, null, Candidates));
    }

    [Fact]
    public void ANameInsideALongerOne_IsTheLongerOnesText()
    {
        // A candidate that is part of another's name is not a second signature where only the longer one stands.
        string[] nested = { "OutlookAI-McpTest-Sig", "OutlookAI-McpTest-SigB" };
        Assert.Equal("OutlookAI-McpTest-SigB", LiveSignatureSettingsStep.Classify("OutlookAI-McpTest-SigB", null, nested));
        Assert.Equal("OutlookAI-McpTest-Sig", LiveSignatureSettingsStep.Classify("OutlookAI-McpTest-Sig only", null, nested));
    }

    [Fact]
    public void NothingToRead_IsNotRead_AndCaseCounts()
    {
        Assert.Equal(LiveSignatureSettingsStep.NotRead, LiveSignatureSettingsStep.Classify(null, null, Candidates));
        Assert.Equal(LiveSignatureSettingsStep.AnotherSignature,
            LiveSignatureSettingsStep.Classify(PolicyNew.ToUpperInvariant(), null, Candidates));
    }
}
