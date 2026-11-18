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
        Assert.Equal(new[] { "GroupPolicy", "Preference" }, LiveSignatureSettingsStep.Scenarios);
        Assert.Equal("OUTLOOKAI_LIVE_SIGNATURE_SETTINGS", LiveSignatureSettingsStep.VariableName);
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
