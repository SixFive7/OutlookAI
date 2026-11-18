using OutlookAI.Core.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Which signature a draft gets, and why (Q155 (c), Q156, Q157 (a), Q160) - the one decision the
/// COM host composes by and new_draft checks before any COM work. Every branch, on fakes: the running
/// profile against the default one (G1's rule), an Exchange account known by an alias (the "silently
/// no default" defect), policy, roaming scope, delegate mailboxes, "none", and the suggestion a draft
/// result carries.
/// </summary>
public sealed class SignatureSelectionTests : IDisposable
{
    private const string Jan = "jan@example.com";
    private readonly TempSignatureFolder _folder = new TempSignatureFolder();

    public SignatureSelectionTests()
    {
        foreach (string name in new[] { "Home", "Office", "Office short", "Corporate", "Nederlands", "Roam (jan@example.com)", "Roam (other@example.com)" })
        {
            _folder.Write(name + ".htm", "<html><body><p>" + name + "</p></body></html>", System.Text.Encoding.ASCII);
        }
    }

    public void Dispose() => _folder.Dispose();

    // ------------------------------------------------------------------ parsing

    [Theory]
    [InlineData(null, SignatureRequestKind.AccountDefault, null)]
    [InlineData("", SignatureRequestKind.AccountDefault, null)]
    [InlineData("  ", SignatureRequestKind.AccountDefault, null)]
    [InlineData("none", SignatureRequestKind.None, null)]
    [InlineData(" NONE ", SignatureRequestKind.None, null)]
    [InlineData("None", SignatureRequestKind.None, null)]
    [InlineData(" Office ", SignatureRequestKind.Named, "Office")]
    [InlineData("none of these", SignatureRequestKind.Named, "none of these")]
    public void TheArgument_LeftOut_None_OrAName(string? argument, SignatureRequestKind kind, string? name)
    {
        SignatureRequest request = SignatureRequest.Parse(argument);

        Assert.Equal(kind, request.Kind);
        Assert.Equal(name, request.Name);
    }

    // ------------------------------------------------------------------ left out: the effective default

    [Fact]
    public void LeftOut_TheRunningProfilesDefault_NotTheDefaultProfiles()
    {
        // G1's case: DefaultProfile names one profile, Outlook runs on another.
        SignatureEnvironment environment = Environment(Accounts()
            .Profile("Corpus", isDefault: true)
            .Add("Corpus", 1, FakeProfileAccountRegistry.Pop3(Jan, newSignature: "Home"))
            .Add("Tier", 1, FakeProfileAccountRegistry.Pop3(Jan, newSignature: "Office", replySignature: "Office short")));

        SignatureDecision newMail = Decide(environment, SignatureRequest.AccountDefault, false, running: "Tier");
        SignatureDecision reply = Decide(environment, SignatureRequest.AccountDefault, true, running: "Tier");

        Assert.Equal("Office", newMail.Name);
        Assert.Equal(SignatureDecision.SourceAccountDefault, newMail.Source);
        Assert.Equal("Tier", newMail.Profile);
        Assert.Contains("the profile Outlook is running, 'Tier'", newMail.Reason, StringComparison.Ordinal);
        Assert.Equal(_folder.PathOf("Office.htm"), newMail.FilePath, ignoreCase: true);
        Assert.Equal("Office short", reply.Name);
        Assert.True(newMail.IsDefault);
    }

    [Fact]
    public void LeftOut_WithoutARunningProfile_TheRecordedDefaultProfile_SaidSo()
    {
        SignatureEnvironment environment = Environment(Accounts()
            .Profile("Tier", isDefault: true)
            .Add("Corpus", 1, FakeProfileAccountRegistry.Pop3(Jan, newSignature: "Home"))
            .Add("Tier", 1, FakeProfileAccountRegistry.Pop3(Jan, newSignature: "Office")));

        SignatureDecision decision = Decide(environment, SignatureRequest.AccountDefault, false, running: null);

        Assert.Equal("Office", decision.Name);
        Assert.Contains("the default profile 'Tier'", decision.Reason, StringComparison.Ordinal);
    }

    [Fact]
    public void LeftOut_SeveralProfiles_NoneRecordedAsDefault_NoRunningOne_NoSignature_NeverTheFirstMatch()
    {
        // The merge defect: both profiles have the address, and the first one used to win.
        SignatureEnvironment environment = Environment(Accounts()
            .Add("A", 1, FakeProfileAccountRegistry.Pop3(Jan, newSignature: "Home"))
            .Add("B", 1, FakeProfileAccountRegistry.Pop3(Jan, newSignature: "Office")));

        SignatureDecision decision = Decide(environment, SignatureRequest.AccountDefault, false, running: null);

        Assert.Null(decision.Name);
        Assert.Equal(SignatureDecision.SourceNone, decision.Source);
        Assert.Null(decision.Profile);
        Assert.Contains("none of the 2 profiles is recorded as the default", decision.Reason, StringComparison.Ordinal);
    }

    [Fact]
    public void LeftOut_TheOnlyProfile_IsUsed()
    {
        SignatureEnvironment environment = Environment(Accounts().Add("Solo", 1, FakeProfileAccountRegistry.Pop3(Jan, newSignature: "Home")));

        Assert.Equal("Home", Decide(environment, SignatureRequest.AccountDefault, false, running: null).Name);
    }

    [Fact]
    public void LeftOut_AnExchangeAccountAddedUnderAnAlias_IsFoundByTheNameOutlookShows()
    {
        // The profile entry is named after the alias; Outlook's Account.SmtpAddress is the primary.
        SignatureEnvironment environment = Environment(Accounts()
            .Add("Outlook", 1, FakeProfileAccountRegistry.Exchange("jan@alias.example.com", newSignature: "Corporate")));

        SignatureDecision byAddressOnly = SignatureSelector.Decide(environment, SignatureRequest.AccountDefault,
            new SignatureSendingAccount("jan.modaal@example.com", null), false, "Outlook");
        SignatureDecision withName = SignatureSelector.Decide(environment, SignatureRequest.AccountDefault,
            new SignatureSendingAccount("jan.modaal@example.com", "jan@alias.example.com"), false, "Outlook");

        // Without the name it is NOT silent any more: the reason says no entry matched.
        Assert.Null(byAddressOnly.Name);
        Assert.Contains("no account of profile 'Outlook' matches jan.modaal@example.com", byAddressOnly.Reason, StringComparison.Ordinal);
        Assert.Equal("Corporate", withName.Name);
    }

    [Fact]
    public void LeftOut_NoDefault_OrItsFilesGone_NoSignature_WithTheReason()
    {
        SignatureEnvironment environment = Environment(Accounts()
            .Add("Outlook", 1, FakeProfileAccountRegistry.Pop3(Jan))
            .Add("Outlook", 2, FakeProfileAccountRegistry.Pop3("ghost@example.com", newSignature: "Deleted Long Ago")));

        SignatureDecision none = Decide(environment, SignatureRequest.AccountDefault, false, running: "Outlook");
        SignatureDecision gone = SignatureSelector.Decide(environment, SignatureRequest.AccountDefault,
            new SignatureSendingAccount("ghost@example.com", null), false, "Outlook");

        Assert.Null(none.Name);
        Assert.Contains(Jan + " has no default signature for new messages", none.Reason, StringComparison.Ordinal);
        Assert.Null(gone.Name);
        Assert.Contains("no signature of that name is in", gone.Reason, StringComparison.Ordinal);
    }

    [Fact]
    public void LeftOut_ADelegateOrSharedMailbox_GetsNone_AndTheResultSaysWhy()
    {
        SignatureEnvironment environment = Environment(Accounts().Add("Outlook", 1, FakeProfileAccountRegistry.Pop3(Jan, newSignature: "Home")));

        SignatureDecision decision = SignatureSelector.Decide(environment, SignatureRequest.AccountDefault, null, true, "Outlook");

        Assert.Null(decision.Name);
        Assert.Contains("delegate or shared mailbox", decision.Reason, StringComparison.Ordinal);
        Assert.Contains("Q160", decision.Reason, StringComparison.Ordinal);
        Assert.Null(decision.Refusal);
    }

    // ------------------------------------------------------------------ named

    [Fact]
    public void Named_GoesIn_AndTheAccountsOwnDefaultIsSuggested()
    {
        SignatureEnvironment environment = Environment(Accounts().Add("Outlook", 1, FakeProfileAccountRegistry.Pop3(Jan, newSignature: "Office")));

        SignatureDecision decision = Decide(environment, SignatureRequest.Named("nederlands"), false, running: "Outlook");

        Assert.Equal("Nederlands", decision.Name);
        Assert.Equal(SignatureDecision.SourceRequested, decision.Source);
        Assert.Equal("Office", decision.Suggestion);
        Assert.Contains("'Office' is the default of " + Jan, decision.SuggestionReason, StringComparison.Ordinal);
        Assert.False(decision.IsDefault);
    }

    [Fact]
    public void Named_TheDefaultItself_SuggestsNothing()
    {
        SignatureEnvironment environment = Environment(Accounts().Add("Outlook", 1, FakeProfileAccountRegistry.Pop3(Jan, newSignature: "Office")));

        Assert.Null(Decide(environment, SignatureRequest.Named("Office"), false, running: "Outlook").Suggestion);
    }

    [Fact]
    public void Named_Unknown_IsRefused_ListingWhatTheAccountCanUse()
    {
        SignatureEnvironment environment = Environment(Accounts().Add("Outlook", 1, FakeProfileAccountRegistry.Pop3(Jan)));

        SignatureDecision decision = Decide(environment, SignatureRequest.Named("Nope"), false, running: "Outlook");

        Assert.NotNull(decision.Refusal);
        Assert.Contains("'Nope' was not found", decision.Refusal, StringComparison.Ordinal);
        Assert.Contains("Office short", decision.Refusal, StringComparison.Ordinal);
        Assert.Contains("\"none\"", decision.Refusal, StringComparison.Ordinal);
    }

    [Fact]
    public void Named_AnotherAccountsRoamingSignature_IsRefused_WhileRoamingIsPossible_AndAllowedWhenItIsOff()
    {
        SignatureEnvironment roaming = Environment(Accounts().Add("Outlook", 1, FakeProfileAccountRegistry.Exchange(Jan)), new FakeSignatureSettings().Microsoft365());
        SignatureEnvironment off = Environment(Accounts().Add("Outlook", 1, FakeProfileAccountRegistry.Exchange(Jan)));

        SignatureDecision other = Decide(roaming, SignatureRequest.Named("Roam (other@example.com)"), false, running: "Outlook");
        SignatureDecision own = Decide(roaming, SignatureRequest.Named("Roam (jan@example.com)"), false, running: "Outlook");
        SignatureDecision device = Decide(roaming, SignatureRequest.Named("Home"), false, running: "Outlook");

        Assert.NotNull(other.Refusal);
        Assert.Contains("roaming signature of other@example.com", other.Refusal, StringComparison.Ordinal);
        Assert.DoesNotContain("Roam (other@example.com)", other.Refusal!.Substring(other.Refusal.IndexOf("can use", StringComparison.Ordinal)), StringComparison.Ordinal);
        Assert.Null(own.Refusal);
        Assert.Null(device.Refusal);
        Assert.Null(Decide(off, SignatureRequest.Named("Roam (other@example.com)"), false, running: "Outlook").Refusal);
    }

    [Fact]
    public void Named_ForADelegateMailbox_ADeviceSignatureGoesIn()
    {
        SignatureEnvironment environment = Environment(Accounts().Add("Outlook", 1, FakeProfileAccountRegistry.Pop3(Jan)));

        SignatureDecision decision = SignatureSelector.Decide(environment, SignatureRequest.Named("Corporate"), null, true, "Outlook");

        Assert.Equal("Corporate", decision.Name);
        Assert.Null(decision.Refusal);
    }

    // ------------------------------------------------------------------ none

    [Fact]
    public void None_NoSignature_RemovesAnExistingOne_AndSuggestsTheDefault()
    {
        SignatureEnvironment environment = Environment(Accounts().Add("Outlook", 1, FakeProfileAccountRegistry.Pop3(Jan, replySignature: "Office short")));

        SignatureDecision decision = Decide(environment, SignatureRequest.NoSignature, true, running: "Outlook");

        Assert.Null(decision.Name);
        Assert.True(decision.RemoveExisting);
        Assert.Equal(SignatureDecision.SourceNone, decision.Source);
        Assert.Equal("Office short", decision.Suggestion);
    }

    // ------------------------------------------------------------------ policy (Q156)

    [Fact]
    public void Policy_DisableSignatures_NoneByDefault_NamedRefused_NoneAllowed()
    {
        SignatureEnvironment environment = Environment(
            Accounts().Add("Outlook", 1, FakeProfileAccountRegistry.Pop3(Jan, newSignature: "Office")),
            new FakeSignatureSettings().Policy(SignaturePolicySource.GroupPolicy, disable: true));

        SignatureDecision byDefault = Decide(environment, SignatureRequest.AccountDefault, false, "Outlook");
        SignatureDecision named = Decide(environment, SignatureRequest.Named("Office"), false, "Outlook");
        SignatureDecision none = Decide(environment, SignatureRequest.NoSignature, false, "Outlook");

        Assert.Null(byDefault.Name);
        Assert.Contains("disabled on this computer by Group Policy", byDefault.Reason, StringComparison.Ordinal);
        Assert.NotNull(named.Refusal);
        Assert.Null(none.Refusal);
        Assert.Empty(SignatureSelector.Candidates(environment, new SignatureSendingAccount(Jan, null), "Outlook").Usable);
    }

    [Theory]
    [InlineData(SignaturePolicySource.CloudPolicy)]
    [InlineData(SignaturePolicySource.GroupPolicy)]
    [InlineData(SignaturePolicySource.Preference)]
    public void Policy_AnAdministratorsDefault_NeverDecidesTheDraftsDefault_TheAccountsOwnDoes(SignaturePolicySource source)
    {
        // Q200 (a), Q201, measured 2026-11-18 on Office LTSC 2024 with the values set before Outlook started: Outlook put
        // the account's own default in a new mail, a reply and a forward, and nothing in a reply or a forward of an
        // account with no reply default of its own - under a policy and under a deployed setting alike.
        SignatureEnvironment environment = Environment(
            Accounts().Add("Outlook", 1, FakeProfileAccountRegistry.Pop3(Jan, newSignature: "Office")),
            new FakeSignatureSettings().Policy(source, newMessage: "Corporate", reply: "Office short"));

        SignatureDecision own = Decide(environment, SignatureRequest.AccountDefault, false, "Outlook");
        Assert.Equal("Office", own.Name);
        Assert.Equal(SignatureDecision.SourceAccountDefault, own.Source);

        SignatureDecision none = Decide(environment, SignatureRequest.AccountDefault, true, "Outlook");
        Assert.Null(none.Name);
        Assert.Equal(SignatureDecision.SourceNone, none.Source);
        Assert.Contains(Jan + " has no default signature for replies and forwards", none.Reason, StringComparison.Ordinal);
    }

    [Fact]
    public void Policy_AForcedDefault_LeavesOneMessageAnotherSignature_OrNone()
    {
        // Q200 (a), measured 2026-11-18: under the policy Outlook's compose window kept its Signature button open in a
        // new mail, a reply and a forward - one message may have another signature, or none.
        SignatureEnvironment environment = Environment(
            Accounts().Add("Outlook", 1, FakeProfileAccountRegistry.Pop3(Jan, newSignature: "Home")),
            new FakeSignatureSettings().Policy(SignaturePolicySource.CloudPolicy, newMessage: "Corporate", reply: "Corporate"));

        Assert.Equal("Home", Decide(environment, SignatureRequest.AccountDefault, false, "Outlook").Name);
        foreach (bool? kind in new bool?[] { false, true, null })
        {
            SignatureDecision other = Decide(environment, SignatureRequest.Named("Office"), kind, "Outlook");
            Assert.Null(other.Refusal);
            Assert.Equal("Office", other.Name);
            Assert.Null(Decide(environment, SignatureRequest.NoSignature, kind, "Outlook").Refusal);
            Assert.Equal("Corporate", Decide(environment, SignatureRequest.Named("corporate"), kind, "Outlook").Name);
        }

        Assert.Equal("Home", Decide(environment, SignatureRequest.Named("Office"), false, "Outlook").Suggestion);
        Assert.True(SignatureSelector.Candidates(environment, new SignatureSendingAccount(Jan, null), "Outlook").DefaultsForced);
    }

    // ------------------------------------------------------------------ the add-in's future pick (Q158/Q162 (b))

    [Fact]
    public void Candidates_WhatAnAccountMayChooseFrom()
    {
        SignatureEnvironment environment = Environment(
            Accounts().Add("Outlook", 1, FakeProfileAccountRegistry.Exchange(Jan, newSignature: "Office", replySignature: "Office short")),
            new FakeSignatureSettings().Microsoft365());

        SignatureCandidates candidates = SignatureSelector.Candidates(environment, new SignatureSendingAccount(Jan, null), "Outlook");

        Assert.Equal("Outlook", candidates.Profile);
        Assert.Equal("Office", candidates.NewMessageDefault);
        Assert.Equal("Office short", candidates.ReplyForwardDefault);
        Assert.Contains(candidates.Usable, s => s.Name == "Roam (jan@example.com)");
        Assert.DoesNotContain(candidates.Usable, s => s.Name == "Roam (other@example.com)");
        Assert.False(candidates.SignaturesDisabled);
        Assert.False(candidates.DefaultsForced);
    }

    [Fact]
    public void Validate_BeforeComWork_ThrowsTheRefusal_AndHandsBackTheSignature()
    {
        SignatureEnvironment environment = Environment(Accounts().Add("Outlook", 1, FakeProfileAccountRegistry.Pop3(Jan)));

        Assert.Null(SignatureSelector.Validate(environment, SignatureRequest.AccountDefault, Jan, false));
        Assert.Null(SignatureSelector.Validate(environment, SignatureRequest.NoSignature, Jan, false));
        Assert.Equal("Office", SignatureSelector.Validate(environment, SignatureRequest.Named("office"), Jan, false)!.Name);
        SignatureRefusedException refusal = Assert.Throws<SignatureRefusedException>(
            () => SignatureSelector.Validate(environment, SignatureRequest.Named("Nope"), Jan, false));
        Assert.Equal("signature_not_usable", refusal.Reason);
    }

    // ------------------------------------------------------------------ helpers

    private static FakeProfileAccountRegistry Accounts() => new FakeProfileAccountRegistry();

    private SignatureEnvironment Environment(FakeProfileAccountRegistry accounts, FakeSignatureSettings? settings = null)
    {
        return SignatureEnvironment.Read(settings ?? new FakeSignatureSettings(), accounts, "16.0", _folder.Root, _folder.Directory);
    }

    private static SignatureDecision Decide(SignatureEnvironment environment, SignatureRequest request, bool? reply, string? running)
    {
        return SignatureSelector.Decide(environment, request, new SignatureSendingAccount(Jan, null), reply, running);
    }
}
