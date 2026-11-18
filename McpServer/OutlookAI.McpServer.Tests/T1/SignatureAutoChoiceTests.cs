using OutlookAI.Core.AddIn;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The add-in's own signature choice (Q158, Q162 (b)): when its AI may choose a draft's signature,
/// among which, and how its answer is matched. The signatures it chooses among are
/// <see cref="SignatureSelector.Candidates"/>'s, pinned by <c>SignatureSelectionTests</c>; what is
/// pinned here is what only the add-in decides.
/// </summary>
public sealed class SignatureAutoChoiceTests
{
    private static SignatureInfo Signature(string name) =>
        new SignatureInfo(name, name + " excerpt", @"C:\sig\" + name + ".htm", null, null, null, 0);

    private static SignatureCandidates Candidates(
        int usable = 2, bool disabled = false, bool forced = false)
    {
        List<SignatureInfo> list = new List<SignatureInfo>();
        for (int i = 0; i < usable; i++)
        {
            list.Add(Signature("Sig" + i));
        }

        return new SignatureCandidates("Outlook", list, "Sig0", "Sig0", disabled, forced);
    }

    // ===== The tick box =====

    [Theory]
    [InlineData(null, true)]   // never set: on, as the maintainer asked
    [InlineData(1, true)]
    [InlineData(2, true)]
    [InlineData(0, false)]
    public void TheTickBox_IsOnUnlessUnticked(int? stored, bool expected)
    {
        Assert.Equal(expected, SignatureAutoChoice.IsOn(stored));
    }

    // ===== A person's signature =====

    [Theory]
    [InlineData("Kind regards\r\nJane", null, false)]                    // nothing put yet that the add-in saw
    [InlineData("Kind regards\r\nJane", "Kind regards\r\nJane", false)]
    [InlineData("Kind regards\r\nJane\r\n", "Kind regards\r\nJane", false)] // Word's trailing marks
    [InlineData("Vriendelijke groeten\r\nJane", "Kind regards\r\nJane", true)]
    [InlineData("", "Kind regards\r\nJane", true)]                        // a person deleted it
    public void ASignatureChangedSinceItWasPut_IsAPersons(string now, string? put, bool expected)
    {
        Assert.Equal(expected, SignatureAutoChoice.ChangedByPerson(now, put));
    }

    // ===== Whether to ask the AI =====

    [Fact]
    public void WithEverythingInPlace_TheAIChoosesAmongTheAccountsSignatures()
    {
        SignatureCandidates candidates = Candidates(usable: 3);

        SignatureAutoChoicePlan plan = SignatureAutoChoice.Plan(true, candidates, draftHasSignature: true, changedByPerson: false);

        Assert.Equal(SignatureAutoChoiceOutcome.Choose, plan.Outcome);
        Assert.Same(candidates.Usable, plan.Candidates);
    }

    [Fact]
    public void TheTickBoxOff_LeavesItAlone() =>
        AssertLeft(SignatureAutoChoiceOutcome.SettingOff, SignatureAutoChoice.Plan(false, Candidates(), true, false));

    [Fact]
    public void PolicyDisablingSignatures_LeavesItAlone() =>
        AssertLeft(SignatureAutoChoiceOutcome.SignaturesDisabled, SignatureAutoChoice.Plan(true, Candidates(disabled: true), true, false));

    [Fact]
    public void PolicyForcingADefault_LeavesItAlone() =>
        AssertLeft(SignatureAutoChoiceOutcome.DefaultsForced, SignatureAutoChoice.Plan(true, Candidates(forced: true), true, false));

    [Fact]
    public void APersonsSignature_IsNeverReplaced() =>
        AssertLeft(SignatureAutoChoiceOutcome.ChosenByPerson, SignatureAutoChoice.Plan(true, Candidates(), true, changedByPerson: true));

    [Fact]
    public void ADraftWithoutASignature_GetsNoneAdded() =>
        AssertLeft(SignatureAutoChoiceOutcome.NoSignatureInDraft, SignatureAutoChoice.Plan(true, Candidates(), draftHasSignature: false, false));

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    public void OneSignatureOrNone_IsNothingToChooseBetween(int usable) =>
        AssertLeft(SignatureAutoChoiceOutcome.NothingToChooseBetween, SignatureAutoChoice.Plan(true, Candidates(usable), true, false));

    [Fact]
    public void ThePolicyRules_ComeBeforeAPersonsChoice()
    {
        // Disabled by policy: nothing may be inserted, whoever chose what.
        AssertLeft(SignatureAutoChoiceOutcome.SignaturesDisabled,
            SignatureAutoChoice.Plan(true, Candidates(disabled: true, forced: true), false, true));
    }

    private static void AssertLeft(SignatureAutoChoiceOutcome expected, SignatureAutoChoicePlan plan)
    {
        Assert.Equal(expected, plan.Outcome);
        Assert.Empty(plan.Candidates);
    }

    // ===== The AI's answer =====

    [Theory]
    [InlineData("Vriendelijke groeten", "Vriendelijke groeten")]
    [InlineData("  vriendelijke GROETEN \n", "Vriendelijke groeten")]
    [InlineData("\"Kind regards\"", "Kind regards")]
    [InlineData("'Kind regards'.", "Kind regards")]
    public void TheAnswer_IsMatchedToANameOnTheList(string answer, string expected)
    {
        IReadOnlyList<SignatureInfo> signatures = new[] { Signature("Kind regards"), Signature("Vriendelijke groeten") };

        Assert.Equal(expected, SignatureAutoChoice.FindNamed(signatures, answer)?.Name);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("Best wishes")]
    [InlineData("Kind regards, Jane")]
    public void AnAnswerNamingNothingOnTheList_ChoosesNothing(string? answer)
    {
        IReadOnlyList<SignatureInfo> signatures = new[] { Signature("Kind regards"), Signature("Vriendelijke groeten") };

        Assert.Null(SignatureAutoChoice.FindNamed(signatures, answer));
    }
}
