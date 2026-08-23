using OutlookAI.Core.Com;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins the freshness sweep's sort key, which shipped WRONG from the beginning and could
/// not be caught by any test that existed (measured defect, 2026-08-23).
/// <para>
/// The sweep ordered each folder's table newest-first with
/// <c>Table.Sort("urn:schemas:httpmail:datereceived", true)</c>. Microsoft's
/// <c>Table.Sort</c> reference says a sort property may be referenced "by their explicit
/// string names only; cannot reference properties by their namespaces", and a read-only
/// probe over five stores of a real profile found exactly that: the explicit name applied
/// on 5 of 5 and the namespace form was refused on 5 of 5. So the sweep never sorted, on
/// any store, for any user - and its 200-item-per-folder cap therefore cut an ARBITRARY
/// slice out of the one search tier whose entire purpose is recent mail.
/// </para>
/// <para>
/// The failure was invisible for two reasons and both are addressed here. The refusal was
/// swallowed by a <c>catch</c>, and the decision lived inside a COM call no CI test can
/// execute. <see cref="OutlookComSession.SweepSortProperties"/> exists so the rule that
/// matters is a pure function: <b>the first spelling tried is an explicit built-in property
/// name, never a namespace reference.</b> That is one assertion away from CI, where before
/// it needed a live mailbox and a probe.
/// </para>
/// </summary>
public sealed class SweepSortPropertyTests
{
    /// <summary>
    /// Every folder kind the sweep can be handed, INCLUDING the null a folder-scoped
    /// subtree walk passes and a kind nobody has defined - a ladder is required in all of
    /// them, because a folder with no ladder is a folder whose cap cuts arbitrarily.
    /// </summary>
    public static TheoryData<string?> AllFolderKinds()
    {
        TheoryData<string?> data = new TheoryData<string?> { null, "unknown-kind" };
        foreach (string kind in OutlookComSession.DefaultSweepFolderKinds)
        {
            data.Add(kind);
        }

        return data;
    }

    /// <summary>
    /// A namespace reference is any spelling with a scheme in it. Both forms the product
    /// uses carry one (<c>urn:schemas:...</c> and <c>http://schemas.microsoft.com/...</c>)
    /// and no explicit built-in name ever does, so the colon is the whole test.
    /// </summary>
    private static bool IsNamespaceReference(string property) => property.Contains(':');

    [Theory]
    [MemberData(nameof(AllFolderKinds))]
    public void EveryFolderKind_GetsALadder(string? folderKind)
    {
        IReadOnlyList<string> ladder = OutlookComSession.SweepSortProperties(folderKind);

        Assert.NotEmpty(ladder);
        Assert.All(ladder, p => Assert.False(string.IsNullOrWhiteSpace(p), "a blank sort spelling is not a spelling"));
        Assert.Equal(ladder.Count, ladder.Distinct(StringComparer.OrdinalIgnoreCase).Count());
    }

    /// <summary>
    /// THE regression. Whatever else changes about the ladders, the spelling Outlook is
    /// asked for FIRST must be one Outlook accepts - and the one it documents as unacceptable
    /// is precisely the one that shipped.
    /// </summary>
    [Theory]
    [MemberData(nameof(AllFolderKinds))]
    public void TheFirstSpellingTried_IsNeverANamespaceReference(string? folderKind)
    {
        IReadOnlyList<string> ladder = OutlookComSession.SweepSortProperties(folderKind);

        Assert.False(
            IsNamespaceReference(ladder[0]),
            "Table.Sort accepts explicit property names only, so a namespace reference in first position means the "
            + "sweep's sort is refused on every folder of kind '" + (folderKind ?? "(null)") + "' - which is the "
            + "defect this test exists for. Got: " + ladder[0]);
    }

    /// <summary>
    /// The namespace forms are kept as last-resort rungs rather than deleted, because the
    /// refusal is measured on one profile and one Outlook build. What must hold is the
    /// ORDER: every explicit name is offered before any namespace reference, so a healthy
    /// profile never reaches the rungs that are expected to fail.
    /// </summary>
    [Theory]
    [MemberData(nameof(AllFolderKinds))]
    public void ExplicitNames_AllComeBeforeAnyNamespaceReference(string? folderKind)
    {
        IReadOnlyList<string> ladder = OutlookComSession.SweepSortProperties(folderKind);

        bool seenNamespace = false;
        foreach (string property in ladder)
        {
            if (IsNamespaceReference(property))
            {
                seenNamespace = true;
                continue;
            }

            Assert.False(
                seenNamespace,
                "explicit name '" + property + "' is offered AFTER a namespace reference, so a profile that "
                + "accepts only explicit names would pay a refused COM call first");
        }
    }

    /// <summary>
    /// A sent folder's natural key is the SUBMIT time, not the delivery time. The sweep's
    /// own restriction already says so - it selects on datereceived OR date - and sorting
    /// only by received time undoes half of that: a sent copy admitted by the submit-time
    /// clause with no delivery time sorts as the oldest row and is the first thing the cap
    /// drops.
    /// </summary>
    [Fact]
    public void ASentFolder_SortsBySubmitTimeFirst()
    {
        Assert.Equal("SentOn", OutlookComSession.SweepSortProperties("sent")[0]);
    }

    /// <summary>
    /// And every other folder the sweep visits is an ARRIVAL folder, where the delivery
    /// time is the natural key.
    /// </summary>
    [Theory]
    [InlineData("inbox")]
    [InlineData("deleted")]
    [InlineData("junk")]
    [InlineData(null)]
    public void AnArrivalFolder_SortsByDeliveryTimeFirst(string? folderKind)
    {
        Assert.Equal("ReceivedTime", OutlookComSession.SweepSortProperties(folderKind)[0]);
    }

    /// <summary>
    /// The sweep passes its own lower-case kinds, but the match is ordinal-ignore-case so a
    /// caller that spells it differently gets the sent ladder rather than a silently wrong
    /// one. Nothing about a folder's sort key depends on how the word was typed.
    /// </summary>
    [Theory]
    [InlineData("Sent")]
    [InlineData("SENT")]
    public void TheSentMatch_IgnoresCase(string folderKind)
    {
        Assert.Equal(
            OutlookComSession.SweepSortProperties("sent"),
            OutlookComSession.SweepSortProperties(folderKind));
    }

    /// <summary>
    /// The sent ladder falls back to the arrival key rather than to nothing: most sent
    /// copies do carry a delivery time, and <c>ReceivedTime</c> is the spelling measured
    /// working on a real profile, so a store that will not carry <c>SentOn</c> still gets a
    /// sorted table instead of an arbitrary cut.
    /// </summary>
    [Fact]
    public void TheSentLadder_FallsBackToTheArrivalKey()
    {
        IReadOnlyList<string> sent = OutlookComSession.SweepSortProperties("sent");

        Assert.Contains("ReceivedTime", sent);
        Assert.True(
            sent.ToList().IndexOf("SentOn") < sent.ToList().IndexOf("ReceivedTime"),
            "the submit time must be preferred over the delivery time in a sent folder");
    }
}
